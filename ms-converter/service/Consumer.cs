using System.Net;
using System.Text;
using Microsoft.Extensions.Options;
using Newtonsoft.Json.Linq;
using RabbitMQ.Client;
using RabbitMQ.Client.Events;

namespace ms_converter.service;

public sealed class RabbitOptions
{
    public string AmqpUrl { get; init; } = "";
    public string VHost   { get; init; } = "/";
    public string Queue { get; init; } = "";
    public string ConsumerTag { get; init; } = "";
    public string? OutcomeQueue { get; init; }
}

public class Consumer(ILogger<Consumer> logger, IOptions<RabbitOptions> opt, Storage storage, Converter converter, S3Uploader s3Uploader, StatusPublisher statusPublisher) : BackgroundService
{
    private readonly RabbitOptions _opt = opt.Value;
    private IConnection? _conn;
    private IChannel? _ch;
    private string? _consumerTag;
    private readonly SemaphoreSlim _oneAtATime = new(1, 1);

    protected override async Task ExecuteAsync(CancellationToken stoppingToken)
    {
        await InitializeAsync(stoppingToken);

        var consumer = CreateConsumer(stoppingToken);
        _consumerTag = await _ch!.BasicConsumeAsync(
            queue: _opt.Queue,
            autoAck: false,
            consumerTag: _opt.ConsumerTag,
            noLocal: false,
            exclusive: false,
            arguments: null,
            consumer: consumer,
            cancellationToken: stoppingToken
        );

        var tcs = new TaskCompletionSource();
        using var reg = stoppingToken.Register(() => tcs.TrySetResult());
        await tcs.Task;

        await DisposeChannelAndConnectionAsync();
    }

    private async Task InitializeAsync(CancellationToken ct)
    {
        var factory = new ConnectionFactory { Uri = new Uri(_opt.AmqpUrl), VirtualHost = _opt.VHost};
        _conn = await factory.CreateConnectionAsync(ct);
        _ch   = await _conn.CreateChannelAsync(cancellationToken: ct);
        
        await _ch.QueueDeclareAsync(_opt.Queue, durable: true, exclusive: false, autoDelete: false, arguments: null, cancellationToken: ct);
        await _ch.BasicQosAsync(0, 1, global: false, cancellationToken: ct);
    }

    private async Task DisposeChannelAndConnectionAsync()
    {
        try { if (_consumerTag is not null) await _ch!.BasicCancelAsync(_consumerTag); } catch { }
        try { if (_ch is not null) { await _ch.CloseAsync(); await _ch.DisposeAsync(); } } catch { }
        try { if (_conn is not null) { await _conn.CloseAsync(); await _conn.DisposeAsync(); } } catch { }
    }

    private AsyncEventingBasicConsumer CreateConsumer(CancellationToken stoppingToken)
    {
        var consumer = new AsyncEventingBasicConsumer(_ch);
        consumer.ReceivedAsync += async (_, ea) =>
        {
            await _oneAtATime.WaitAsync(stoppingToken);
            try
            {
                await OnReceivedAsync(ea, stoppingToken);
            }
            finally
            {
                _oneAtATime.Release();
            }
        };
        return consumer;
    }

    private async Task OnReceivedAsync(BasicDeliverEventArgs ea, CancellationToken ct)
    {
        string? uuidForCleanup = null;
        var success = false;
        try
        {
            var content = Encoding.UTF8.GetString(ea.Body.ToArray());
            logger.LogInformation("recieve message with content: {content}", content);

            var (uuid, fileName, extension) = ParseMessage(content);
            uuidForCleanup = uuid;
            var (source, saveName) = BuildSourceAndSaveName(uuid, fileName, extension);
            logger.LogInformation("source={source} | saveName={saveName}", source, saveName);

            storage.DeleteUuidFolder(uuid);
            storage.CreateUuidFolder(uuid);
            await storage.DownloadSourceDocumentAsync(source, saveName, ct);

            var expectedPdfPath = storage.GetResultPdfPath();
            converter.ConvertToPdf(storage.GetTempFullPath(saveName));

            if (File.Exists(expectedPdfPath))
            {
                logger.LogInformation("PDF saved: {pdf}", expectedPdfPath);
                await UploadPdfToS3Async(uuid, fileName, extension, expectedPdfPath, ct);
                success = true;
            }
            else
            {
                logger.LogWarning("PDF not found at expected path: {pdf}", expectedPdfPath);
            }
        }
        catch (OperationCanceledException) when (ct.IsCancellationRequested) {}
        catch (Exception ex)
        {
            logger.LogError(ex, "handler error");
        }
        finally
        {
            try
            {
                if (!string.IsNullOrWhiteSpace(uuidForCleanup))
                    await statusPublisher.PublishAsync(uuidForCleanup, success, CancellationToken.None);
            }
            catch (Exception pubEx)
            {
                logger.LogWarning(pubEx, "failed to publish outcome for uuid={uuid}", uuidForCleanup);
            }
            try
            {
                if (!string.IsNullOrWhiteSpace(uuidForCleanup))
                    storage.DeleteUuidFolder(uuidForCleanup);
            }
            catch (Exception ex)
            {
                logger.LogWarning(ex, "cleanup temp failed for uuid {uuid}", uuidForCleanup);
            }
            storage.DeleteResultPdf();
        }
        if (success)
            await AckAsync(ea.DeliveryTag, ct);
        else
            await NackAsync(ea.DeliveryTag, requeue: false, ct);
    }


    private static (string uuid, string fileName, string extension) ParseMessage(string content)
    {
        var json = JObject.Parse(content);
        var uuid = json["uuid"]?.ToString() ?? throw new InvalidOperationException("uuid missing");
        var fileName = json["urlEncodedFileName"]?.ToString() ?? throw new InvalidOperationException("urlEncodedFileName missing");
        var extension = json["extension"]?.ToString() ?? throw new InvalidOperationException("extension missing");
        return (uuid, fileName, extension);
    }
    
    private static (string source, string saveName) BuildSourceAndSaveName(string uuid, string urlEncodedFileName, string extension)
    {
        if (Uri.TryCreate(urlEncodedFileName, UriKind.Absolute, out var abs) && (abs.Scheme == Uri.UriSchemeHttp || abs.Scheme == Uri.UriSchemeHttps))
        {
            var saveNameAbs = $"{uuid}/file.{extension}".Replace("//", "/");
            return (abs.ToString(), saveNameAbs);
        }
        var encodedName = urlEncodedFileName.Contains('%') ? urlEncodedFileName : Uri.EscapeDataString(urlEncodedFileName);
        var source = $"{uuid}/{encodedName}.{extension}".Replace("//", "/");
        var saveName = $"{uuid}/file.{extension}".Replace("//", "/");
        return (source, saveName);
    }
    
    private async Task UploadPdfToS3Async(string uuid, string fileName, string extension, string pdfFull, CancellationToken ct)
    {
        var origRaw = WebUtility.UrlDecode(fileName);
        string baseName;
        if (Uri.TryCreate(origRaw, UriKind.Absolute, out var absUri))
        {
            baseName = Path.GetFileNameWithoutExtension(absUri.LocalPath);
        }
        else
        {
            var suffix = "." + extension.TrimStart('.');
            baseName = origRaw.EndsWith(suffix, StringComparison.OrdinalIgnoreCase) ? origRaw[..^suffix.Length] : origRaw;
        }
        var s3Key = $"{uuid}/{baseName}.pdf";
        await s3Uploader.UploadPdfAsync(pdfFull, s3Key, ct);
    }

    private async Task AckAsync(ulong deliveryTag, CancellationToken ct) => 
        await _ch!.BasicAckAsync(deliveryTag, multiple: false, cancellationToken: ct);

    private async Task NackAsync(ulong deliveryTag, bool requeue, CancellationToken ct) =>
        await _ch!.BasicNackAsync(deliveryTag, multiple: false, requeue: requeue, cancellationToken: ct);
    
    private static bool IsIrrecoverableOfficeError(Exception ex)
    {
        if (ex is NetOffice.Exceptions.MethodCOMException)
            return true;
        if (ex is System.Runtime.InteropServices.COMException com)
            return com.HResult is unchecked((int)0x80004005) or unchecked((int)0x80020005);
        return false;
    }

    private static bool IsTransient(Exception ex)
    {
        if (ex is OperationCanceledException) 
            return false;
        if (ex is HttpRequestException hre)
        {
            if (hre.StatusCode is null) return true;
            var code = (int)hre.StatusCode.Value;
            return code is >= 500 or 408 or 429;
        }
        if (ex is Amazon.S3.AmazonS3Exception s3Ex)
        {
            var sc = (int)s3Ex.StatusCode;
            return sc is >= 500 or 408 or 429;
        }
        if (ex is IOException ioex)
            return ioex.HResult is unchecked((int)0x80070020) or unchecked((int)0x80070021);
        return false;
    }
}
