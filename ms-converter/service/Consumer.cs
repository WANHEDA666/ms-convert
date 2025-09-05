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
    public string Exchange { get; init; } = "";
    public string Queue { get; init; } = "";
    public int Prefetch { get; init; } = 1;
    public string ConsumerTag { get; init; } = "";
}

public class Consumer(ILogger<Consumer> logger, IOptions<RabbitOptions> opt, Storage storage, Converter converter, S3Uploader s3Uploader) : BackgroundService
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
        var factory = new ConnectionFactory { Uri = new Uri(_opt.AmqpUrl) };
        _conn = await factory.CreateConnectionAsync(ct);
        _ch   = await _conn.CreateChannelAsync(cancellationToken: ct);

        await _ch.ExchangeDeclareAsync(_opt.Exchange, ExchangeType.Topic, durable: true, cancellationToken: ct);
        await _ch.QueueDeclareAsync(_opt.Queue, durable: true, exclusive: false, autoDelete: false, arguments: null, cancellationToken: ct);
        await _ch.QueueBindAsync(_opt.Queue, _opt.Exchange, routingKey: "#", cancellationToken: ct);

        var prefetch = (ushort)Math.Max(1, _opt.Prefetch);
        await _ch.BasicQosAsync(0, prefetch, global: false, cancellationToken: ct);
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
            ConvertToPdf(GetTempFullPath(saveName));
            if (File.Exists(expectedPdfPath))
            {
                logger.LogInformation("PDF saved: {pdf}", expectedPdfPath);
                await UploadPdfToS3Async(uuid, fileName, expectedPdfPath, ct);
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
        var decoded = WebUtility.UrlDecode(urlEncodedFileName);
        var source = (Uri.TryCreate(decoded, UriKind.Absolute, out var abs) && (abs.Scheme == Uri.UriSchemeHttp || abs.Scheme == Uri.UriSchemeHttps))
            ? abs.ToString() : (uuid + "/" + urlEncodedFileName + "." + extension).Replace("//", "/");
        var saveName = $"{uuid}/file.{extension}".Replace("//", "/");
        return (source, saveName);
    }

    private async Task UploadPdfToS3Async(string uuid, string fileName, string pdfFull, CancellationToken ct)
    {
        var origRaw  = WebUtility.UrlDecode(fileName);
        var baseName = Path.GetFileNameWithoutExtension(
            Uri.TryCreate(origRaw, UriKind.Absolute, out var absUri) ? absUri.LocalPath : origRaw
        );
        var s3Key = $"{uuid}/{baseName}.pdf";
        await s3Uploader.UploadPdfAsync(pdfFull, s3Key, ct);
    }

    private static string GetTempFullPath(string saveName) => Path.Combine(AppContext.BaseDirectory, "temp", saveName.Replace('/', Path.DirectorySeparatorChar));

    private void ConvertToPdf(string srcFull) => converter.ConvertToPdf(srcFull);

    private async Task AckAsync(ulong deliveryTag, CancellationToken ct) => 
        await _ch!.BasicAckAsync(deliveryTag, multiple: false, cancellationToken: ct);

    private async Task NackAsync(ulong deliveryTag, bool requeue, CancellationToken ct) =>
        await _ch!.BasicNackAsync(deliveryTag, multiple: false, requeue: requeue, cancellationToken: ct);
}
