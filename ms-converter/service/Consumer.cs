using System.Net;
using System.Text;
using Microsoft.Extensions.Options;
using ms_converter.service.errors;
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
        bool success;
        bool forDelete = false;
        try
        {
            var content = Encoding.UTF8.GetString(ea.Body.ToArray());
            logger.LogInformation("recieve message with content: {content}", content);

            var (uuid, fileName, extension, output) = ParseMessage(content);
            uuidForCleanup = uuid;
            var (source, saveName) = BuildSourceAndSaveName(uuid, fileName, extension);
            logger.LogInformation("source={source} | saveName={saveName}", source, saveName);

            storage.DeleteUuidFolder(uuid);
            storage.CreateUuidFolder(uuid);
            await storage.DownloadSourceDocumentAsync(source, saveName, ct);

            var expectedPdfPath = storage.GetResultPdfPath();
            var expectedHtmlPath = storage.GetResultHtmlPath();
            
            if (output == "html")
            {
                converter.ConvertToHtml(storage.GetTempFullPath(saveName));
                
                if (!File.Exists(expectedHtmlPath))
                {
                    throw new LocalOfficeApiException($"HTML not found at expected path: {expectedHtmlPath}");
                }
                var ext = Path.GetExtension(storage.GetTempFullPath(saveName)).TrimStart('.').ToLowerInvariant();
                switch (ext)
                {
                    case "ppt":
                    case "pptx":
                    case "pps":
                    case "ppsx":
                    case "pptm":
                    case "pot":
                    case "odp":
                        forDelete =  true;
                        await UploadHtmlToS3Async(uuid, expectedHtmlPath, ct);
                        break;
                }

            } else if (output == "pdf")
            {
                converter.ConvertToPdf(storage.GetTempFullPath(saveName));
            
                if (!File.Exists(expectedPdfPath))
                {
                    throw new LocalOfficeApiException($"PDF not found at expected path: {expectedPdfPath}");
                }
                await UploadPdfToS3Async(uuid, fileName, extension, expectedPdfPath, ct);
                forDelete =  true;
            }
            success = true;
        }
        catch (OperationCanceledException) when (ct.IsCancellationRequested)
        {
            return;
        }
        catch (HttpRequestException hre) when (hre.StatusCode == HttpStatusCode.NotFound)
        {
            logger.LogError(hre, "404 error");
            try
            {
                if (!string.IsNullOrWhiteSpace(uuidForCleanup))
                    await statusPublisher.PublishAsync(uuidForCleanup, false, CancellationToken.None);
            }
            catch (Exception pubEx)
            {
                logger.LogWarning(pubEx, "failed to publish outcome for uuid={uuid}", uuidForCleanup);
            }
            await AckAsync(ea.DeliveryTag, ct);
            return;
        }
        catch (LocalOfficeApiException oex)
        {
            logger.LogError(oex, "office error");
            try
            {
                if (!string.IsNullOrWhiteSpace(uuidForCleanup))
                    await statusPublisher.PublishAsync(uuidForCleanup, false, CancellationToken.None);
            }
            catch (Exception pubEx)
            {
                logger.LogWarning(pubEx, "failed to publish outcome for uuid={uuid}", uuidForCleanup);
            }
            await AckAsync(ea.DeliveryTag, ct);
            return;
        }
        catch (Exception ex)
        {
            logger.LogError(ex, "handle error");
            await NackAsync(ea.DeliveryTag, requeue: true, ct);
            return;
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
            if (forDelete)
            {
                storage.DeleteResultPdf();
                storage.DeleteResultHtml();
            }
        }
        if (success)
        {
            if (!string.IsNullOrWhiteSpace(uuidForCleanup))
            {
                try
                {
                   await statusPublisher.PublishAsync(uuidForCleanup, true, CancellationToken.None);
                }
                catch (Exception pubEx)
                {
                    logger.LogWarning(pubEx, "failed to publish outcome for uuid={uuid}", uuidForCleanup);
                }
            }
            await AckAsync(ea.DeliveryTag, ct);
        }
    }

    private static (string uuid, string fileName, string extension, string output) ParseMessage(string content)
    {
        var json = JObject.Parse(content);
        var uuid = json["uuid"]?.ToString() ?? throw new InvalidOperationException("uuid missing");
        var fileName = json["urlEncodedFileName"]?.ToString() ?? throw new InvalidOperationException("urlEncodedFileName missing");
        var extension = json["extension"]?.ToString() ?? throw new InvalidOperationException("extension missing");
        var output = json["output"]?.ToString()?.ToLowerInvariant() ?? "pdf";
        return (uuid, fileName, extension, output);
    }
    
    private static (string source, string saveName) BuildSourceAndSaveName(string uuid, string urlEncodedFileName, string extension)
    {
        var decoded = WebUtility.UrlDecode(urlEncodedFileName);
        if (Uri.TryCreate(decoded, UriKind.Absolute, out var abs) && (abs.Scheme == Uri.UriSchemeHttp || abs.Scheme == Uri.UriSchemeHttps))
        {
            var path = abs.AbsolutePath;
            var hasFileName = !string.IsNullOrEmpty(path) && path != "/" && Path.HasExtension(path);
            if (hasFileName)
            {
                var saveNameAbs = $"{uuid}/file.{extension}".Replace("//", "/");
                return (abs.ToString(), saveNameAbs);
            }
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
        await s3Uploader.UploadFileAsync(pdfFull, s3Key, "application/pdf", ct);
    }
    
    private async Task UploadHtmlToS3Async(string uuid, string htmlFullPath, CancellationToken ct)
    {
        var htmlKey = $"{uuid}/presentation.html";
        await s3Uploader.UploadFileAsync(htmlFullPath, htmlKey, "text/html", ct);

        var imagesDir = Path.Combine(Path.GetDirectoryName(htmlFullPath) ?? ".", Path.GetFileNameWithoutExtension(htmlFullPath) + ".files");
        if (Directory.Exists(imagesDir))
        {
            foreach (var f in Directory.GetFiles(imagesDir))
            {
                var fileName = Path.GetFileName(f);
                var fileKey = $"{uuid}/{Path.GetFileNameWithoutExtension(htmlFullPath)}.files/{fileName}";
                await s3Uploader.UploadFileAsync(f, fileKey, "image/jpeg", ct);
            }
        }
    }

    private async Task AckAsync(ulong deliveryTag, CancellationToken ct) => 
        await _ch!.BasicAckAsync(deliveryTag, multiple: false, cancellationToken: ct);

    private async Task NackAsync(ulong deliveryTag, bool requeue, CancellationToken ct) =>
        await _ch!.BasicNackAsync(deliveryTag, multiple: false, requeue: requeue, cancellationToken: ct);
}