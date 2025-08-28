using System.Text;
using Microsoft.Extensions.Options;
using Newtonsoft.Json.Linq;
using RabbitMQ.Client;
using RabbitMQ.Client.Events;

namespace ms_converter.service;

public sealed class RabbitOptions
{
    public string AmqpUrl { get; set; } = "";
    public string Exchange { get; set; } = "";
    public string Queue { get; set; } = "";
    public int Prefetch { get; set; } = 50;
    public string ConsumerTag { get; set; } = "csharp-worker";
}

public class Consumer(ILogger<Consumer> logger, IOptions<RabbitOptions> opt, Storage storage, Converter converter) : BackgroundService
{
    private readonly RabbitOptions _opt = opt.Value;
    private IConnection? _conn;
    private IChannel? _ch;
    private string? _consumerTag;

    protected override async Task ExecuteAsync(CancellationToken stoppingToken)
    {
        var factory = new ConnectionFactory
        {
            Uri = new Uri(_opt.AmqpUrl),
        };

        _conn = await factory.CreateConnectionAsync(stoppingToken);
        _ch   = await _conn.CreateChannelAsync(cancellationToken: stoppingToken);

        await _ch.ExchangeDeclareAsync(_opt.Exchange, ExchangeType.Topic, durable: true, cancellationToken: stoppingToken);
        await _ch.QueueDeclareAsync(_opt.Queue, durable: true, exclusive: false, autoDelete: false, arguments: null, cancellationToken: stoppingToken);
        await _ch.QueueBindAsync(_opt.Queue, _opt.Exchange, routingKey: "#", cancellationToken: stoppingToken);
        await _ch.BasicQosAsync(0, (ushort)_opt.Prefetch, global: false, cancellationToken: stoppingToken);

        var consumer = new AsyncEventingBasicConsumer(_ch);
        consumer.ReceivedAsync += async (_, ea) =>
        {
            try
            {
                var content = Encoding.UTF8.GetString(ea.Body.ToArray());
                var json    = JObject.Parse(content);
                var uuid      = json["uuid"]?.ToString() ?? throw new InvalidOperationException("uuid missing");
                var fileName  = json["urlEncodedFileName"]?.ToString() ?? throw new InvalidOperationException("urlEncodedFileName missing");
                var extension = json["extension"]?.ToString() ?? throw new InvalidOperationException("extension missing");
                logger.LogInformation("recieve message with content: {content}", content);

                var downloadPath = uuid + "/" + fileName + "." + extension;
                var saveName     = $"{uuid}/file.{extension}";
                logger.LogInformation("downloadPath={downloadPath} | saveName={saveName}", downloadPath, saveName);

                storage.DeleteUuidFolder(uuid);
                storage.CreateUuidFolder(uuid);
                await storage.DownloadSourceDocumentAsync(downloadPath, saveName, stoppingToken);
                
                var srcFull = Path.Combine(AppContext.BaseDirectory, "temp", saveName.Replace('/', Path.DirectorySeparatorChar));
                converter.ConvertToPdf(srcFull);
                logger.LogInformation("converted to pdf: {src}", srcFull);

                await _ch!.BasicAckAsync(ea.DeliveryTag, multiple: false, cancellationToken: stoppingToken);
            }
            catch (Exception ex)
            {
                logger.LogError(ex, "handler error");
                await _ch!.BasicNackAsync(ea.DeliveryTag, multiple: false, requeue: false, cancellationToken: stoppingToken);
            }
        };

        _consumerTag = await _ch.BasicConsumeAsync(
            queue: _opt.Queue,
            autoAck: false,
            consumerTag: _opt.ConsumerTag,
            noLocal: false,
            exclusive: false,
            arguments: null,
            consumer: consumer,
            cancellationToken: stoppingToken
        );

        logger.LogInformation("Consuming from '{Queue}'...", _opt.Queue);

        var tcs = new TaskCompletionSource();
        using var reg = stoppingToken.Register(() => tcs.TrySetResult());
        await tcs.Task;

        try { if (_consumerTag is not null) await _ch.BasicCancelAsync(_consumerTag); } catch { }
        try { if (_ch is not null) { await _ch.CloseAsync(); await _ch.DisposeAsync(); } } catch { }
        try { if (_conn is not null) { await _conn.CloseAsync(); await _conn.DisposeAsync(); } } catch { }
    }
}
