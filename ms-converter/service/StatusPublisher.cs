using System.Text;
using Microsoft.Extensions.Options;
using RabbitMQ.Client;

namespace ms_converter.service;

public sealed class StatusPublisher(ILogger<StatusPublisher> logger, IOptions<RabbitOptions> opt)
{
    private readonly RabbitOptions _opt = opt.Value;

    public async Task PublishAsync(string uuid, bool success, CancellationToken ct = default)
    {
        if (string.IsNullOrWhiteSpace(_opt.OutcomeQueue))
        {
            logger.LogDebug("OutcomeQueue is empty, skip publish");
            return;
        }

        var factory = new ConnectionFactory { Uri = new Uri(_opt.AmqpUrl), VirtualHost = _opt.VHost };
        await using var conn = await factory.CreateConnectionAsync(ct);
        await using var ch   = await conn.CreateChannelAsync(cancellationToken: ct);
        await ch.QueueDeclareAsync(queue: _opt.OutcomeQueue, durable: true, exclusive: false, autoDelete: false, arguments: null, cancellationToken: ct);

        var payload = $"{{\"uuid\":\"{uuid}\",\"success\":{(success ? "true" : "false")}}}";
        var body    = Encoding.UTF8.GetBytes(payload);
        var props = new BasicProperties
        {
            DeliveryMode = DeliveryModes.Persistent,
            ContentType  = "application/json"
        };
        await ch.BasicPublishAsync(exchange: "", routingKey: _opt.OutcomeQueue, mandatory: false, basicProperties: props, body: body, cancellationToken: ct);
        logger.LogInformation("Published outcome to '{Queue}': {Payload}", _opt.OutcomeQueue, payload);
    }
}