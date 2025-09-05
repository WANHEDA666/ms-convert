using Amazon.Runtime;
using Amazon.S3;
using Amazon.S3.Model;
using Microsoft.Extensions.Options;

namespace ms_converter.service;

public sealed class S3Options
{
    public string AwsKey { get; init; } = "";
    public string AwsSecret { get; init; } = "";
    public string AwsS3EndPoint { get; init; } = "";
    public string AwsS3BucketName { get; init; } = "";
}

public class S3Uploader(ILogger<S3Uploader> logger, IOptionsMonitor<S3Options> opt) : IDisposable
{
    private readonly IDisposable? _reg = opt.OnChange(s3Options => logger.LogInformation("S3 options reloaded: Endpoint={Endpoint}, Bucket={Bucket}", 
        s3Options.AwsS3EndPoint, s3Options.AwsS3BucketName));

    public async Task<string> UploadPdfAsync(string localPath, string key, CancellationToken ct = default)
    {
        var s3Options = opt.CurrentValue;
        using var s3 = new AmazonS3Client(
            new BasicAWSCredentials(s3Options.AwsKey, s3Options.AwsSecret),
            new AmazonS3Config { ServiceURL = s3Options.AwsS3EndPoint });
        
        logger.LogInformation("S3 put: endpoint={Endpoint} bucket={Bucket} key={Key}", s3Options.AwsS3EndPoint, s3Options.AwsS3BucketName, key);
        var resp = await s3.PutObjectAsync(new PutObjectRequest {
            BucketName  = s3Options.AwsS3BucketName,
            Key         = key,
            FilePath    = localPath,
            CannedACL   = S3CannedACL.PublicRead,
            ContentType = "application/pdf",
            Headers = {CacheControl = "max-age=31536000"}
        }, ct);
        logger.LogInformation("Uploaded to s3://{Bucket}/{Key} | ETag={ETag}", s3Options.AwsS3BucketName, key, resp.ETag);
        return $"{s3Options.AwsS3EndPoint.TrimEnd('/')}/{s3Options.AwsS3BucketName}/{Uri.EscapeDataString(key).Replace("%2F","/")}";
    }

    public void Dispose() => _reg?.Dispose();
}
