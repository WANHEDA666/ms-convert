using Microsoft.Extensions.Options;

namespace ms_converter.service;

public sealed class StorageOptions
{
    public string AwsS3BucketName { get; init; } = "";
    public string? BaseDownloadUrl { get; init; }
}

public class Storage(ILogger<Storage> logger, HttpClient http, IOptionsMonitor<StorageOptions> opt) : IDisposable
{
    private readonly IDisposable? _reloadReg = opt.OnChange(storageOptions =>
        logger.LogInformation("Storage options reloaded: Base={Base} | Bucket={Bucket}", storageOptions.BaseDownloadUrl, storageOptions.AwsS3BucketName));

    private string ExtractDirectory => Path.Combine(AppContext.BaseDirectory, "temp");
    private string ResultDirectory  => Path.Combine(AppContext.BaseDirectory, "result");
    
    public void DeleteUuidFolder(string uuid)
    {
        var dir = Path.Combine(ExtractDirectory, uuid);
        if (Directory.Exists(dir))
        {
            new DirectoryInfo(dir).Delete(true);
            logger.LogInformation("Deleted dir: {Dir}", dir);
        }
    }

    public void CreateUuidFolder(string uuid)
    {
        var dir = Path.Combine(ExtractDirectory, uuid);
        Directory.CreateDirectory(dir);
        logger.LogInformation("Created dir: {Dir}", dir);
    }
    
    public async Task DownloadSourceDocumentAsync(string downloadPathOrUrl, string savePath, CancellationToken ct = default)
    {
        var storageOptions = opt.CurrentValue;
        string url;
        if (Uri.TryCreate(downloadPathOrUrl, UriKind.Absolute, out var abs) && (abs.Scheme == Uri.UriSchemeHttp || abs.Scheme == Uri.UriSchemeHttps))
        {
            url = abs.ToString();
        }
        else
        {
            var baseUrl = !string.IsNullOrWhiteSpace(storageOptions.BaseDownloadUrl) ? storageOptions.BaseDownloadUrl!.TrimEnd('/') : $"https://{storageOptions.AwsS3BucketName.TrimEnd('/')}";
            url = $"{baseUrl}/{downloadPathOrUrl.TrimStart('/')}";
        }
        logger.LogInformation("Downloading from: {Url}", url);

        var destFull = Path.Combine(ExtractDirectory, savePath.Replace('/', Path.DirectorySeparatorChar));
        Directory.CreateDirectory(Path.GetDirectoryName(destFull)!);

        using var resp = await http.GetAsync(url, HttpCompletionOption.ResponseHeadersRead, ct);
        resp.EnsureSuccessStatusCode();

        var tmp = destFull + ".part";
        await using (var src = await resp.Content.ReadAsStreamAsync(ct))
        await using (var dst = File.Create(tmp))
        {
            await src.CopyToAsync(dst, ct);
        }
        if (File.Exists(destFull)) File.Delete(destFull);
        File.Move(tmp, destFull);
    }
    
    public string GetTempFullPath(string saveName) => Path.Combine(AppContext.BaseDirectory, "temp", saveName.Replace('/', Path.DirectorySeparatorChar));
    
    public string GetResultPdfPath()
    {
        Directory.CreateDirectory(ResultDirectory);
        return Path.Combine(ResultDirectory, "file.pdf");
    }

    public void DeleteResultPdf()
    {
        var pdf = Path.Combine(ResultDirectory, "file.pdf");
        if (File.Exists(pdf))
        {
            try
            {
                File.Delete(pdf);
                logger.LogInformation("Deleted temp pdf: {Path}", pdf);
            }
            catch (Exception ex)
            {
                logger.LogWarning(ex, "Failed to delete temp pdf: {Path}", pdf);
            }
        }
    }

    public void Dispose() => _reloadReg?.Dispose();
}
