using Microsoft.Extensions.Options;
namespace ms_converter.service;

public sealed class StorageOptions
{
    public string AwsS3BucketName { get; set; } = "";
    public string? BaseDownloadUrl { get; set; }
}

public class Storage(ILogger<Storage> logger, HttpClient http, IOptions<StorageOptions> opt)
{
    private readonly StorageOptions _opt = opt.Value;
    private string ExtractDirectory => Path.Combine(AppContext.BaseDirectory, "temp");

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
    
    public async Task DownloadSourceDocumentAsync(string downloadPath, string savePath, CancellationToken ct = default)
    {
        var baseUrl = !string.IsNullOrWhiteSpace(_opt.BaseDownloadUrl) ? _opt.BaseDownloadUrl!.TrimEnd('/') : $"https://{_opt.AwsS3BucketName.TrimEnd('/')}";
        var url = $"{baseUrl}/{downloadPath.TrimStart('/')}";
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

        logger.LogInformation("Saved to: {Path}", destFull);
    }
}
