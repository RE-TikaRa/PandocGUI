using System.Diagnostics;
using System.IO.Compression;
using System.Text.Json;

namespace PandocGUI.Services;

public static class PandocService
{
    public static PandocDetectionResult DetectPandoc(string? preferredPath)
    {
        var steps = new List<string>();

        if (IsPandocExecutable(preferredPath))
        {
            steps.Add($"命中：已选择路径 -> {preferredPath}");
            return new PandocDetectionResult(preferredPath, "Preferred", steps);
        }

        steps.Add("未命中：已选择路径为空或不存在");

        var saved = AppSettings.PandocPath;
        if (IsPandocExecutable(saved))
        {
            steps.Add($"命中：已保存路径 -> {saved}");
            return new PandocDetectionResult(saved, "Saved", steps);
        }

        steps.Add("未命中：已保存路径为空或不存在");

        var fromWhere = LocatePandocFromWhere(steps);
        if (IsPandocExecutable(fromWhere))
        {
            steps.Add($"命中：where pandoc -> {fromWhere}");
            return new PandocDetectionResult(fromWhere, "Where", steps);
        }

        steps.Add("未命中：where pandoc");

        var fromPath = LocatePandocFromPathOnly(steps);
        if (IsPandocExecutable(fromPath))
        {
            steps.Add($"命中：PATH -> {fromPath}");
            return new PandocDetectionResult(fromPath, "PATH", steps);
        }

        steps.Add("未命中：PATH");
        return new PandocDetectionResult(null, "None", steps);
    }

    public static string? LocatePandoc(string? preferredPath)
    {
        if (IsPandocExecutable(preferredPath))
        {
            return preferredPath;
        }

        var embedded = AppSettings.PandocPath;
        if (IsPandocExecutable(embedded))
        {
            return embedded;
        }

        return LocatePandocFromPath();
    }

    public static string? LocatePandocFromPath()
    {
        var fromWhere = LocatePandocFromWhere(null);
        if (IsPandocExecutable(fromWhere))
        {
            return fromWhere;
        }

        return LocatePandocFromPathOnly(null);
    }

    public static bool IsPandocExecutable(string? path)
    {
        return !string.IsNullOrWhiteSpace(path) && File.Exists(path);
    }

    private static string? LocatePandocFromWhere(List<string>? steps)
    {
        try
        {
            steps?.Add("检测：where pandoc");
            var psi = new ProcessStartInfo
            {
                FileName = "where",
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                UseShellExecute = false,
                CreateNoWindow = true
            };
            psi.ArgumentList.Add("pandoc");

            using var process = new Process { StartInfo = psi };
            if (!process.Start())
            {
                steps?.Add("where 启动失败");
                return null;
            }

            if (!process.WaitForExit(2000))
            {
                steps?.Add("where 超时（2s）");
                try
                {
                    process.Kill(true);
                }
                catch
                {
                    // ignore
                }
                return null;
            }

            var output = process.StandardOutput.ReadToEnd();
            var error = process.StandardError.ReadToEnd();

            if (process.ExitCode != 0 || string.IsNullOrWhiteSpace(output))
            {
                steps?.Add($"where 失败：ExitCode={process.ExitCode}");
                if (!string.IsNullOrWhiteSpace(error))
                {
                    steps?.Add($"where 输出错误：{error.Trim()}");
                }
                return null;
            }

            var lines = output.Split(new[] { "\r\n", "\n" }, StringSplitOptions.RemoveEmptyEntries);
            steps?.Add($"where 返回 {lines.Length} 行");
            foreach (var line in lines)
            {
                var candidate = line.Trim().Trim('"');
                if (string.IsNullOrWhiteSpace(candidate))
                {
                    continue;
                }

                if (!Path.HasExtension(candidate))
                {
                    var withExe = candidate + ".exe";
                    if (File.Exists(withExe))
                    {
                        return withExe;
                    }
                }

                if (File.Exists(candidate))
                {
                    return candidate;
                }
            }

            return null;
        }
        catch
        {
            steps?.Add("where 异常");
            return null;
        }
    }

    private static string? LocatePandocFromPathOnly(List<string>? steps)
    {
        steps?.Add("检测：PATH");
        var pathEnv = Environment.GetEnvironmentVariable("PATH");
        if (string.IsNullOrWhiteSpace(pathEnv))
        {
            steps?.Add("PATH 为空");
            return null;
        }

        var parts = pathEnv.Split(Path.PathSeparator, StringSplitOptions.RemoveEmptyEntries);
        steps?.Add($"PATH 项数：{parts.Length}");

        var skipped = 0;
        foreach (var part in parts)
        {
            var trimmed = part.Trim().Trim('"');
            if (string.IsNullOrWhiteSpace(trimmed))
            {
                continue;
            }

            if (trimmed.StartsWith(@"\\", StringComparison.Ordinal))
            {
                skipped++;
                continue;
            }

            var candidate = Path.Combine(trimmed, "pandoc.exe");
            if (File.Exists(candidate))
            {
                return candidate;
            }

            candidate = Path.Combine(trimmed, "pandoc");
            if (File.Exists(candidate))
            {
                return candidate;
            }
        }

        if (skipped > 0)
        {
            steps?.Add($"跳过 UNC 路径：{skipped}");
        }

        return null;
    }

    public static async Task<PandocInfo?> GetInfoAsync(string pandocPath, CancellationToken cancellationToken)
    {
        if (!IsPandocExecutable(pandocPath))
        {
            return null;
        }

        using var timeoutCts = new CancellationTokenSource(TimeSpan.FromSeconds(5));
        using var linkedCts = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken, timeoutCts.Token);
        PandocRunResult result;
        try
        {
            result = await RunAsync(pandocPath, new[] { "--version" }, linkedCts.Token);
        }
        catch (OperationCanceledException)
        {
            return null;
        }
        if (!result.Succeeded || string.IsNullOrWhiteSpace(result.StandardOutput))
        {
            return null;
        }
        var firstLine = result.StandardOutput
            .Split(new[] { "\r\n", "\n" }, StringSplitOptions.RemoveEmptyEntries)
            .FirstOrDefault() ?? string.Empty;

        var version = firstLine.Replace("pandoc", string.Empty, StringComparison.OrdinalIgnoreCase)
            .Trim();

        return new PandocInfo(pandocPath, version, firstLine);
    }

    public static Task<IReadOnlyList<string>> ListInputFormatsAsync(string pandocPath, CancellationToken ct)
        => ListFormatsAsync(pandocPath, "--list-input-formats", ct);

    public static Task<IReadOnlyList<string>> ListOutputFormatsAsync(string pandocPath, CancellationToken ct)
        => ListFormatsAsync(pandocPath, "--list-output-formats", ct);

    private static async Task<IReadOnlyList<string>> ListFormatsAsync(string pandocPath, string argument, CancellationToken ct)
    {
        using var timeoutCts = new CancellationTokenSource(TimeSpan.FromSeconds(8));
        using var linkedCts = CancellationTokenSource.CreateLinkedTokenSource(ct, timeoutCts.Token);
        PandocRunResult result;
        try
        {
            result = await RunAsync(pandocPath, new[] { argument }, linkedCts.Token);
        }
        catch (OperationCanceledException)
        {
            return Array.Empty<string>();
        }
        if (!result.Succeeded)
        {
            return Array.Empty<string>();
        }
        var lines = result.StandardOutput
            .Split(new[] { "\r\n", "\n" }, StringSplitOptions.RemoveEmptyEntries)
            .Select(line => line.Trim())
            .Where(line => !string.IsNullOrWhiteSpace(line))
            .OrderBy(line => line, StringComparer.OrdinalIgnoreCase)
            .ToList();

        return lines;
    }

    public static async Task<PandocRunResult> RunAsync(string pandocPath, IReadOnlyList<string> arguments, CancellationToken ct)
    {
        var psi = new ProcessStartInfo
        {
            FileName = pandocPath,
            RedirectStandardOutput = true,
            RedirectStandardError = true,
            UseShellExecute = false,
            CreateNoWindow = true
        };

        foreach (var arg in arguments)
        {
            psi.ArgumentList.Add(arg);
        }
        using var process = new Process { StartInfo = psi };
        process.Start();

        try
        {
            var outputTask = process.StandardOutput.ReadToEndAsync(ct);
            var errorTask = process.StandardError.ReadToEndAsync(ct);

            await process.WaitForExitAsync(ct);

            var output = await outputTask;
            var error = await errorTask;

            return new PandocRunResult(process.ExitCode == 0, output, error);
        }
        catch (OperationCanceledException)
        {
            try
            {
                if (!process.HasExited)
                {
                    process.Kill(true);
                }
            }
            catch
            {
                // ignore
            }

            throw;
        }
    }

    public static async Task<PandocDownloadResult> DownloadLatestAsync(IProgress<double>? progress, CancellationToken ct)
    {
        var targetRoot = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "PandocGUI",
            "pandoc");

        Directory.CreateDirectory(targetRoot);

        var apiUrl = "https://api.github.com/repos/jgm/pandoc/releases/latest";
        using var http = new HttpClient();
        http.DefaultRequestHeaders.UserAgent.ParseAdd("PandocGUI");

        using var response = await http.GetAsync(apiUrl, ct);
        response.EnsureSuccessStatusCode();

        var json = await response.Content.ReadAsStringAsync(ct);
        var downloadUrl = GetWindowsZipUrl(json);
        if (string.IsNullOrWhiteSpace(downloadUrl))
        {
            return new PandocDownloadResult(false, null, "无法找到 Windows 版 Pandoc 资源");
        }
        var zipPath = Path.Combine(targetRoot, "pandoc.zip");
        await DownloadFileAsync(http, downloadUrl, zipPath, progress, ct);

        ZipFile.ExtractToDirectory(zipPath, targetRoot, true);

        var pandocPath = Directory
            .GetFiles(targetRoot, "pandoc.exe", SearchOption.AllDirectories)
            .FirstOrDefault();

        if (string.IsNullOrWhiteSpace(pandocPath))
        {
            return new PandocDownloadResult(false, null, "下载完成但未找到 pandoc.exe");
        }

        AppSettings.PandocPath = pandocPath;
        return new PandocDownloadResult(true, pandocPath, null);
    }

    private static async Task DownloadFileAsync(
        HttpClient http,
        string url,
        string outputPath,
        IProgress<double>? progress,
        CancellationToken ct)
    {
        using var response = await http.GetAsync(url, HttpCompletionOption.ResponseHeadersRead, ct);
        response.EnsureSuccessStatusCode();

        var total = response.Content.Headers.ContentLength ?? -1L;
        await using var stream = await response.Content.ReadAsStreamAsync(ct);
        await using var file = File.Create(outputPath);

        var buffer = new byte[81920];
        long readTotal = 0;
        int read;
        while ((read = await stream.ReadAsync(buffer, ct)) > 0)
        {
            await file.WriteAsync(buffer.AsMemory(0, read), ct);
            readTotal += read;
            if (total > 0 && progress is not null)
            {
                progress.Report(readTotal / (double)total);
            }
        }
    }
    private static string? GetWindowsZipUrl(string json)
    {
        using var doc = JsonDocument.Parse(json);
        if (!doc.RootElement.TryGetProperty("assets", out var assets))
        {
            return null;
        }

        foreach (var asset in assets.EnumerateArray())
        {
            if (!asset.TryGetProperty("name", out var nameProp))
            {
                continue;
            }

            var name = nameProp.GetString();
            if (string.IsNullOrWhiteSpace(name))
            {
                continue;
            }

            if (!name.Contains("windows-x86_64.zip", StringComparison.OrdinalIgnoreCase))
            {
                continue;
            }

            if (asset.TryGetProperty("browser_download_url", out var urlProp))
            {
                return urlProp.GetString();
            }
        }

        return null;
    }
}

public sealed record PandocInfo(string Path, string Version, string RawHeader);

public sealed record PandocRunResult(bool Succeeded, string StandardOutput, string StandardError);

public sealed record PandocDownloadResult(bool Succeeded, string? PandocPath, string? ErrorMessage);

public sealed record PandocDetectionResult(string? PandocPath, string Source, IReadOnlyList<string> Steps);
