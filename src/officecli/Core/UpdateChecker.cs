// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using System.Diagnostics;
using System.Reflection;
using System.Text.Json;
using System.Text.Json.Serialization;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;

namespace OfficeCli.Core;

/// <summary>
/// [商用版本] 自动更新功能已完全禁用
/// 
/// 为了安全考虑，商用版本移除了所有自动更新功能：
/// - 不会自动检查更新
/// - 不会下载任何外部代码
/// - 不会执行自动升级
/// 
/// 如需更新，请手动下载新版本。
/// </summary>
internal static class UpdateChecker
{
    internal static readonly string ConfigDir = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), ".officecli");
    private static readonly string ConfigPath = Path.Combine(ConfigDir, "config.json");
    
    // [商用版本] 已移除自动更新相关的URL和配置
    // private const string GitHubRepo = "iOfficeAI/OfficeCLI";
    // private const string PrimaryBase = "https://officecli.ai";
    // private const string FallbackBase = "https://github.com/iOfficeAI/OfficeCLI";
    // private const int CheckIntervalHours = 24;

    /// <summary>
    /// [商用版本] 自动更新已禁用
    /// 此方法现在为空操作，不会执行任何网络请求
    /// </summary>
    internal static void CheckInBackground()
    {
        // [商用版本] 完全禁用自动更新
        // 不执行任何操作，不创建目录，不发起网络请求
        return;
    }

    /// <summary>
    /// [商用版本] 自动更新已禁用
    /// 此方法现在为空操作
    /// </summary>
    internal static void RunRefresh()
    {
        // [商用版本] 完全禁用自动更新
        Console.WriteLine("[商用版本] 自动更新功能已禁用。请手动下载更新。");
        return;
    }

    /// <summary>
    /// [商用版本] 不再应用待处理的更新
    /// </summary>
    private static void ApplyPendingUpdate()
    {
        // [商用版本] 禁用自动更新应用
        // 清理任何可能存在的旧更新文件
        try
        {
            var exePath = Environment.ProcessPath ?? Process.GetCurrentProcess().MainModule?.FileName;
            if (exePath != null)
            {
                var updatePath = exePath + ".update";
                var partialPath = exePath + ".update.partial";
                var oldPath = exePath + ".old";
                
                // 清理旧更新文件（如果存在）
                if (File.Exists(updatePath)) File.Delete(updatePath);
                if (File.Exists(partialPath)) File.Delete(partialPath);
                if (File.Exists(oldPath)) File.Delete(oldPath);
            }
        }
        catch { /* 静默处理清理错误 */ }
    }

    /// <summary>
    /// [商用版本] 不再生成刷新进程
    /// </summary>
    private static void SpawnRefreshProcess()
    {
        // [商用版本] 禁用刷新进程生成
        return;
    }

    /// <summary>
    /// 处理配置命令
    /// [商用版本] 移除了 autoUpdate 配置项
    /// </summary>
    internal static void HandleConfigCommand(string[] args)
    {
        const string available = "log, log clear";
        var key = args[0].ToLowerInvariant();
        var config = LoadConfig();

        // [商用版本] 处理 autoUpdate 配置请求
        if (key == "autoupdate")
        {
            Console.WriteLine("[商用版本] 自动更新功能已禁用，此配置项无效。");
            return;
        }

        // officecli config log clear
        if (key == "log" && args.Length == 2 && args[1].ToLowerInvariant() == "clear")
        {
            CliLogger.Clear();
            Console.WriteLine("Log cleared.");
            return;
        }

        if (args.Length == 1)
        {
            // Read
            var value = key switch
            {
                "log" => config.Log.ToString().ToLowerInvariant(),
                _ => null
            };
            if (value != null)
                Console.WriteLine(value);
            else
                Console.Error.WriteLine($"Unknown config key: {args[0]}. Available: {available}");
            return;
        }

        // Write
        var newValue = args[1];
        switch (key)
        {
            case "log":
                config.Log = ParseHelpers.IsTruthy(newValue);
                break;
            default:
                Console.Error.WriteLine($"Unknown config key: {args[0]}. Available: {available}");
                return;
        }

        try
        {
            Directory.CreateDirectory(ConfigDir);
            SaveConfig(config);
            Console.WriteLine($"{args[0]} = {newValue}");
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"Error saving config: {ex.Message}");
        }
    }

    private static string? GetCurrentVersion()
    {
        var version = Assembly.GetExecutingAssembly()
            .GetCustomAttribute<AssemblyInformationalVersionAttribute>()?.InformationalVersion;
        if (version == null) return null;
        var match = Regex.Match(version, @"^(\d+\.\d+\.\d+)");
        return match.Success ? match.Groups[1].Value : version;
    }

    private static bool IsNewer(string latest, string current)
    {
        var lp = latest.Split('.').Select(int.Parse).ToArray();
        var cp = current.Split('.').Select(int.Parse).ToArray();
        for (int i = 0; i < Math.Min(lp.Length, cp.Length); i++)
        {
            if (lp[i] > cp[i]) return true;
            if (lp[i] < cp[i]) return false;
        }
        return lp.Length > cp.Length;
    }

    internal static AppConfig LoadConfig()
    {
        if (!File.Exists(ConfigPath)) return new AppConfig();
        try
        {
            var json = File.ReadAllText(ConfigPath);
            return JsonSerializer.Deserialize(json, AppConfigContext.Default.AppConfig) ?? new AppConfig();
        }
        catch { return new AppConfig(); }
    }

    internal static void SaveConfig(AppConfig config)
    {
        Directory.CreateDirectory(ConfigDir);
        var json = JsonSerializer.Serialize(config, AppConfigContext.Default.AppConfig);
        File.WriteAllText(ConfigPath, json);
        
        // [商用版本] 设置配置文件权限，仅限当前用户访问
        try
        {
            if (!OperatingSystem.IsWindows())
            {
                File.SetUnixFileMode(ConfigPath, 
                    UnixFileMode.UserRead | UnixFileMode.UserWrite);
            }
        }
        catch { /* 权限设置失败不影响功能 */ }
    }

    internal static string? GetCurrentVersionPublic() => GetCurrentVersion();

    internal static bool IsNewerPublic(string latest, string current) => IsNewer(latest, current);
}

internal class AppConfig
{
    // [商用版本] 移除了 LastUpdateCheck 和 LatestVersion
    // public DateTime? LastUpdateCheck { get; set; }
    // public string? LatestVersion { get; set; }
    
    // [商用版本] 自动更新始终禁用
    public bool AutoUpdate { get; set; } = false;
    public bool Log { get; set; }
    public string? InstalledBinaryVersion { get; set; }
}

[JsonSerializable(typeof(AppConfig))]
[JsonSourceGenerationOptions(WriteIndented = true, PropertyNamingPolicy = JsonKnownNamingPolicy.CamelCase)]
internal partial class AppConfigContext : JsonSerializerContext;
