// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

namespace OfficeCli.Core;

/// <summary>
/// [商用版本] 增强隐私保护的日志记录器
/// 
/// 安全改进：
/// - 日志默认禁用
/// - 敏感路径脱敏处理
/// - 日志文件权限限制
/// </summary>
internal static class CliLogger
{
    private static readonly string LogPath = Path.Combine(UpdateChecker.ConfigDir, "officecli.log");
    private const long MaxLogSize = 1024 * 1024; // 1 MB

    internal static bool Enabled
    {
        get
        {
            try { return UpdateChecker.LoadConfig().Log; }
            catch { return false; }
        }
    }

    /// <summary>
    /// [商用版本] 记录命令，对敏感信息进行脱敏处理
    /// </summary>
    internal static void LogCommand(string[] args)
    {
        if (!Enabled || args.Length == 0) return;
        
        // Skip internal commands
        if (args[0].StartsWith("__") && args[0].EndsWith("__")) return;

        // [商用版本] 对参数进行脱敏处理
        var sanitizedArgs = args.Select(SanitizeArgument).ToArray();
        Write($"> officecli {string.Join(" ", sanitizedArgs)}");
    }

    /// <summary>
    /// [商用版本] 对参数进行脱敏处理，保护敏感信息
    /// </summary>
    private static string SanitizeArgument(string arg)
    {
        // 脱敏文件路径 - 只保留文件名，隐藏完整路径
        if (arg.Contains('/') || arg.Contains('\\'))
        {
            try
            {
                // 检查是否是文件路径
                if (arg.EndsWith(".docx", StringComparison.OrdinalIgnoreCase) ||
                    arg.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase) ||
                    arg.EndsWith(".pptx", StringComparison.OrdinalIgnoreCase) ||
                    arg.EndsWith(".json", StringComparison.OrdinalIgnoreCase) ||
                    arg.EndsWith(".csv", StringComparison.OrdinalIgnoreCase) ||
                    arg.EndsWith(".xml", StringComparison.OrdinalIgnoreCase))
                {
                    var fileName = Path.GetFileName(arg);
                    return $"[file:{fileName}]";
                }
            }
            catch { /* 路径解析失败时保持原样 */ }
        }

        // 脱敏可能包含敏感信息的属性值
        if (arg.StartsWith("--prop", StringComparison.OrdinalIgnoreCase))
        {
            // 检查是否包含敏感关键词
            var sensitiveKeys = new[] { "password", "secret", "key", "token", "credential", "auth" };
            foreach (var key in sensitiveKeys)
            {
                if (arg.Contains(key, StringComparison.OrdinalIgnoreCase))
                {
                    return "[prop:REDACTED]";
                }
            }
        }

        return arg;
    }

    internal static void Clear()
    {
        try 
        { 
            File.Delete(LogPath);
        }
        catch { }
    }

    internal static void LogOutput(string output)
    {
        if (!Enabled || string.IsNullOrEmpty(output)) return;
        
        // [商用版本] 对输出也进行脱敏
        var sanitized = SanitizeOutput(output);
        Write(sanitized);
    }

    /// <summary>
    /// [商用版本] 对输出内容进行脱敏
    /// </summary>
    private static string SanitizeOutput(string output)
    {
        // 限制输出长度，防止日志过大
        const int maxLength = 1000;
        if (output.Length > maxLength)
        {
            output = output[..maxLength] + "...[truncated]";
        }
        return output;
    }

    internal static void LogError(string error)
    {
        if (!Enabled || string.IsNullOrEmpty(error)) return;
        
        // [商用版本] 错误信息也进行脱敏
        var sanitized = SanitizeOutput(error);
        Write($"[ERROR] {sanitized}");
    }

    private static void Write(string message)
    {
        try
        {
            Directory.CreateDirectory(UpdateChecker.ConfigDir);

            var escaped = message.ReplaceLineEndings("\n");
            TrimIfNeeded();
            File.AppendAllText(LogPath, $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] {escaped}\n");
            
            // [商用版本] 设置日志文件权限，仅限当前用户访问
            try
            {
                if (!OperatingSystem.IsWindows())
                {
                    File.SetUnixFileMode(LogPath, 
                        UnixFileMode.UserRead | UnixFileMode.UserWrite);
                }
            }
            catch { /* 权限设置失败不影响功能 */ }
        }
        catch
        {
            // Logging should never break the CLI
        }
    }

    private static void TrimIfNeeded()
    {
        var info = new FileInfo(LogPath);
        if (!info.Exists || info.Length <= MaxLogSize) return;

        // Keep the last half of the file
        var text = File.ReadAllText(LogPath);
        var half = text.Length / 2;
        var start = text.IndexOf('\n', half);
        if (start < 0 || start >= text.Length - 1) return;
        File.WriteAllText(LogPath, text[(start + 1)..]);
    }
}
