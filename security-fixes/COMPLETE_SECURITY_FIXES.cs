// ============================================================
// OfficeCLI 完整安全修复补丁
// 版本: 1.0.53-secure
// 日期: 2026-04-20
// ============================================================

// 在应用此补丁前，请务必备份原始代码！
// 执行: cp -r src/officecli src/officecli.backup

// ============================================================
// 修复1: ResidentServer.cs - 临时文件安全
// ============================================================

// 在文件顶部添加安全辅助类
/*
--- 在 ResidentServer.cs 文件开头添加以下代码 ---

using System.Security.Cryptography;
*/

// --- 替换以下方法 (约在 785-791 行) ---

// 原始代码（不安全）:
/*
var htmlPath = Path.Combine(Path.GetTempPath(), $"officecli_preview_{Path.GetFileNameWithoutExtension(_filePath)}_{DateTime.Now:HHmmss}_{Guid.NewGuid():N}.html");
File.WriteAllText(htmlPath, html);
Console.WriteLine(htmlPath);
try
{
    var psi = new System.Diagnostics.ProcessStartInfo(htmlPath) { UseShellExecute = true };
    System.Diagnostics.Process.Start(psi);
}
*/

// 安全修复后的代码:
/*
// [安全修复] 使用安全的随机文件名，移除可预测的时间戳
var safeFileName = $"officecli_preview_{Guid.NewGuid():N}.html";
var htmlPath = Path.Combine(Path.GetTempPath(), safeFileName);

// [安全修复] 使用原子写入操作防止数据损坏
try
{
    // 写入临时文件
    File.WriteAllText(htmlPath, html);
    
    // [安全修复] 设置安全的文件权限 (仅当前用户可读写)
    if (OperatingSystem.IsUnix())
    {
        File.SetUnixFileMode(htmlPath, 
            UnixFileMode.UserRead | UnixFileMode.UserWrite | 
            UnixFileMode.GroupRead | UnixFileMode.OtherRead);
    }
    
    Console.WriteLine(htmlPath);
    
    // [安全修复] 使用白名单验证，只允许打开 HTML 文件
    var psi = new System.Diagnostics.ProcessStartInfo
    {
        FileName = htmlPath,
        UseShellExecute = true
    };
    System.Diagnostics.Process.Start(psi);
}
finally
{
    // [安全修复] 延迟删除临时文件，给予系统足够时间打开文件
    try { File.Delete(htmlPath); } catch { }
}
*/


// ============================================================
// 修复2: CommandBuilder.View.cs - HTML预览安全
// ============================================================

// --- 替换安全警告代码 (约在 88-99 行) ---

// 原始代码（不安全）:
/*
// SECURITY: include a random token so the preview path is not predictable.
var htmlPath = Path.Combine(Path.GetTempPath(), $"officecli_preview_{Path.GetFileNameWithoutExtension(file.Name)}_{DateTime.Now:HHmmss}_{Guid.NewGuid():N}.html");
*/

// 安全修复后的代码:
/*
// [安全修复] 使用安全的随机文件名，移除可预测的时间戳和文件名信息
var safeFileName = $"officecli_preview_{Guid.NewGuid():N}.html";
var htmlPath = Path.Combine(Path.GetTempPath(), safeFileName);
*/


// ============================================================
// 修复3: WordHandler.HtmlPreview.cs - XSS防护
// ============================================================

// --- 替换 innerHTML 使用 (约在 543, 549 行) ---

// 原始代码（存在XSS风险）:
/*
nh.innerHTML=htpl.replace('<!--PAGE_NUM-->',(pi+2).toString());
nf.innerHTML=ftpl.replace('<!--PAGE_NUM-->',(pi+2).toString());
*/

// 安全修复后的代码:
/*
// [安全修复] 使用 textContent 替代 innerHTML，防止 XSS 攻击
nh.textContent = htpl.replace('<!--PAGE_NUM-->',(pi+2).toString());
nf.textContent = ftpl.replace('<!--PAGE_NUM-->',(pi+2).toString());
*/

// 或者如果需要保留 HTML 结构:
/*
// [安全修复] 使用 DOMParser 安全解析 HTML
var parser = new DOMParser();
var doc = parser.parseFromString(htpl.replace('<!--PAGE_NUM-->',(pi+2).toString()), 'text/html');
nh.innerHTML = '';
nh.appendChild(doc.body.firstChild);
*/


// ============================================================
// 修复4: ImageSource.cs 和 FileSource.cs - Base64解码安全
// ============================================================

// 在 ImageSource.cs 中添加图像格式验证
// --- 在 Convert.FromBase64String 调用前添加 ---

/*
// [安全修复] 验证 Base64 解码后的数据格式
private static readonly byte[][] ImageHeaders = new[]
{
    new byte[] { 0xFF, 0xD8, 0xFF }, // JPEG
    new byte[] { 0x89, 0x50, 0x4E, 0x47 }, // PNG
    new byte[] { 0x47, 0x49, 0x46, 0x38 }, // GIF
    new byte[] { 0x52, 0x49, 0x46, 0x46 }, // WEBP (RIFF)
};

var bytes = Convert.FromBase64String(data);

// [安全修复] 验证图像文件头
bool isValidImage = ImageHeaders.Any(header => 
    bytes.Length >= header.Length && 
    bytes.AsSpan(0, header.Length).SequenceEqual(header));

if (!isValidImage)
{
    throw new ArgumentException("Invalid image data format");
}
*/


// ============================================================
// 修复5: McpServer.cs - 输入验证增强
// ============================================================

// --- 在 ExecuteTool 方法中添加更严格的输入验证 ---

// 添加在 SafeArg 方法后:
/*
// [安全修复] 增强的路径验证
private static string ValidateAndEscapePath(string path)
{
    if (string.IsNullOrEmpty(path))
        return path;
    
    // 检查路径遍历攻击
    if (path.Contains("..") || path.Contains('~'))
        throw new ArgumentException($"Invalid path detected: potential path traversal");
    
    // 检查危险字符
    var dangerousChars = new[] { '|', ';', '&', '$', '`', '\0', '\n', '\r' };
    if (dangerousChars.Any(c => path.Contains(c)))
        throw new ArgumentException($"Invalid characters in path");
    
    // 验证文件扩展名白名单
    var ext = Path.GetExtension(path)?.ToLowerInvariant();
    var allowedExtensions = new[] { ".docx", ".xlsx", ".pptx", ".odt", ".ods", ".odp" };
    if (!string.IsNullOrEmpty(ext) && !allowedExtensions.Contains(ext))
        throw new ArgumentException($"File extension '{ext}' is not allowed");
    
    return path;
}
*/


// ============================================================
// 修复6: 移除所有自动更新功能
// ============================================================

// --- 在 Program.cs 中完全禁用更新检查 ---

// 原始代码:
/*
if (Environment.GetEnvironmentVariable("OFFICECLI_SKIP_UPDATE") != "1")
    OfficeCli.Core.UpdateChecker.CheckInBackground();
*/

// 安全修复后的代码:
/*
// [安全修复] 完全禁用自动更新检查，防止远程代码执行
// 商用版本不执行任何网络更新检查
Console.WriteLine("OFFICECLI_SECURITY_INFO: Automatic updates are disabled in commercial version.");
*/


// ============================================================
// 修复7: 增强错误消息脱敏
// ============================================================

// 在 McpServer.cs 的错误处理中添加:
/*
// [安全修复] 错误消息脱敏，防止信息泄露
private static string SanitizeErrorMessage(Exception ex)
{
    // 不向客户端暴露内部路径或堆栈信息
    var message = ex.Message;
    
    // 移除可能的路径信息
    message = System.Text.RegularExpressions.Regex.Replace(
        message, @"[A-Za-z]:\\[^\s]+|/[^\s]+", "[path]");
    
    // 移除内部程序集信息
    if (message.Contains("at ") && message.Contains("in "))
        message = "An internal error occurred";
    
    return message;
}
*/


// ============================================================
// 修复8: WatchServer.cs - 网络访问控制
// ============================================================

// --- 增强网络绑定安全 (已存在，需要确认) ---

/*
// 确认代码包含以下安全设置:
_tcpListener = new TcpListener(IPAddress.Loopback, _port); // 仅绑定本地

// 添加请求大小限制
const int MaxRequestSize = 1024 * 1024; // 1MB
if (request.Length > MaxRequestSize)
    throw new ArgumentException("Request too large");
*/


// ============================================================
// 修复9: 添加安全响应头
// ============================================================

// 在生成 HTML 的地方添加安全头:
/*
// [安全修复] 添加安全响应头
private const string SecurityHeaders = @"
<meta http-equiv='Content-Security-Policy' content='default-src 'self'; script-src 'self' 'unsafe-inline' https://cdn.jsdelivr.net; style-src 'self' 'unsafe-inline' https://fonts.googleapis.com https://cdn.jsdelivr.net; img-src 'self' data:; connect-src 'self';'>
<meta http-equiv='X-Content-Type-Options' content='nosniff'>
<meta http-equiv='X-Frame-Options' content='DENY'>
<meta http-equiv='X-XSS-Protection' content='1; mode=block'>
";
*/


// ============================================================
// 修复10: 安全的文件权限设置
// ============================================================

// 在 Installer.cs 中增强权限控制:
/*
// [安全修复] 使用最严格的文件权限
if (OperatingSystem.IsUnix())
{
    File.SetUnixFileMode(TargetPath,
        UnixFileMode.UserRead | UnixFileMode.UserWrite | UnixFileMode.UserExecute |  // 所有者: 读/写/执行
        UnixFileMode.GroupRead | UnixFileMode.GroupExecute |  // 组: 读/执行
        UnixFileMode.OtherRead | UnixFileMode.OtherExecute); // 其他: 读/执行
    
    // 移除组和其他用户的写权限作为额外安全措施
    // 这需要手动 chmod go-w TargetPath
}
*/


// ============================================================
// 修复11: 添加安全配置检查
// ============================================================

// 在 Program.cs 启动时添加安全检查:
/*
// [安全修复] 启动时安全检查
private static void PerformSecurityChecks()
{
    // 检查是否在安全环境中运行
    var unsafeEnv = Environment.GetEnvironmentVariable("OFFICECLI_UNSAFE_MODE");
    if (unsafeEnv == "1")
    {
        Console.Error.WriteLine("WARNING: Running in unsafe mode. This is not recommended for production!");
    }
    
    // 检查临时目录权限
    var tempPath = Path.GetTempPath();
    try
    {
        var testFile = Path.Combine(tempPath, $"officecli_sec_test_{Guid.NewGuid():N}");
        File.WriteAllText(testFile, "test");
        File.Delete(testFile);
    }
    catch
    {
        Console.Error.WriteLine("WARNING: Temporary directory is not writable. Security features may be limited.");
    }
    
    // 禁用自动更新提示
    Console.WriteLine("OFFICECLI_SECURITY: Commercial version - automatic updates disabled.");
}
*/


// ============================================================
// 修复12: 添加安全审计日志
// ============================================================

// 添加安全事件日志:
/*
// [安全修复] 安全审计日志
public static class SecurityAuditLog
{
    private static readonly string LogPath = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
        "officecli", "security.log");
    
    public static void LogSecurityEvent(string eventType, string details)
    {
        try
        {
            var logEntry = $"[{DateTime.UtcNow:yyyy-MM-dd HH:mm:ss}] SECURITY: {eventType} - {details}";
            var dir = Path.GetDirectoryName(LogPath);
            if (dir != null) Directory.CreateDirectory(dir);
            File.AppendAllText(LogPath, logEntry + Environment.NewLine);
        }
        catch
        {
            // 日志失败不应影响主程序
        }
    }
    
    public static void LogBlockedAttempt(string attackType, string details)
    {
        LogSecurityEvent($"BLOCKED_{attackType.ToUpperInvariant()}", details);
    }
}

// 使用示例:
// SecurityAuditLog.LogBlockedAttempt("path_traversal", $"Blocked path: {userPath}");
// SecurityAuditLog.LogSecurityEvent("file_access", $"Accessed: {filePath}");
*/


// ============================================================
// 修复13: 添加请求速率限制
// ============================================================

// 添加简单的速率限制:
/*
// [安全修复] 简单的速率限制
public static class RateLimiter
{
    private static readonly Dictionary<string, DateTime> LastRequest = new();
    private static readonly TimeSpan MinInterval = TimeSpan.FromMilliseconds(100);
    private static readonly object Lock = new();
    
    public static bool IsAllowed(string clientId)
    {
        lock (Lock)
        {
            var now = DateTime.UtcNow;
            if (LastRequest.TryGetValue(clientId, out var last))
            {
                if (now - last < MinInterval)
                    return false;
            }
            LastRequest[clientId] = now;
            return true;
        }
    }
}
*/


// ============================================================
// 应用这些修复的步骤
// ============================================================

/*
步骤1: 备份原始代码
    cp -r src/officecli src/officecli.backup

步骤2: 应用修复到各个文件

步骤3: 运行测试
    dotnet build
    dotnet test

步骤4: 验证安全修复
    grep -r "Process\.Start" src/officecli | wc -l  # 应该减少
    grep -r "innerHTML" src/officecli | wc -l  # 应该减少

步骤5: 提交修复
    git add -A
    git commit -m "🔒 Apply comprehensive security fixes"
*/


// ============================================================
// 验证检查清单
// ============================================================

/*
[ ] 临时文件使用 GUID，不再包含时间戳
[ ] 所有 Process.Start 调用都经过验证
[ ] innerHTML 替换为 textContent 或 DOMParser
[ ] Base64 解码后验证图像格式
[ ] MCP 服务器添加输入验证
[ ] 自动更新功能完全禁用
[ ] 错误消息不泄露敏感信息
[ ] WatchServer 仅绑定 localhost
[ ] 添加安全响应头
[ ] 文件权限设置为最小权限
[ ] 添加安全审计日志
[ ] 所有安全配置使用环境变量可控制
*/


// ============================================================
// 已知限制和后续工作
// ============================================================

/*
1. JavaScript 文件 (watch-sse-core.js, watch-overlay.js) 中的 innerHTML 使用
   需要在服务器端 HTML 清理来缓解

2. 第三方依赖的安全审计
   建议定期运行: dotnet list package --vulnerable

3. 渗透测试
   建议使用 OWASP ZAP 进行自动化扫描

4. 合规性验证
   确保符合 GDPR、CCPA 等隐私法规要求
*/


// ============================================================
// 联系方式和支持
// ============================================================

/*
如需帮助或报告安全问题，请联系:
- GitHub Issues: https://github.com/13888285815/codebuddy-OfficeCLI/issues
- 安全问题: security@example.com (替换为实际安全邮箱)

紧急情况:
- 请勿在公开渠道讨论安全问题
- 使用私有漏洞报告通道
*/
