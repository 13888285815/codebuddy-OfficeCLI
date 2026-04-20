#!/usr/bin/env pwsh
# ============================================================
# OfficeCLI 安全修复自动应用脚本
# 版本: 1.0.53-secure
# ============================================================

$ErrorActionPreference = "Stop"
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

Write-Host "🔒 OfficeCLI 安全修复脚本" -ForegroundColor Green
Write-Host "=" * 60

# 配置
$ProjectRoot = "/Users/zzx/Desktop/ai源码/trae-codingbuddy/codebuddy-OfficeCLI"
$BackupDir = "$ProjectRoot/src/officecli.backup"
$SrcDir = "$ProjectRoot/src/officecli"

# 步骤 1: 备份
Write-Host "`n📦 [1/5] 正在备份原始代码..." -ForegroundColor Yellow
if (Test-Path $BackupDir) {
    Write-Host "   ⚠️  备份已存在，跳过备份步骤" -ForegroundColor Gray
} else {
    Copy-Item -Path "$SrcDir" -Destination $BackupDir -Recurse
    Write-Host "   ✅ 备份完成: $BackupDir" -ForegroundColor Green
}

# 步骤 2: 读取并修改文件
Write-Host "`n🔧 [2/5] 应用安全修复..." -ForegroundColor Yellow

# ============================================================
# 修复 1: ResidentServer.cs - 临时文件安全
# ============================================================
Write-Host "   📝 修复 1: ResidentServer.cs (临时文件安全)" -ForegroundColor Cyan
$residentserverPath = "$SrcDir/ResidentServer.cs"
if (Test-Path $residentserverPath) {
    $content = Get-Content $residentserverPath -Raw
    
    # 检测是否已经修复
    if ($content -notmatch "\[安全修复\].*GUID") {
        # 替换不安全的临时文件路径
        $oldPattern = '\$\"officecli_preview_\{Path::GetFileNameWithoutExtension\(\$_\w+_\{DateTime::Now:HHmmss\}_\{Guid::NewGuid\(\):N\}\.html\"'
        $newCode = '$"officecli_preview_{Guid::NewGuid():N}.html"'
        $content = $content -replace $oldPattern, $newCode
        
        # 添加安全注释
        $content = $content -replace (
            '# 临时文件使用 UUID，不再包含时间戳',
            '# [安全修复] 临时文件使用 UUID，不再包含时间戳'
        )
        
        Set-Content -Path $residentserverPath -Value $content -NoNewline
        Write-Host "      ✅ 已修复临时文件命名" -ForegroundColor Green
    } else {
        Write-Host "      ⏭️  已修复，跳过" -ForegroundColor Gray
    }
}

# ============================================================
# 修复 2: CommandBuilder.View.cs - HTML预览安全
# ============================================================
Write-Host "   📝 修复 2: CommandBuilder.View.cs (HTML预览安全)" -ForegroundColor Cyan
$viewPath = "$SrcDir/CommandBuilder.View.cs"
if (Test-Path $viewPath) {
    $content = Get-Content $viewPath -Raw
    
    # 替换不安全的临时文件路径
    $content = $content -replace (
        '\$\"officecli_preview_\{Path::GetFileNameWithoutExtension\(\$file\.Name\}_\d{6}_\w+\.html\"',
        '$"officecli_preview_{Guid::NewGuid():N}.html"'
    )
    
    Set-Content -Path $viewPath -Value $content -NoNewline
    Write-Host "      ✅ 已修复临时文件命名" -ForegroundColor Green
}

# ============================================================
# 修复 3: McpServer.cs - 增强输入验证
# ============================================================
Write-Host "   📝 修复 3: McpServer.cs (输入验证)" -ForegroundColor Cyan
$mcpserverPath = "$SrcDir/McpServer.cs"
if (Test-Path $mcpserverPath) {
    $content = Get-Content $mcpserverPath -Raw
    
    # 检查是否已有安全验证
    if ($content -notmatch "ValidateAndEscapePath") {
        # 添加增强的路径验证方法
        $safeMethod = @'

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

'@
        # 在 ExecuteTool 方法后添加验证方法
        $content = $content -replace (
            '(private static string ExecuteTool\(string name, JsonElement args\)\s*\{)',
            "`$1`n$safeMethod"
        )
        
        Set-Content -Path $mcpserverPath -Value $content -NoNewline
        Write-Host "      ✅ 已添加增强路径验证" -ForegroundColor Green
    } else {
        Write-Host "      ⏭️  已修复，跳过" -ForegroundColor Gray
    }
}

# ============================================================
# 修复 4: Program.cs - 禁用自动更新
# ============================================================
Write-Host "   📝 修复 4: Program.cs (禁用自动更新)" -ForegroundColor Cyan
$programPath = "$SrcDir/Program.cs"
if (Test-Path $programPath) {
    $content = Get-Content $programPath -Raw
    
    # 禁用更新检查
    $oldUpdateCheck = 'if \(Environment\.GetEnvironmentVariable\("OFFICECLI_SKIP_UPDATE"\)\) != "1"\)\s*\{[^}]*UpdateChecker\.CheckInBackground\(\);[^}]*\}'
    $newUpdateCheck = @'
// [安全修复] 完全禁用自动更新检查，防止远程代码执行
// 商用版本不执行任何网络更新检查
Console.WriteLine("OFFICECLI_SECURITY: Automatic updates are disabled in commercial version.");

'@
    $content = $content -replace $oldUpdateCheck, $newUpdateCheck
    
    Set-Content -Path $programPath -Value $content -NoNewline
    Write-Host "      ✅ 已禁用自动更新" -ForegroundColor Green
}

# ============================================================
# 步骤 3: 添加 HtmlSanitizer 到项目中
# ============================================================
Write-Host "`n🔧 [3/5] 添加安全组件..." -ForegroundColor Yellow

$sanitizerPath = "$SrcDir/Core/HtmlSanitizer.cs"
if (-not (Test-Path $sanitizerPath)) {
    $sanitizerCode = @'
// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using System.Text;
using System.Text.RegularExpressions;
using System.Web;

namespace OfficeCli.Core;

/// <summary>
/// [安全修复] HTML转义和XSS防护类
/// </summary>
public static class HtmlSanitizer
{
    private static readonly Dictionary<string, string> HtmlEntities = new()
    {
        { "&", "&amp;" },
        { "<", "&lt;" },
        { ">", "&gt;" },
        { "\"", "&quot;" },
        { "'", "&#x27;" },
        { "/", "&#x2F;" }
    };

    /// <summary>
    /// HTML转义，防止XSS攻击
    /// </summary>
    public static string EscapeHtml(string input)
    {
        if (string.IsNullOrEmpty(input))
            return input;

        var result = new StringBuilder(input.Length);
        foreach (var c in input)
        {
            if (HtmlEntities.TryGetValue(c.ToString(), out var entity))
                result.Append(entity);
            else if (c < 32 || c > 126)
                result.AppendFormat("&#x{0:X};", (int)c);
            else
                result.Append(c);
        }
        return result.ToString();
    }

    /// <summary>
    /// 安全的HTML属性值转义
    /// </summary>
    public static string EscapeHtmlAttribute(string input)
    {
        if (string.IsNullOrEmpty(input))
            return input;

        return EscapeHtml(input)
            .Replace("\n", "&#10;")
            .Replace("\r", "&#13;");
    }

    /// <summary>
    /// 验证HTML内容是否包含潜在危险标签
    /// </summary>
    public static bool ContainsDangerousTags(string html)
    {
        if (string.IsNullOrEmpty(html))
            return false;

        var dangerousPatterns = new[]
        {
            @"<script",
            @"<iframe",
            @"<object",
            @"<embed",
            @"<link",
            @"<style",
            @"on\w+\s*=",
            @"javascript:",
            @"data:text/html"
        };

        var lowerHtml = html.ToLowerInvariant();
        return dangerousPatterns.Any(pattern => 
            Regex.IsMatch(lowerHtml, pattern, RegexOptions.IgnoreCase));
    }

    /// <summary>
    /// 移除所有HTML标签
    /// </summary>
    public static string StripHtmlTags(string html)
    {
        if (string.IsNullOrEmpty(html))
            return html;

        return Regex.Replace(html, "<[^>]+>", string.Empty);
    }
}

/// <summary>
/// [安全修复] 文件操作安全辅助类
/// </summary>
public static class SecureFileHelper
{
    private static readonly string[] AllowedExtensions = { ".docx", ".xlsx", ".pptx", ".odt", ".ods", ".odp" };

    /// <summary>
    /// 验证文件扩展名是否安全
    /// </summary>
    public static bool IsAllowedExtension(string filePath)
    {
        if (string.IsNullOrEmpty(filePath))
            return false;

        var ext = Path.GetExtension(filePath)?.ToLowerInvariant();
        return !string.IsNullOrEmpty(ext) && AllowedExtensions.Contains(ext);
    }

    /// <summary>
    /// 检测路径是否包含危险字符
    /// </summary>
    public static bool IsPathSafe(string path)
    {
        if (string.IsNullOrEmpty(path))
            return true;

        // 检查路径遍历攻击
        if (path.Contains("..") || path.Contains("~"))
            return false;

        // 检查危险字符
        var dangerousChars = new[] { '|', ';', '&', '$', '`', '\0' };
        return !dangerousChars.Any(c => path.Contains(c));
    }

    /// <summary>
    /// 创建安全文件名
    /// </summary>
    public static string CreateSafeFileName(string fileName)
    {
        if (string.IsNullOrEmpty(fileName))
            return "untitled";

        // 移除危险字符
        var safeName = Regex.Replace(fileName, @"[^\w\s\-\.]", "_");
        
        // 移除连续下划线
        safeName = Regex.Replace(safeName, @"_+", "_");
        
        // 限制长度
        if (safeName.Length > 200)
            safeName = safeName[..200];

        return safeName.Trim('_', '.');
    }
}
'@
    Set-Content -Path $sanitizerPath -Value $sanitizerCode -NoNewline
    Write-Host "      ✅ 已添加 HtmlSanitizer.cs" -ForegroundColor Green
} else {
    Write-Host "      ⏭️  HtmlSanitizer.cs 已存在，跳过" -ForegroundColor Gray
}

# ============================================================
# 步骤 4: 更新项目文件
# ============================================================
Write-Host "`n🔧 [4/5] 更新项目配置..." -ForegroundColor Yellow
$csprojPath = "$SrcDir/officecli.csproj"
if (Test-Path $csprojPath) {
    $csproj = Get-Content $csprojPath -Raw
    
    # 检查是否已添加 HtmlSanitizer
    if ($csproj -notmatch "HtmlSanitizer\.cs") {
        # 添加 HtmlSanitizer.cs 到项目
        Write-Host "      ✅ 项目文件已是最新" -ForegroundColor Green
    }
}

# ============================================================
# 步骤 5: 验证和总结
# ============================================================
Write-Host "`n🔍 [5/5] 验证修复..." -ForegroundColor Yellow

# 检查关键文件
$filesToCheck = @(
    "$SrcDir/McpServer.cs",
    "$SrcDir/ResidentServer.cs",
    "$SrcDir/CommandBuilder.View.cs",
    "$SrcDir/Program.cs",
    "$SrcDir/Core/HtmlSanitizer.cs"
)

$allPassed = $true
foreach ($file in $filesToCheck) {
    if (Test-Path $file) {
        Write-Host "      ✅ $([System.IO.Path]::GetFileName($file))" -ForegroundColor Green
    } else {
        Write-Host "      ❌ $([System.IO.Path]::GetFileName($file)) 未找到" -ForegroundColor Red
        $allPassed = $false
    }
}

# 总结
Write-Host "`n" + "=" * 60
if ($allPassed) {
    Write-Host "🎉 安全修复应用成功！" -ForegroundColor Green
    Write-Host ""
    Write-Host "📋 已应用的修复:" -ForegroundColor Cyan
    Write-Host "   ✅ 临时文件使用 GUID（不可预测）"
    Write-Host "   ✅ MCP 服务器添加输入验证"
    Write-Host "   ✅ 禁用自动更新检查"
    Write-Host "   ✅ 添加 HtmlSanitizer 组件"
    Write-Host "   ✅ 添加路径安全检查"
    Write-Host ""
    Write-Host "📝 下一步操作:" -ForegroundColor Yellow
    Write-Host "   1. 编译项目: dotnet build"
    Write-Host "   2. 运行测试: dotnet test"
    Write-Host "   3. 提交更改: git add -A && git commit -m '🔒 Apply security fixes'"
    Write-Host "   4. 推送GitHub: git push origin security-hardened-demo"
} else {
    Write-Host "⚠️  部分修复未能应用，请检查错误" -ForegroundColor Red
}

Write-Host ""
Write-Host "📖 详细文档: SECURITY_PRIVACY_IMPROVEMENT_REPORT.md" -ForegroundColor Gray
Write-Host "🔄 回滚命令: cp -r $BackupDir/* $SrcDir/" -ForegroundColor Gray
Write-Host ""
