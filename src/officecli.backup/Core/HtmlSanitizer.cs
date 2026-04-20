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
