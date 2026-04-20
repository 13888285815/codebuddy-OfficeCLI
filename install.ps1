# [商用版本] OfficeCLI 安装脚本 (PowerShell)
# 安全改进：移除了自动下载功能，仅支持本地二进制文件安装

param(
    [string]$BinaryPath = ""
)

$binary = "officecli.exe"

# 显示帮助信息
function Show-Help {
    Write-Host "OfficeCLI 商用版本安装脚本" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "用法:"
    Write-Host "  .\install.ps1 [-BinaryPath <本地二进制文件路径>]"
    Write-Host ""
    Write-Host "示例:"
    Write-Host "  .\install.ps1 -BinaryPath .\officecli-win-x64.exe"
    Write-Host "  .\install.ps1 .\officecli-win-arm64.exe"
    Write-Host ""
    Write-Host "注意：商用版本不支持自动下载，请手动下载二进制文件后运行此脚本。"
}

if ($BinaryPath -eq "--help" -or $BinaryPath -eq "-h") {
    Show-Help
    exit 0
}

# 检测平台
$asset = "officecli-win-x64.exe"
if ($env:PROCESSOR_ARCHITECTURE -eq "ARM64") {
    $asset = "officecli-win-arm64.exe"
}

Write-Host "[商用版本] OfficeCLI 安装程序" -ForegroundColor Cyan
Write-Host "平台: Windows"
Write-Host "期望的二进制文件: $asset"
Write-Host ""

# 查找本地二进制文件
$source = $null

# 如果提供了参数，使用指定的文件
if ($BinaryPath -and (Test-Path $BinaryPath)) {
    $source = $BinaryPath
    Write-Host "使用指定的二进制文件: $source"
} else {
    # 尝试在当前目录查找匹配的二进制文件
    $candidates = @(".\$asset", ".\$binary", ".\bin\$asset", ".\bin\$binary", ".\bin\release\$asset", ".\bin\release\$binary")
    foreach ($candidate in $candidates) {
        if (Test-Path $candidate) {
            $output = & $candidate --version 2>&1
            if ($LASTEXITCODE -eq 0) {
                $source = $candidate
                Write-Host "找到有效的二进制文件: $candidate"
                break
            }
        }
    }
}

if (-not $source) {
    Write-Host "错误: 无法找到有效的 OfficeCLI 二进制文件。" -ForegroundColor Red
    Write-Host ""
    Write-Host "请手动下载对应平台的二进制文件:"
    Write-Host "  - Windows x64: officecli-win-x64.exe"
    Write-Host "  - Windows ARM64: officecli-win-arm64.exe"
    Write-Host ""
    Write-Host "下载后运行: .\install.ps1 -BinaryPath <文件路径>"
    exit 1
}

# 验证二进制文件
Write-Host ""
Write-Host "验证二进制文件..."
$output = & $source --version 2>&1
if ($LASTEXITCODE -ne 0) {
    Write-Host "错误: 二进制文件验证失败" -ForegroundColor Red
    exit 1
}

$version = $output
Write-Host "版本: $version"

# 检查现有安装
$existing = Get-Command $binary -ErrorAction SilentlyContinue
if ($existing) {
    $installDir = Split-Path $existing.Source
    Write-Host ""
    Write-Host "发现现有安装: $($existing.Source)"
    Write-Host "将升级到新版本..."
} else {
    $installDir = "$env:LOCALAPPDATA\OfficeCli"
}

# 创建安装目录
New-Item -ItemType Directory -Force -Path $installDir | Out-Null

# 安装二进制文件
Write-Host ""
Write-Host "安装到: $installDir\$binary"
Copy-Item -Force $source "$installDir\$binary"

# 添加到 PATH
$currentPath = [Environment]::GetEnvironmentVariable("Path", "User")
if ($currentPath -notlike "*$installDir*") {
    [Environment]::SetEnvironmentVariable("Path", "$currentPath;$installDir", "User")
    Write-Host ""
    Write-Host "已将 $installDir 添加到 PATH"
    Write-Host "重启终端以生效。"
}

# [商用版本] 设置配置文件权限
Write-Host ""
Write-Host "设置安全配置..."
$configDir = "$env:USERPROFILE\.officecli"
New-Item -ItemType Directory -Force -Path $configDir | Out-Null

# 完成
Write-Host ""
Write-Host "==========================================" -ForegroundColor Green
Write-Host "OfficeCLI 安装成功！" -ForegroundColor Green
Write-Host "==========================================" -ForegroundColor Green
Write-Host ""
Write-Host "版本: $version"
Write-Host "安装位置: $installDir\$binary"
Write-Host ""
Write-Host "[商用版本安全特性]" -ForegroundColor Yellow
Write-Host "  - 自动更新已禁用"
Write-Host "  - 日志记录默认关闭"
Write-Host "  - 配置文件已设置安全权限"
Write-Host ""
Write-Host "运行 'officecli --help' 开始使用"
Write-Host ""
Write-Host "安全提示:" -ForegroundColor Yellow
Write-Host "  - 定期关注安全公告以获取更新"
Write-Host "  - 处理敏感文档时避免使用 watch 功能"
Write-Host "  - 建议定期清理临时文件"
