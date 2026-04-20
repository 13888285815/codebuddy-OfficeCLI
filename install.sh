#!/bin/bash
# [商用版本] OfficeCLI 安装脚本
# 安全改进：移除了自动下载功能，仅支持本地二进制文件安装

set -e

BINARY_NAME="officecli"
INSTALL_DIR="$HOME/.local/bin"

# 显示帮助信息
show_help() {
    echo "OfficeCLI 商用版本安装脚本"
    echo ""
    echo "用法:"
    echo "  ./install.sh [本地二进制文件路径]"
    echo ""
    echo "示例:"
    echo "  ./install.sh ./officecli-mac-arm64"
    echo "  ./install.sh ./officecli-linux-x64"
    echo ""
    echo "注意：商用版本不支持自动下载，请手动下载二进制文件后运行此脚本。"
}

# 检查参数
if [ "$1" = "--help" ] || [ "$1" = "-h" ]; then
    show_help
    exit 0
fi

# 检测平台
OS=$(uname -s | tr '[:upper:]' '[:lower:]')
ARCH=$(uname -m)

case "$OS" in
    darwin)
        case "$ARCH" in
            arm64) EXPECTED_ASSET="officecli-mac-arm64" ;;
            x86_64) EXPECTED_ASSET="officecli-mac-x64" ;;
            *) echo "不支持的架构: $ARCH"; exit 1 ;;
        esac
        ;;
    linux)
        case "$ARCH" in
            x86_64) EXPECTED_ASSET="officecli-linux-x64" ;;
            aarch64|arm64) EXPECTED_ASSET="officecli-linux-arm64" ;;
            *) echo "不支持的架构: $ARCH"; exit 1 ;;
        esac
        ;;
    *)
        echo "不支持的操作系统: $OS"
        echo "Windows 用户请使用 install.ps1 脚本"
        exit 1
        ;;
esac

echo "[商用版本] OfficeCLI 安装程序"
echo "平台: $OS / $ARCH"
echo "期望的二进制文件: $EXPECTED_ASSET"
echo ""

# 查找本地二进制文件
SOURCE=""

# 如果提供了参数，使用指定的文件
if [ -n "$1" ]; then
    if [ -f "$1" ]; then
        SOURCE="$1"
        echo "使用指定的二进制文件: $SOURCE"
    else
        echo "错误: 指定的文件不存在: $1"
        exit 1
    fi
else
    # 尝试在当前目录查找匹配的二进制文件
    for candidate in "./$EXPECTED_ASSET" "./$BINARY_NAME" "./bin/$EXPECTED_ASSET" "./bin/$BINARY_NAME" "./bin/release/$EXPECTED_ASSET" "./bin/release/$BINARY_NAME"; do
        if [ -f "$candidate" ]; then
            if [ ! -x "$candidate" ]; then
                chmod +x "$candidate"
            fi
            if "$candidate" --version >/dev/null 2>&1; then
                SOURCE="$candidate"
                echo "找到有效的二进制文件: $candidate"
                break
            fi
        fi
    done
fi

if [ -z "$SOURCE" ]; then
    echo "错误: 无法找到有效的 OfficeCLI 二进制文件。"
    echo ""
    echo "请手动下载对应平台的二进制文件:"
    echo "  - macOS Apple Silicon: officecli-mac-arm64"
    echo "  - macOS Intel: officecli-mac-x64"
    echo "  - Linux x64: officecli-linux-x64"
    echo "  - Linux ARM64: officecli-linux-arm64"
    echo ""
    echo "下载后运行: ./install.sh [二进制文件路径]"
    exit 1
fi

# 验证二进制文件
echo ""
echo "验证二进制文件..."
if ! "$SOURCE" --version >/dev/null 2>&1; then
    echo "错误: 二进制文件验证失败"
    exit 1
fi

VERSION=$("$SOURCE" --version 2>&1 | head -1)
echo "版本: $VERSION"

# 检查现有安装
EXISTING=$(command -v "$BINARY_NAME" 2>/dev/null || true)
if [ -n "$EXISTING" ]; then
    INSTALL_DIR=$(dirname "$EXISTING")
    echo ""
    echo "发现现有安装: $EXISTING"
    echo "将升级到新版本..."
fi

# 创建安装目录
mkdir -p "$INSTALL_DIR"

# 安装二进制文件
echo ""
echo "安装到: $INSTALL_DIR/$BINARY_NAME"
cp "$SOURCE" "$INSTALL_DIR/$BINARY_NAME.new"
chmod +x "$INSTALL_DIR/$BINARY_NAME.new"

# macOS: 移除隔离标志并临时签名
if [ "$(uname -s)" = "Darwin" ]; then
    xattr -d com.apple.quarantine "$INSTALL_DIR/$BINARY_NAME.new" 2>/dev/null || true
    codesign -s - -f "$INSTALL_DIR/$BINARY_NAME.new" 2>/dev/null || true
fi

# 原子替换
mv -f "$INSTALL_DIR/$BINARY_NAME.new" "$INSTALL_DIR/$BINARY_NAME"

# 添加到 PATH
case ":$PATH:" in
    *":$INSTALL_DIR:"*) ;;
    *)
        PATH_LINE="export PATH=\"$INSTALL_DIR:\$PATH\""
        if [ "$(uname -s)" = "Darwin" ]; then
            SHELL_RC="$HOME/.zshrc"
        elif [ -n "$ZSH_VERSION" ]; then
            SHELL_RC="$HOME/.zshrc"
        else
            SHELL_RC="$HOME/.bashrc"
        fi
        if ! grep -qF "$INSTALL_DIR" "$SHELL_RC" 2>/dev/null; then
            echo "" >> "$SHELL_RC"
            echo "$PATH_LINE" >> "$SHELL_RC"
            echo ""
            echo "已将 $INSTALL_DIR 添加到 PATH ($SHELL_RC)"
            echo "运行 'source $SHELL_RC' 或重启终端以生效。"
        fi
        ;;
esac

# [商用版本] 设置配置文件权限
echo ""
echo "设置安全配置..."
CONFIG_DIR="$HOME/.officecli"
mkdir -p "$CONFIG_DIR"
if [ -f "$CONFIG_DIR/config.json" ]; then
    chmod 600 "$CONFIG_DIR/config.json" 2>/dev/null || true
fi

# 完成
echo ""
echo "=========================================="
echo "OfficeCLI 安装成功！"
echo "=========================================="
echo ""
echo "版本: $VERSION"
echo "安装位置: $INSTALL_DIR/$BINARY_NAME"
echo ""
echo "[商用版本安全特性]"
echo "  - 自动更新已禁用"
echo "  - 日志记录默认关闭"
echo "  - 配置文件已设置安全权限"
echo ""
echo "运行 'officecli --help' 开始使用"
echo ""
echo "安全提示:"
echo "  - 定期关注安全公告以获取更新"
echo "  - 处理敏感文档时避免使用 watch 功能"
echo "  - 建议定期清理临时文件"
