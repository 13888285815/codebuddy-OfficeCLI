#!/bin/bash
# ============================================================
# OfficeCLI 一键推送到 GitHub
# ============================================================

echo "🚀 开始推送到 GitHub..."

cd "/Users/zzx/Desktop/ai源码/trae-codingbuddy/codebuddy-OfficeCLI"

# 检查远程仓库配置
echo "📡 检查远程仓库..."
git remote -v

# 推送到 GitHub
echo "📤 推送 security-hardened-demo 分支..."
git push -u origin security-hardened-demo

# 推送 security-hardened 分支（包含所有文档）
echo "📤 推送 security-hardened 分支..."
git push -u origin security-hardened

echo ""
echo "✅ 推送完成！"
echo ""
echo "🔗 访问以下链接查看安全版本："
echo "   • 仓库主页: https://github.com/13888285815/codebuddy-OfficeCLI"
echo "   • 安全分支: https://github.com/13888285815/codebuddy-OfficeCLI/tree/security-hardened-demo"
echo "   • 安全文档: https://github.com/13888285815/codebuddy-OfficeCLI/blob/security-hardened-demo/SECURITY_FIXES_APPLIED.md"
echo "   • 完整报告: https://github.com/13888285815/codebuddy-OfficeCLI/blob/security-hardened-demo/SECURITY_PRIVACY_IMPROVEMENT_REPORT.md"
echo ""
echo "📋 下一步："
echo "   1. 访问 GitHub Actions: https://github.com/13888285815/codebuddy-OfficeCLI/actions"
echo "   2. 创建 Release（可选）"
echo "   3. 分享链接给朋友！"
echo ""
