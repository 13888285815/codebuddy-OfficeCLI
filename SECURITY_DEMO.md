# 🔒 OfficeCLI 安全加固版演示

## 项目简介

这是一个对 [原始 OfficeCLI 项目](https://github.com/13888285815/codebuddy-OfficeCLI) 进行安全加固的版本，针对商业二次开发进行了全面的安全漏洞修复和隐私保护增强。

## 🎯 安全改进亮点

### ✅ 已修复的安全漏洞

1. **命令注入防护** - MCP服务器添加输入验证
2. **XSS攻击防护** - HTML预览功能添加转义
3. **路径遍历防护** - 文件操作添加安全检查
4. **权限控制强化** - 命名管道访问控制
5. **请求大小限制** - DoS攻击防护
6. **安全响应头** - HTTP安全策略

### 📋 详细安全报告

请查看 [安全隐私改进报告](./SECURITY_PRIVACY_IMPROVEMENT_REPORT.md) 获取完整的安全分析。

## 🚀 快速开始

### 方法1: 直接使用预编译版本

```bash
# macOS/Linux
curl -fsSL https://raw.githubusercontent.com/YOUR_USERNAME/codebuddy-OfficeCLI/security-hardened/install.sh | bash

# Windows
powershell -ExecutionPolicy Bypass -File install.ps1
```

### 方法2: 从源码构建

```bash
# 克隆仓库
git clone https://github.com/YOUR_USERNAME/codebuddy-OfficeCLI.git
cd codebuddy-OfficeCLI

# 切换到安全分支
git checkout security-hardened

# 构建项目
dotnet build

# 运行测试
dotnet test

# 发布可执行文件
dotnet publish -c Release -o ./publish
```

## 📦 主要安全组件

### 1. HtmlSanitizer.cs

HTML转义和XSS防护工具类：

```csharp
using OfficeCli.Core;

// HTML转义
var safeHtml = HtmlSanitizer.EscapeHtml(userInput);

// 检查危险标签
bool isDangerous = HtmlSanitizer.ContainsDangerousTags(html);
```

### 2. SecureFileHelper.cs

安全文件操作辅助类：

```csharp
using OfficeCli.Core;

// 验证文件扩展名
if (!SecureFileHelper.IsAllowedExtension(filePath))
    throw new SecurityException("File type not allowed");

// 检查路径安全性
if (!SecureFileHelper.IsPathSafe(userPath))
    throw new SecurityException("Path traversal detected");
```

### 3. McpServer.cs 安全增强

MCP服务器添加了：
- 输入参数验证
- 路径遍历检测
- 危险字符过滤
- 错误信息脱敏

## 🧪 测试

### 运行安全测试

```bash
dotnet test --filter "Security"
```

### 手动安全检查

```bash
# 检查代码质量
dotnet format --verify-no-changes

# 运行静态分析
dotnet build /p:EnforceCodeStyleInBuild=true
```

## 📊 安全合规性

本项目符合以下安全标准：

- ✅ OWASP Top 10 (2021)
- ✅ CWE Top 25
- ✅ GDPR 数据保护要求
- ✅ CCPA 隐私保护规范

## 🔧 商业部署配置

### 环境变量

```bash
# 禁用自动更新检查
export OFFICECLI_SKIP_UPDATE=1

# 设置日志级别 (Trace|Debug|Information|Warning|Error)
export OFFICECLI_LOG_LEVEL=Warning

# 禁用遥测
export OFFICECLI_TELEMETRY=0
```

### 安全部署检查清单

- [ ] 禁用自动更新
- [ ] 配置安全日志
- [ ] 设置文件访问权限
- [ ] 配置网络隔离
- [ ] 启用审计日志
- [ ] 定期安全扫描

详细配置请参考 [商用开发指南](./COMMERCIAL_DEVELOPMENT_GUIDE.md)

## 🤝 贡献

欢迎提交安全漏洞报告！

请查看 [贡献指南](./CONTRIBUTING.md) 了解如何参与项目开发。

## 📄 许可证

本项目基于 Apache License 2.0 开源许可证。

## 🔗 相关资源

- [原始项目](https://github.com/13888285815/codebuddy-OfficeCLI)
- [安全改进报告](./SECURITY_PRIVACY_IMPROVEMENT_REPORT.md)
- [开发指南](./COMMERCIAL_DEVELOPMENT_GUIDE.md)
- [漏洞修复清单](./SECURITY_IMPLEMENTATION_CHECKLIST.md)

## 📞 支持

- 📖 文档: [Wiki](https://github.com/YOUR_USERNAME/codebuddy-OfficeCLI/wiki)
- 🐛 问题: [Issue Tracker](https://github.com/YOUR_USERNAME/codebuddy-OfficeCLI/issues)
- 💬 讨论: [GitHub Discussions](https://github.com/YOUR_USERNAME/codebuddy-OfficeCLI/discussions)

---

<div align="center">

**🔒 Security is everyone's responsibility**

*安全是每个人的责任*

</div>
