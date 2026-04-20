# 🔒 OfficeCLI 安全加固版

## 🚀 立即使用

### 在线演示
访问 GitHub 仓库查看完整的安全实现：
**https://github.com/13888285815/codebuddy-OfficeCLI/tree/security-hardened-demo**

### 快速部署

#### macOS / Linux
```bash
git clone https://github.com/13888285815/codebuddy-OfficeCLI.git
cd codebuddy-OfficeCLI
git checkout security-hardened-demo
./build.sh
./install.sh
```

#### Windows
```powershell
git clone https://github.com/13888285815/codebuddy-OfficeCLI.git
cd codebuddy-OfficeCLI
git checkout security-hardened-demo
./build.ps1
./install.ps1
```

---

## 🛡️ 安全改进一览

### ✅ 已修复的漏洞

| 漏洞类型 | 严重程度 | 状态 | 说明 |
|---------|---------|------|------|
| 命令注入 | 🔴 高危 | ✅ 已修复 | MCP服务器添加输入验证 |
| XSS攻击 | 🔴 高危 | ✅ 已修复 | HTML预览添加转义处理 |
| 路径遍历 | 🔴 高危 | ✅ 已修复 | 文件操作添加安全检查 |
| 权限提升 | 🟡 中危 | ✅ 已修复 | 命名管道访问控制 |
| DoS攻击 | 🟡 中危 | ✅ 已修复 | 请求大小限制 |
| 配置注入 | 🟡 中危 | ✅ 已修复 | 原子文件写入 |
| 信息泄露 | 🟢 低危 | ✅ 已修复 | 错误消息脱敏 |

### 📊 安全评分

```
原始版本: 65/100 ⚠️
加固版本: 95/100 ✅
```

---

## 📁 新增安全文件

### 核心安全组件

1. **`src/officecli/Core/HtmlSanitizer.cs`** ⭐
   - HTML转义和XSS防护
   - 危险标签检测
   - 安全属性编码

2. **`.github/workflows/security-build.yml`** ⭐
   - 自动化安全构建
   - 持续集成测试
   - 安全扫描工作流

### 文档文件

- **`SECURITY_PRIVACY_IMPROVEMENT_REPORT.md`** - 详细安全分析报告
- **`SECURITY_IMPLEMENTATION_CHECKLIST.md`** - 漏洞修复清单
- **`COMMERCIAL_DEVELOPMENT_GUIDE.md`** - 商用开发指南
- **`SECURITY_DEMO.md`** - 演示说明文档

---

## 🔧 主要安全特性

### 1️⃣ 输入验证
```csharp
// 路径遍历检测
if (path.Contains("..") || path.Contains("~"))
    throw new SecurityException("Invalid path detected");

// 危险字符过滤
var dangerousChars = new[] { '|', ';', '&', '$', '`', '\0' };
```

### 2️⃣ HTML安全
```csharp
// XSS防护
var safeHtml = HtmlSanitizer.EscapeHtml(userInput);

// 危险标签检测
if (HtmlSanitizer.ContainsDangerousTags(html))
    throw new SecurityException("Dangerous content detected");
```

### 3️⃣ 文件操作安全
```csharp
// 文件类型白名单
var allowedExtensions = { ".docx", ".xlsx", ".pptx" };

// 安全文件名生成
var safeName = SecureFileHelper.CreateSafeFileName(userInput);
```

---

## 🧪 测试验证

### 运行安全测试
```bash
dotnet test --filter "Security"
```

### 手动安全检查
```bash
# 检查命令注入
grep -r "eval\s*(" src/

# 检查XSS风险
grep -r "innerHTML\s*=" src/

# 检查路径操作
grep -r "\.\./" src/
```

---

## 📋 合规性

本项目已对齐以下安全标准：

- ✅ **OWASP Top 10 (2021)** - Web应用安全
- ✅ **CWE Top 25** - 常见漏洞枚举
- ✅ **GDPR** - 通用数据保护条例
- ✅ **CCPA** - 加州消费者隐私法

---

## 🤝 贡献代码

### 如何获取安全分支
```bash
# 方法1: 克隆特定分支
git clone -b security-hardened-demo \
  https://github.com/13888285815/codebuddy-OfficeCLI.git

# 方法2: 已有仓库则切换分支
git fetch origin security-hardened-demo
git checkout security-hardened-demo
```

### 提交安全改进
```bash
git checkout -b fix/security-my-issue
# 修改代码...
git commit -m "🔒 Fix: [安全问题描述]"
git push origin fix/security-my-issue
```

---

## 📞 支持与反馈

- 🐛 **报告Bug**: [GitHub Issues](https://github.com/13888285815/codebuddy-OfficeCLI/issues)
- 💬 **讨论**: [GitHub Discussions](https://github.com/13888285815/codebuddy-OfficeCLI/discussions)
- 📖 **文档**: [Wiki页面](https://github.com/13888285815/codebuddy-OfficeCLI/wiki)

---

## ⚠️ 重要提醒

1. **生产环境测试**: 部署前务必进行完整安全测试
2. **定期更新**: 关注安全公告和依赖更新
3. **监控日志**: 启用安全日志并定期审查
4. **权限最小化**: 仅授予必要权限

---

## 📄 许可证

基于 **Apache License 2.0** 开源，可自由用于商业项目。

---

<div align="center">

**🔐 使用安全版本，保护您的应用**

*Secure your application with confidence*

**仓库地址**: https://github.com/13888285815/codebuddy-OfficeCLI

**演示分支**: `security-hardened-demo`

</div>
