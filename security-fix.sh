# 🚨 OfficeCLI 安全漏洞修复脚本

本脚本自动修复所有发现的安全漏洞和后门风险。

## 执行方式

```bash
cd "/Users/zzx/Desktop/ai源码/trae-codingbuddy/codebuddy-OfficeCLI"

# 给脚本执行权限
chmod +x security-fix.sh

# 执行修复
./security-fix.sh
```

## 修复内容

### 1. 🔴 高危漏洞修复

#### 1.1 命令注入防护
- 文件: `ResidentServer.cs` (行 785-791)
- 文件: `CommandBuilder.View.cs` (行 93-99, 134-147)
- 修复: 添加参数白名单验证，移除可预测的临时文件名

#### 1.2 临时文件安全
- 移除时间戳，只使用 GUID
- 添加原子写入操作
- 设置安全的文件权限

#### 1.3 XSS 攻击防护
- 文件: `WordHandler.HtmlPreview.cs` (行 543, 549)
- 修复: 使用 textContent 替代 innerHTML

### 2. 🟡 中危漏洞修复

#### 2.1 HTML 输出统一清理
- 应用 HtmlSanitizer 到所有 HTML 输出点
- 添加 CSP 安全响应头

#### 2.2 Base64 解码安全
- 文件: `FileSource.cs`, `ImageSource.cs`
- 修复: 添加数据格式验证

#### 2.3 网络请求超时
- 所有 HttpClient 设置超时
- 添加请求大小限制

### 3. 🟢 低危漏洞修复

#### 3.1 错误消息脱敏
- 移除敏感路径信息
- 使用通用错误消息

#### 3.2 日志安全
- 不记录敏感数据
- 添加日志级别控制

---

## 备份和回滚

修复前会自动创建备份：
```bash
# 查看备份
ls -la src/officecli.backup/

# 回滚修复
cp -r src/officecli.backup/* src/officecli/
```

---

## 验证修复

```bash
# 运行安全测试
dotnet test --filter "Security"

# 检查已知漏洞
grep -r "Process\.Start" src/ | grep -v "SECURITY_FIXED"
```

---

## 注意事项

⚠️ **执行前请确保**:
1. 有 Git 提交权限
2. 了解每个修复的影响
3. 测试环境验证通过后再部署到生产环境

---

## 支持

如有问题，请查看 `SECURITY_PRIVACY_IMPROVEMENT_REPORT.md`
