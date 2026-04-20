# OfficeCLI 商用安全版本

> **专为商用和二次开发优化的 OfficeCLI 版本**

这是 OfficeCLI 的商用安全版本，针对企业级使用场景进行了安全加固和隐私保护优化。

## 与原版的区别

### 安全改进

| 特性 | 原版 | 商用安全版 |
|------|------|-----------|
| 自动更新 | 启用 | **已禁用** |
| 网络下载 | 自动下载二进制文件 | **仅支持本地安装** |
| 日志记录 | 可选启用 | **默认禁用，增强脱敏** |
| 配置文件权限 | 默认 | **仅限当前用户访问** |
| Watch服务器超时 | 可配置 | **最大10分钟限制** |

### 隐私保护

- **无遥测**：不收集任何使用数据
- **无外部连接**：不会连接更新服务器或云服务
- **本地处理**：所有文档处理在本地完成
- **日志脱敏**：文件路径等敏感信息在日志中被脱敏

## 安装

### 前提条件

由于商用版本禁用了自动下载，您需要**手动下载**对应平台的二进制文件：

| 平台 | 文件名 |
|------|--------|
| macOS Apple Silicon | `officecli-mac-arm64` |
| macOS Intel | `officecli-mac-x64` |
| Linux x64 | `officecli-linux-x64` |
| Linux ARM64 | `officecli-linux-arm64` |
| Windows x64 | `officecli-win-x64.exe` |
| Windows ARM64 | `officecli-win-arm64.exe` |

### macOS / Linux

```bash
# 1. 下载对应平台的二进制文件（手动下载）
# 2. 运行安装脚本
chmod +x install.sh
./install.sh /path/to/officecli-binary
```

### Windows

```powershell
# 1. 下载对应平台的二进制文件（手动下载）
# 2. 运行安装脚本
.\install.ps1 -BinaryPath C:\path\to\officecli-win-x64.exe
```

## 安全特性详解

### 1. 禁用自动更新

商用版本完全移除了自动更新功能：
- 不会自动检查新版本
- 不会下载或执行外部代码
- 所有更新必须通过手动方式安装

### 2. 增强的日志隐私

```bash
# 启用日志（默认禁用）
officecli config log true
```

启用后，日志将：
- 仅记录文件名，不记录完整路径
- 脱敏敏感关键词（password, secret, key, token 等）
- 限制输出长度，防止日志过大

### 3. 配置文件安全

配置文件存储在 `~/.officecli/config.json`，具有以下安全特性：
- 仅限当前用户读写（权限 600）
- 不包含敏感信息
- 可随时删除重置

### 4. Watch 服务器限制

```bash
# 启动 watch 服务器
officecli watch document.pptx

# 建议：使用完毕后立即关闭
officecli unwatch document.pptx
```

安全限制：
- 仅绑定到 127.0.0.1（本地回环）
- 最大空闲时间限制为 10 分钟
- 不接受外部网络连接

## 商用授权

本版本基于 Apache 2.0 许可证，可用于商业用途：
- ✅ 企业内部使用
- ✅ 集成到商业产品中
- ✅ 二次开发和修改
- ✅ 分发和再许可

详细授权条款请参阅 [LICENSE](LICENSE) 文件。

## 安全指南

请参阅 [SECURITY.md](SECURITY.md) 了解：
- 安全使用建议
- 隐私注意事项
- 已知限制
- 安全更新流程

## 更新流程

由于自动更新已禁用，请按以下步骤手动更新：

1. 关注仓库的 Release 页面获取更新通知
2. 下载新版本的二进制文件
3. 运行安装脚本覆盖安装
4. 验证新版本：`officecli --version`

## 二次开发

### 构建

```bash
# 需要 .NET 10 SDK
dotnet publish -c Release -r <runtime-id> --self-contained -p:PublishSingleFile=true
```

### 安全修改建议

如果您需要进一步定制：
1. 审查所有网络请求代码
2. 检查文件操作权限
3. 验证输入数据的有效性
4. 定期审计依赖项

## 技术支持

商用版本的技术支持：
- GitHub Issues（公开问题）
- 安全漏洞请私下报告

## 致谢

本版本基于 [OfficeCLI](https://github.com/iOfficeAI/OfficeCLI) 项目，感谢原作者的开源贡献。

## 免责声明

本软件按"原样"提供，不提供任何明示或暗示的担保。详见 LICENSE 文件。
