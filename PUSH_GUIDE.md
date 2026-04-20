# 🚀 推送到GitHub指南

## 快速开始

### 方式1: 使用GitHub CLI（推荐）

如果你已安装 GitHub CLI：

```bash
cd "/Users/zzx/Desktop/ai源码/trae-codingbuddy/codebuddy-OfficeCLI"

# 推送安全分支到你的仓库
git push -u origin security-hardened-demo

# 推送security分支
git push -u origin security-hardened
```

### 方式2: 使用HTTPS认证

```bash
cd "/Users/zzx/Desktop/ai源码/trae-codingbuddy/codebuddy-OfficeCLI"

# 设置远程仓库
git remote set-url origin https://github.com/13888285815/codebuddy-OfficeCLI.git

# 推送分支
git push -u origin security-hardened-demo
git push -u origin security-hardened
```

### 方式3: 使用SSH密钥

```bash
# 1. 检查SSH密钥
ls -la ~/.ssh/

# 2. 如果没有，创建SSH密钥
ssh-keygen -t ed25519 -C "your_email@example.com"

# 3. 添加SSH密钥到GitHub账户
# 访问: https://github.com/settings/keys
# 点击 "New SSH key"，粘贴 ~/.ssh/id_ed25519.pub 内容

# 4. 切换到SSH远程仓库
git remote set-url origin git@github.com:13888285815/codebuddy-OfficeCLI.git

# 5. 推送
git push -u origin security-hardened-demo
git push -u origin security-hardened
```

---

## 🎯 GitHub 演示设置

### 1. 创建演示分支（推荐）

```bash
# 在GitHub网站上操作：
# 1. 访问 https://github.com/13888285815/codebuddy-OfficeCLI
# 2. 点击 "Branch: main" 下拉菜单
# 3. 输入 "security-hardened-demo"
# 4. 点击 "Create branch: security-hardened-demo"
```

### 2. 启用GitHub Pages

```bash
# 在GitHub网站上操作：
# 1. 进入仓库 Settings
# 2. 左侧菜单点击 "Pages"
# 3. Source 选择 "main" 分支
# 4. 点击 Save

# 演示地址格式：
# https://13888285815.github.io/codebuddy-OfficeCLI/
```

### 3. 创建Releases

```bash
# 使用GitHub CLI
gh release create v1.0.0 \
  --title "🔒 Security Hardened Version" \
  --notes "安全加固版本，包含以下改进：
  - 命令注入防护
  - XSS攻击防护
  - 路径遍历检测
  - 完整安全文档" \
  --target security-hardened-demo
```

---

## 🌐 GitHub 演示链接格式

推送成功后，你将拥有以下演示链接：

### 1. 主仓库地址
```
https://github.com/13888285815/codebuddy-OfficeCLI
```

### 2. 安全分支
```
https://github.com/13888285815/codebuddy-OfficeCLI/tree/security-hardened-demo
```

### 3. 安全文档
```
https://github.com/13888285815/codebuddy-OfficeCLI/blob/security-hardened-demo/SECURITY_README.md
```

### 4. 安全组件
```
https://github.com/13888285815/codebuddy-OfficeCLI/blob/security-hardened-demo/src/officecli/Core/HtmlSanitizer.cs
```

### 5. CI/CD工作流
```
https://github.com/13888285815/codebuddy-OfficeCLI/actions/workflows/security-build.yml
```

---

## 💡 分享技巧

### 创建可分享的链接

1. **固定安全分支**:
   ```
   https://github.com/13888285815/codebuddy-OfficeCLI/tree/security-hardened-demo
   ```

2. **直接查看文件**:
   ```
   https://github.com/13888285815/codebuddy-OfficeCLI/blob/security-hardened-demo/SECURITY_README.md
   ```

3. **比较分支差异**:
   ```
   https://github.com/13888285815/codebuddy-OfficeCLI/compare/main...security-hardened-demo
   ```

4. **下载源代码**:
   ```
   https://github.com/13888285815/codebuddy-OfficeCLI/archive/refs/heads/security-hardened-demo.zip
   ```

---

## 🎉 完成后的分享模板

你可以使用以下模板分享：

---

**🔒 OfficeCLI 安全加固版已发布！**

**GitHub 仓库**: https://github.com/13888285815/codebuddy-OfficeCLI

**演示分支**: `security-hardened-demo`

**主要安全改进**:
✅ 命令注入防护
✅ XSS攻击防护
✅ 路径遍历检测
✅ 完整的合规文档

**快速开始**:
```bash
git clone -b security-hardened-demo https://github.com/13888285815/codebuddy-OfficeCLI.git
cd codebuddy-OfficeCLI
./build.sh
```

**文档**: 查看 `SECURITY_README.md` 了解更多！

---

