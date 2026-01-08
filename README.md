# iOS 作业排版助手

> 📄 将老师发送的 Word/Excel 作业文件完美转换为 PDF，解决 iPhone 打印排版错乱问题

## ✅ 项目状态：可用

通过 **iOS 快捷指令** 实现文档转换，无需企业认证，无需域名备案。

---

## 🚀 快速开始

### 一键安装快捷指令

1. 按照 [IOS_SHORTCUT_GUIDE.md](IOS_SHORTCUT_GUIDE.md) 创建快捷指令
2. 将 URL 设为你的服务器地址：`http://<服务器IP>:18080/api/convert`

### 使用方法

适用于 **任何支持共享的 iOS App**（微信、文件、邮件、网盘等）：

1. **长按**需要转换的 Word/Excel/PPT 文件
2. 点击 **"共享"** 或 **"用其他应用打开"**
3. 选择 **"PDF转换打印"**
4. 等待 5-15 秒，PDF 自动打开
5. 点击分享 → **打印** / **保存** / **AirDrop**

---

## 📱 功能特点

- ✅ **100% 保真**：使用 Windows Office 引擎转换，排版完美还原
- ✅ **无需认证**：不需要企业微信认证或域名备案
- ✅ **原生体验**：通过 iOS 共享菜单调用，无需打开额外 App
- ✅ **支持格式**：Word (.doc/.docx)、Excel (.xls/.xlsx)、PPT (.ppt/.pptx)
- ✅ **文件大小**：最大支持 100MB

---

## 🏗️ 系统架构

```
iPhone                          Linux Server                     Windows VM
  │                                  │                               │
  │  [共享菜单]                       │                               │
  │      │                           │                               │
  │      ▼                           │                               │
  │  [快捷指令] ──HTTP POST──────────>│  [Nginx] ──────────────────────>│
  │                                  │     │                          │
  │                                  │     ▼                          │
  │                                  │  [Flask API]                   │
  │                                  │     │                          │
  │                                  │     ▼                          │
  │  <──────────── PDF ──────────────│  [文档转换] <── Office COM ────│
  │                                  │                               │
  │  [快速查看/打印]                  │                               │
```

---

## 🛠️ 服务端部署

详见 [DEPLOYMENT.md](DEPLOYMENT.md)

### 快速部署

```bash
# 克隆项目
git clone https://github.com/freeitaly/ios_better_printer.git
cd ios_better_printer

# 配置环境变量
cp .env.example .env
vim .env

# 启动服务
sudo docker compose up -d --build
```

---

## 📁 项目结构

```
ios_better_printer/
├── app.py                    # Flask API (/api/convert)
├── converter.py              # 文档转换引擎
├── wecom_api.py              # 企业微信API（保留兼容）
├── windows_converter_service.py  # Windows Office 转换服务
├── nginx.conf                # Nginx 配置
├── docker-compose.yml        # Docker 编排
├── IOS_SHORTCUT_GUIDE.md     # iOS 快捷指令设置指南
├── DEPLOYMENT.md             # 部署指南
└── TROUBLESHOOTING.md        # 故障排查指南
```

---

## 📖 文档

- [iOS 快捷指令设置指南](IOS_SHORTCUT_GUIDE.md)
- [服务端部署指南](DEPLOYMENT.md)
- [开发历程存档](DEVELOPMENT_HISTORY.md)

---

## 🔄 开发历程

本项目经历了多次技术方案迭代：

| 阶段 | 方案 | 结果 |
|-----|------|------|
| 1 | 微信公众号 | ❌ 不支持接收文件消息 |
| 2 | 企业微信自建应用 | ❌ 需要企业认证+域名备案 |
| 3 | **iOS 快捷指令** | ✅ **成功** |

详细历程见 [DEVELOPMENT_HISTORY.md](DEVELOPMENT_HISTORY.md)

---

## 许可证

MIT License
