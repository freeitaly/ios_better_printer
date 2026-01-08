# 文档转换服务部署指南

本指南介绍如何部署 iOS 作业排版助手的服务端。

---

## 系统架构

```
┌─────────────────────────────────────────────────────────────────┐
│  iPhone                                                         │
│  [快捷指令] ────────── HTTP POST ──────────────────────────────>│
└─────────────────────────────────────────────────────────────────┘
                              │
                              ▼
┌─────────────────────────────────────────────────────────────────┐
│  Linux Server (Docker)                                          │
│                                                                 │
│  [Nginx :18080] ──> [Flask /api/convert] ──> [Windows转换服务]   │
│                              │                                  │
│                              ▼                                  │
│                     [LibreOffice 备用]                          │
└─────────────────────────────────────────────────────────────────┘
                              │
                              ▼
┌─────────────────────────────────────────────────────────────────┐
│  Windows VM                                                     │
│  [Python服务 :8080] ──> [Microsoft Office / WPS COM]            │
└─────────────────────────────────────────────────────────────────┘
```

### 端口说明

| 端口 | 用途 |
|-----|------|
| **18080** | 对外服务端口（Nginx） |
| 5000 | Flask 内部端口（容器内） |
| 8080 | Windows 转换服务端口 |

> ⚠️ **为什么使用 18080？**
> 
> 电信专线独立 IP 通常会封禁 80 和 8080 端口。部署时发现这些端口无法访问，因此改用 18080 端口。
> 
> 如需使用其他端口，修改 `docker-compose.yml` 中的端口映射：
> ```yaml
> ports:
>   - "你的端口:80"
> ```

---

## 第一部分：Windows VM 部署（推荐）

### 1.1 环境要求
- Windows 10/11 或 Server 2022
- Microsoft Office 或 WPS Office
- Python 3.11+

### 1.2 部署步骤

```powershell
cd C:\
git clone https://github.com/freeitaly/ios_better_printer.git converter
cd converter
pip install flask pywin32==306

# 开放防火墙
New-NetFirewallRule -DisplayName "OfficeConverter" -Direction Inbound -LocalPort 8080 -Protocol TCP -Action Allow

# 启动服务
python windows_converter_service.py
```

### 1.3 验证

```bash
curl http://localhost:8080/health
```

---

## 第二部分：Linux VM 部署

### 2.1 安装 Docker

```bash
# 国内镜像
curl -fsSL https://mirrors.aliyun.com/docker-ce/linux/ubuntu/gpg | sudo gpg --dearmor -o /usr/share/keyrings/docker-archive-keyring.gpg
echo "deb [arch=$(dpkg --print-architecture) signed-by=/usr/share/keyrings/docker-archive-keyring.gpg] https://mirrors.aliyun.com/docker-ce/linux/ubuntu $(lsb_release -cs) stable" | sudo tee /etc/apt/sources.list.d/docker.list > /dev/null
sudo apt update && sudo apt install -y docker-ce docker-ce-cli containerd.io docker-compose-plugin
```

### 2.2 部署

```bash
sudo mkdir -p /opt/ios_better_printer
cd /opt/ios_better_printer
git clone https://github.com/freeitaly/ios_better_printer.git .
cp .env.example .env
vim .env
```

### 2.3 配置 .env

```ini
# Windows 转换服务
WINDOWS_CONVERTER_ENABLED=true
WINDOWS_CONVERTER_URL=http://<Windows-VM-IP>:8080

# 企业微信（可选，如果不使用可留空）
WECOM_CORP_ID=
WECOM_AGENT_ID=
WECOM_SECRET=
WECOM_TOKEN=
WECOM_ENCODING_AES_KEY=
```

### 2.4 启动

```bash
sudo docker compose up -d --build
```

---

## 第三部分：验证部署

### 测试转换 API

```bash
# 上传 Word 文件测试
curl -X POST -F "file=@test.docx" http://localhost:18080/api/convert -o output.pdf
file output.pdf  # 应显示: PDF document
```

### 检查日志

```bash
sudo docker compose logs -f app
```

---

## 维护命令

```bash
# 查看日志
sudo docker compose logs -f app

# 更新代码
git pull && sudo docker compose up -d --build

# 重启
sudo docker compose restart
```

---

## 客户端配置

服务端部署完成后，在 iPhone 上创建快捷指令：

详见 [IOS_SHORTCUT_GUIDE.md](IOS_SHORTCUT_GUIDE.md)
