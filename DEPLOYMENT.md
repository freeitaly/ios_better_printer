# 微信文档转换服务部署指南 (双引擎版)

本文档提供完整的生产环境部署步骤。本系统采用**双引擎架构**，结合了Linux容器的高效性和Windows Office的高保真转换能力。

---

## 🏗️ 系统架构

为了实现100%的排版还原度（特别是页数一致性），系统由两部分组成：

```
┌───────────────────────────┐           ┌──────────────────────────┐
│        Linux VM           │           │        Windows VM        │
│    (前端接入 + 备份引擎)    │           │      (主转换引擎)         │
│                           │           │                          │
│  [微信] -> [Flask App] ───┼── HTTP ──>│ [Python服务] -> [Office]  │
│              │            │           │                          │
│              ▼            │           └──────────────────────────┘
│        [LibreOffice]      │
│         (降级备份)         │
└───────────────────────────┘
```

- **Linux VM**: 运行Docker容器，处理微信消息，负责与用户交互。
- **Windows VM**: 运行Office COM服务，提供原汁原味的文档转换。

> **注意**：如果不部署Windows VM，系统将自动降级使用Linux内置的LibreOffice进行转换（可能会有排版/页数差异）。

---

## 🛠️ 部署流程概览

1. **部署Windows VM** (作为上游服务)
2. **部署Linux VM** (作为下游应用)
3. **联调测试**

---

## 第一部分：Windows VM 部署 (主转换引擎)

此步骤在ESXi上创建一个Windows虚拟机，用于运行Microsoft Office。

### 1.1 创建虚拟机
- **OS**: Windows 10/11 (推荐) 或 Server 2022
- **配置**: 2核 CPU, 4GB 内存, 60GB 硬盘
- **网络**: 需固定IP (例如: `192.168.1.101`)

### 1.2 安装必要软件
1. **Microsoft Office**: 安装Word, Excel, PowerPoint组件。
2. **Python 3.11+**: [下载安装](https://www.python.org/downloads/windows/) (务必勾选 "Add to PATH")。
3. **Git**: [下载安装](https://git-scm.com/download/win)。

### 1.3 部署转换服务

打开PowerShell执行：

```powershell
# 1. 下载代码
cd C:\
git clone https://github.com/your-repo/ios_better_printer.git converter
cd converter

# 2. 安装依赖
pip install flask pywin32==306
# 如果报错，执行: python C:\Python311\Scripts\pywin32_postinstall.py -install

# 3. 开放防火墙端口 8080
New-NetFirewallRule -DisplayName "OfficeConverter" -Direction Inbound -LocalPort 8080 -Protocol TCP -Action Allow

# 4. 启动服务测试
python windows_converter_service.py
```

### 1.4 配置开机自启 (推荐使用NSSM)
1. 下载 [NSSM](https://nssm.cc/release/nssm-2.24.zip) 并解压到 `C:\nssm`.
2. 管理员运行CMD:
   ```cmd
   C:\nssm\win64\nssm.exe install OfficeConverter "C:\Python311\python.exe" "C:\converter\windows_converter_service.py"
   C:\nssm\win64\nssm.exe set OfficeConverter AppDirectory "C:\converter"
   C:\nssm\win64\nssm.exe start OfficeConverter
   ```

---

## 第二部分：Linux VM 部署 (应用主程序)

此步骤部署处理微信消息的主程序。

### 2.1 环境准备 (Ubuntu 22.04)

```bash
# 安装Docker
curl -fsSL https://get.docker.com | sudo sh
sudo systemctl enable docker

# 安装Docker Compose
sudo apt install docker-compose -y
```

### 2.2 部署代码

```bash
# 1. 创建目录
sudo mkdir -p /opt/wechat-converter
cd /opt/wechat-converter

# 2. 上传项目代码 (除windows_*.py外的所有文件)
# ...使用git clone或scp...
```

### 2.3 配置文件

复制并编辑配置文件：

```bash
cp .env.example .env
vim .env
```

**关键配置项**：

```ini
# 微信配置
WECHAT_APP_ID=你的AppID
WECHAT_APP_SECRET=你的AppSecret
WECHAT_TOKEN=你的Token

# 🔌 连接到Windows VM (关键步骤)
WINDOWS_CONVERTER_ENABLED=true
WINDOWS_CONVERTER_URL=http://192.168.1.101:8080
WINDOWS_CONVERTER_TIMEOUT=60
```
> 将 `192.168.1.101` 替换为你实际的Windows VM IP。

### 2.4 启动服务

```bash
sudo docker-compose up -d --build
```

---

## 第三部分：联调与验证

### 3.1 检查连接

在Linux VM内测试能否通过内网访问Windows服务：

```bash
sudo docker exec -it wechat-doc-converter bash
curl http://192.168.1.101:8080/health
# 应返回 {"status": "ok", ...}
```

### 3.2 微信端测试

1. 关注测试号/订阅号。
2. 发送一个Word文档。
3. 观察Linux容器日志：
   ```bash
   sudo docker-compose logs -f app
   ```
4. 成功标志：
   - 日志显示 `尝试使用Windows转换服务...`
   - 日志显示 `✅ Windows转换成功`
   - 手机收到PDF，且**页数与Windows上一致**。

---

## 故障排查

| 现象 | 可能原因 | 解决方案 |
|------|----------|----------|
| **转换显示"使用LibreOffice转换"** | Linux无法连接Windows服务 | 检查Windows防火墙8080端口；检查.env中的IP配置 |
| **Windows服务报错 0x80080005** | Office未激活或卡死 | 登录Windows VM打开Word确认无弹窗；重启Windows服务 |
| **一直显示"处理中"无结果** | 转换超时 | 在.env中增加 `WINDOWS_CONVERTER_TIMEOUT` 到 120 |
| **手机未收到任何回复** | 微信回调配置错误 | 检查微信后台URL配置；检查Linux防火墙80端口 |

## 维护

- **日志查看**: `sudo docker-compose logs -f --tail=100`
- **临时文件清理**: 系统已配置自动清理，无需手动干预。
- **更新**: `git pull && sudo docker-compose up -d --build`
