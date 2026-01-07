# 企业微信文档转换服务部署指南

> ⚠️ **注意**: 本项目目前处于暂停状态。请先阅读 [README.md](README.md) 了解当前技术限制。

---

## 系统架构

```
┌───────────────────────────┐           ┌──────────────────────────┐
│        Linux VM           │           │        Windows VM        │
│    (前端接入 + 备份引擎)    │           │      (主转换引擎)         │
│                           │           │                          │
│  [企业微信] -> [Flask] ───┼── HTTP ──>│ [Python服务] -> [Office]  │
│              │            │           │                          │
│              ▼            │           └──────────────────────────┘
│        [LibreOffice]      │
│         (降级备份)         │
└───────────────────────────┘
```

---

## 第一部分：Windows VM 部署

### 1.1 环境要求
- Windows 10/11 或 Server 2022
- 2核 CPU, 4GB 内存
- Microsoft Office 或 WPS Office
- Python 3.11+

### 1.2 部署步骤

```powershell
cd C:\
git clone https://github.com/freeitaly/ios_better_printer.git converter
cd converter
pip install flask pywin32==306

# 开放防火墙端口
New-NetFirewallRule -DisplayName "OfficeConverter" -Direction Inbound -LocalPort 8080 -Protocol TCP -Action Allow

# 启动服务
python windows_converter_service.py
```

### 1.3 验证服务

```bash
curl http://localhost:8080/health
```

---

## 第二部分：Linux VM 部署

### 2.1 安装Docker

```bash
# 国内使用阿里云镜像
curl -fsSL https://mirrors.aliyun.com/docker-ce/linux/ubuntu/gpg | sudo gpg --dearmor -o /usr/share/keyrings/docker-archive-keyring.gpg
echo "deb [arch=$(dpkg --print-architecture) signed-by=/usr/share/keyrings/docker-archive-keyring.gpg] https://mirrors.aliyun.com/docker-ce/linux/ubuntu $(lsb_release -cs) stable" | sudo tee /etc/apt/sources.list.d/docker.list > /dev/null
sudo apt update
sudo apt install -y docker-ce docker-ce-cli containerd.io docker-compose-plugin
```

### 2.2 部署代码

```bash
sudo mkdir -p /opt/wecom-converter
cd /opt/wecom-converter
git clone https://github.com/freeitaly/ios_better_printer.git .
```

### 2.3 配置环境变量

```bash
cp .env.example .env
vim .env
```

配置示例：
```ini
WECOM_CORP_ID=你的企业ID
WECOM_AGENT_ID=你的应用AgentId
WECOM_SECRET=你的应用Secret
WECOM_TOKEN=自定义Token
WECOM_ENCODING_AES_KEY=43位随机字符串

WINDOWS_CONVERTER_ENABLED=true
WINDOWS_CONVERTER_URL=http://<Windows-VM-IP>:8080
```

### 2.4 启动服务

```bash
sudo docker compose up -d --build
```

---

## 第三部分：配置企业微信

### 3.1 创建自建应用
1. 访问 https://work.weixin.qq.com/
2. 管理后台 → 应用管理 → 创建应用
3. 记录 AgentId、Secret、企业ID

### 3.2 配置API接收
1. 应用设置 → 接收消息 → 设置API接收
2. 填写：
   - **URL**: `http://<公网IP>:18080/wecom`
   - **Token**: 与.env一致
   - **EncodingAESKey**: 与.env一致

---

## 已知问题

⚠️ **当前企业微信自建应用无法接收file类型消息回调**，导致核心功能无法使用。

详见 [README.md](README.md#项目状态说明)

---

## 维护命令

```bash
# 查看日志
sudo docker compose logs -f app

# 更新代码
git pull && sudo docker compose up -d --build

# 重启服务
sudo docker compose restart
```
