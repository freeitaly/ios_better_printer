# 企业微信文档转换服务部署指南 (双引擎版)

本文档提供完整的生产环境部署步骤。

---

## 🏗️ 系统架构

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

## 🛠️ 部署流程

1. 部署Windows VM (可选，用于高保真转换)
2. 部署Linux VM
3. 注册企业微信并配置应用

---

## 第一部分：Windows VM 部署 (可选)

### 1.1 创建虚拟机
- **OS**: Windows 10/11 或 Server 2022
- **配置**: 2核 CPU, 4GB 内存

### 1.2 安装必要软件
1. Microsoft Office 或 WPS Office
2. Python 3.11+
3. Git

### 1.3 部署转换服务

```powershell
cd C:\
git clone https://github.com/freeitaly/ios_better_printer.git converter
cd converter
pip install flask pywin32==306
New-NetFirewallRule -DisplayName "OfficeConverter" -Direction Inbound -LocalPort 8080 -Protocol TCP -Action Allow
python windows_converter_service.py
```

---

## 第二部分：Linux VM 部署

### 2.1 环境准备 (Ubuntu 22.04)

```bash
# 安装Docker (国内使用阿里云镜像)
curl -fsSL https://mirrors.aliyun.com/docker-ce/linux/ubuntu/gpg | sudo gpg --dearmor -o /usr/share/keyrings/docker-archive-keyring.gpg
echo "deb [arch=$(dpkg --print-architecture) signed-by=/usr/share/keyrings/docker-archive-keyring.gpg] https://mirrors.aliyun.com/docker-ce/linux/ubuntu $(lsb_release -cs) stable" | sudo tee /etc/apt/sources.list.d/docker.list > /dev/null
sudo apt update
sudo apt install -y docker-ce docker-ce-cli containerd.io docker-compose-plugin
sudo systemctl enable docker
```

### 2.2 部署代码

```bash
sudo mkdir -p /opt/wecom-converter
cd /opt/wecom-converter
git clone https://github.com/freeitaly/ios_better_printer.git .
```

### 2.3 配置文件

```bash
cp .env.example .env
vim .env
```

**关键配置**：

```ini
# 企业微信配置
WECOM_CORP_ID=你的企业ID
WECOM_AGENT_ID=你的应用AgentId
WECOM_SECRET=你的应用Secret
WECOM_TOKEN=自定义Token
WECOM_ENCODING_AES_KEY=43位随机字符串

# Windows转换服务
WINDOWS_CONVERTER_ENABLED=true
WINDOWS_CONVERTER_URL=http://<Windows-VM-IP>:8080
```

### 2.4 启动服务

```bash
sudo docker compose up -d --build
```

> ⚠️ 服务使用 **18080** 端口，需在路由器配置端口转发。

---

## 第三部分：配置企业微信

### 3.1 注册企业微信
1. 访问 https://work.weixin.qq.com/
2. 个人可选择"其他组织"类型注册

### 3.2 创建自建应用
1. 管理后台 → 应用管理 → 创建应用
2. 记录 **AgentId** 和 **Secret**
3. 在"我的企业"页面记录 **企业ID**

### 3.3 配置接收消息
1. 应用设置 → 接收消息 → 设置API接收
2. 填写：
   - **URL**: `http://<公网IP>:18080/wecom`
   - **Token**: 与.env中WECOM_TOKEN一致
   - **EncodingAESKey**: 与.env中WECOM_ENCODING_AES_KEY一致
3. 点击保存

### 3.4 测试
在企业微信中打开应用，发送一个Word文件，等待返回PDF。

---

## 故障排查

| 现象 | 解决方案 |
|------|----------|
| 回调URL验证失败 | 检查Token和EncodingAESKey是否一致 |
| 转换使用LibreOffice | 检查WINDOWS_CONVERTER_URL配置 |
| 外网无法访问 | 检查路由器18080端口转发 |

## 维护

```bash
# 查看日志
sudo docker compose logs -f app

# 更新代码
git pull && sudo docker compose up -d --build
```
