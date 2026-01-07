# 企业微信作业排版助手

> 📄 将老师发送的Word/Excel作业文件完美转换为PDF，解决iPhone打印排版错乱问题

## 项目简介

这是一个基于**企业微信**的Office文档转PDF转换服务。系统采用**双引擎架构**（Windows Office + LibreOffice），确保转换后的PDF排版与Windows环境**100%一致**。

### 核心功能

- ✅ **企业微信内闭环**：直接在企业微信中发送文件即可转换
- ✅ **100%保真**：支持调用Windows Office引擎，排版完美还原
- ✅ **高可用**：主要引擎故障时自动降级到LibreOffice
- ✅ **即转即用**：5-15秒快速转换，支持Word/Excel/PPT

### 适用场景

- 家长接收老师发送的作业文件，需远程打印
- iPhone用户需要保证文档排版与Windows一致

## 快速开始

### 前置要求

- Linux服务器（Ubuntu 22.04推荐）
- Docker 和 Docker Compose
- **Windows虚拟机**（可选，用于实现100%排版还原）
- 固定公网IP地址
- **企业微信账号**（个人可免费注册）

### 一、注册企业微信

1. 访问：https://work.weixin.qq.com/
2. 注册企业微信（个人可选择"其他组织"类型）
3. 进入管理后台 → 应用管理 → 创建自建应用
4. 记录以下信息：
   - 企业ID (`corpid`)
   - 应用AgentId
   - 应用Secret

### 二、部署服务

#### 1. 克隆项目

```bash
git clone https://github.com/freeitaly/ios_better_printer.git
cd ios_better_printer
```

#### 2. 配置环境变量

```bash
cp .env.example .env
vim .env
```

修改以下配置：

```ini
# 企业微信配置
WECOM_CORP_ID=你的企业ID
WECOM_AGENT_ID=你的应用AgentId
WECOM_SECRET=你的应用Secret
WECOM_TOKEN=随机字符串（自己设置）
WECOM_ENCODING_AES_KEY=43位随机字符串（自己设置）

# Windows转换服务（可选）
WINDOWS_CONVERTER_ENABLED=true
WINDOWS_CONVERTER_URL=http://your-windows-ip:8080
```

#### 3. 启动服务

```bash
sudo docker compose up -d --build
```

#### 4. 检查服务状态

```bash
curl http://localhost:18080/health
```

### 三、配置企业微信应用

1. 进入企业微信管理后台 → 应用管理 → 你的应用
2. 找到 **接收消息** → **设置API接收**
3. 填写配置：
   - **URL**: `http://你的公网IP:18080/wecom`
   - **Token**: 与`.env`中的`WECOM_TOKEN`一致
   - **EncodingAESKey**: 与`.env`中的`WECOM_ENCODING_AES_KEY`一致
4. 点击 **保存**

### 四、开始使用

1. 在企业微信中打开你创建的应用
2. 直接发送Word/Excel文件
3. 等待5-15秒，收到转换后的PDF

## 支持的文件格式

- Microsoft Word: `.doc`, `.docx`
- Microsoft Excel: `.xls`, `.xlsx`
- Microsoft PowerPoint: `.ppt`, `.pptx`

## 项目结构

```
ios_better_printer/
├── app.py              # Flask主应用（企业微信回调处理）
├── wecom_api.py        # 企业微信API封装（含AES加解密）
├── converter.py        # 文档转换引擎
├── config.py           # 配置管理
├── Dockerfile          # Docker镜像定义
├── docker-compose.yml  # 服务编排
├── nginx.conf          # Nginx反向代理配置
└── windows_converter_service.py  # Windows Office转换服务
```

## 开发与调试

```bash
# 查看日志
sudo docker compose logs -f app

# 重启服务
sudo docker compose restart

# 重新构建
sudo docker compose up -d --build
```

## 许可证

MIT License

## 致谢

- LibreOffice - 开源办公套件
- Flask - Python Web框架
