# 微信作业排版助手

> 📄 将老师发送的Word/Excel作业文件完美转换为PDF，解决iPhone打印排版错乱问题

## 项目简介

这是一个基于微信公众号的Office文档转PDF转换服务。系统采用**双引擎架构**（Windows Office + LibreOffice），确保转换后的PDF排版与Windows环境**100%一致**，彻底解决iOS预览导致的页数增加、内容重叠等问题。

### 核心功能

- ✅ **微信内闭环**：无需切换App，直接在微信完成转换
- ✅ **100%保真**：支持调用Windows Office引擎，排版完美还原
- ✅ **高可用**：主要引擎故障时自动降级到LibreOffice
- ✅ **即转即用**：3-10秒快速转换，支持Word/Excel/PPT
- ✅ **零学习成本**：转发文件即可，自动处理并返回PDF

### 适用场景

- 家长接收老师发送的作业文件，需远程打印
- iPhone用户需要保证文档排版与Windows一致
- 避免iOS Office预览导致的格式错误

## 快速开始

### 前置要求

- ESXi虚拟机或Linux服务器（Ubuntu 22.04推荐）
- Docker 和 Docker Compose
- **Windows虚拟机**（可选，用于实现100%排版还原）
- 固定公网IP地址
- **微信测试号**（推荐先用此方案验证）或正式公众号

### 一、申请微信测试号

1. 访问：https://mp.weixin.qq.com/debug/cgi-bin/sandbox?t=sandbox/login
2. 扫码登录后获取：
   - `appID`
   - `appsecret`
3. 记录这两个值，后续配置需要

### 二、部署服务

#### 1. 克隆项目

```bash
git clone <repository-url>
cd ios_better_printer
```

#### 2. 配置环境变量

```bash
cp .env.example .env
vim .env
```

修改以下配置：

```bash
WECHAT_APP_ID=你的AppID
WECHAT_APP_SECRET=你的AppSecret
WECHAT_TOKEN=随机字符串（自己设置，例如：mytoken123）
```

#### 3. 启动服务

```bash
# 构建并启动
sudo docker compose up -d --build

# 查看日志
sudo docker compose logs -f
```

> ⚠️ **端口说明**：服务默认使用 **18080** 端口（80/8080常被运营商封锁），需在路由器配置端口转发。

#### 4. 检查服务状态

```bash
curl http://localhost:18080/health
# 应返回: {"status":"ok","service":"wechat-doc-converter"}
```

### 三、配置微信测试号

1. 返回微信测试号管理页面
2. 找到 **接口配置信息** 部分
3. 填写配置：
   - **URL**: `http://你的公网IP:18080/wechat`
   - **Token**: 与`.env`中的`WECHAT_TOKEN`一致
4. 点击 **提交**，等待验证通过（显示绿色对勾）

### 四、开始使用

1. 用手机微信扫描测试号页面的二维码，关注测试号
2. 在任意聊天中，长按Word/Excel文件 → **转发** → 选择测试号
3. 等待3-10秒，收到转换后的PDF
4. 将PDF转发给打印机小程序，完成打印

## 使用说明

### 支持的文件格式

- Microsoft Word: `.doc`, `.docx`
- Microsoft Excel: `.xls`, `.xlsx`
- Microsoft PowerPoint: `.ppt`, `.pptx`

### 常见问题

**Q: 转换速度多快？**  
A: 通常5-10秒，取决于文件大小和复杂度。

**Q: 会不会丢失原始格式？**  
A: LibreOffice渲染引擎兼容性非常好，绝大多数格式都能完美保留。复杂的特殊字体可能会被替换为近似字体。

**Q: 测试号100人限制够用吗？**  
A: 对于家庭使用完全足够。如需更多用户，可升级到正式订阅号。

**Q: 转换失败怎么办？**  
A: 检查文件是否损坏，或者文件格式是否支持。查看服务日志：`sudo docker compose logs -f`

**Q: 如何升级到正式公众号？**  
A: 参考 [DEPLOYMENT.md](DEPLOYMENT.md) 的正式部署章节，需要购买域名并完成ICP备案。

## 项目结构

```
ios_better_printer/
├── app.py              # Flask主应用（webhook处理）
├── converter.py        # LibreOffice转换引擎
├── wechat_api.py       # 微信API封装
├── config.py           # 配置管理
├── Dockerfile          # Docker镜像定义
├── docker-compose.yml  # 服务编排
├── nginx.conf          # Nginx反向代理配置
├── requirements.txt    # Python依赖
├── .env.example        # 环境变量模板
└── README.md           # 本文件
```

## 技术栈

- **后端**: Python 3.11 + Flask
- **转换引擎**: LibreOffice Headless
- **字体**: Noto Sans CJK、文泉驿微米黑
- **容器化**: Docker + Docker Compose
- **反向代理**: Nginx

## 开发与调试

### 查看日志

```bash
# 查看所有服务日志
sudo docker compose logs -f

# 仅查看应用日志
sudo docker compose logs -f app
```

### 重启服务

```bash
sudo docker compose restart
```

### 停止服务

```bash
sudo docker compose down
```

### 重新构建

```bash
sudo docker compose up -d --build
```

## 安全建议

1. **不要在公共仓库提交 `.env` 文件**（已加入`.gitignore`）
2. **使用强密码**作为`WECHAT_TOKEN`
3. **启用HTTPS**：生产环境建议配置SSL证书
4. **定期清理临时文件**：`temp_files/`目录会自动清理，但建议定期检查

## 许可证

MIT License

## 贡献

欢迎提交Issue和Pull Request！

## 致谢

- LibreOffice - 开源办公套件
- Flask - Python Web框架
- Noto Fonts & 文泉驿 - 开源中文字体
