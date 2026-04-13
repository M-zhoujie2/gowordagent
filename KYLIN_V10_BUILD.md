# GOWordAgent 银河麒麟 V10 构建指南

## 项目结构

```
GOWordAgent/
├── GOWordAgent.Core/           # .NET 8 共享类库
├── GOWordAgent.WpsService/     # .NET 8 后端服务
├── GOWordAgent.WpsAddon/       # WPS 加载项
├── Scripts/                    # 部署脚本
└── GOWordAgentAddIn/           # 原有 Word VSTO（不变）
```

## 环境要求

- .NET 8 SDK
- Windows（开发）/ 银河麒麟 V10（部署）
- WPS Office for Linux 12.1+

## 构建步骤

### 1. 构建 Core 类库

```bash
cd GOWordAgent.Core
dotnet build -c Release
```

### 2. 构建后端服务

```bash
cd GOWordAgent.WpsService
dotnet publish -c Release -r linux-x64 --self-contained true -p:PublishSingleFile=true
```

输出目录：`bin/Release/net8.0/linux-x64/publish/`

### 3. 打包

```bash
# 创建发布目录
mkdir -p release/backend
mkdir -p release/addon

# 复制后端文件
cp GOWordAgent.WpsService/bin/Release/net8.0/linux-x64/publish/* release/backend/

# 复制加载项文件
cp -r GOWordAgent.WpsAddon/* release/addon/

# 复制脚本
cp Scripts/* release/
```

## 部署到银河麒麟 V10

### 1. 复制到目标机器

```bash
scp -r release/ user@kylin-host:/tmp/gowordagent-release
```

### 2. 执行安装

```bash
ssh user@kylin-host
cd /tmp/gowordagent-release
./install.sh
```

### 3. 验证

```bash
# 检查服务状态
systemctl --user status gowordagent

# 检查端口文件
cat /tmp/gowordagent-port.json

# 测试 API
curl http://127.0.0.1:$(cat /tmp/gowordagent-port.json | grep -o '"port":[0-9]*' | cut -d: -f2)/api/proofread/health
```

### 4. 启动 WPS

重启 WPS 文字，在右侧边栏应该能看到"智能校对"面板。

## 卸载

```bash
cd /tmp/gowordagent-release
./uninstall.sh
```

## 日志查看

```bash
# 服务日志
journalctl --user -u gowordagent -f

# 后端日志（如配置了文件日志）
tail -f ~/.local/share/gowordagent/logs/service.log
```

## 故障排查

### 1. 服务无法启动

```bash
# 检查端口占用
netstat -tlnp | grep gowordagent

# 手动运行查看错误
/opt/gowordagent/gowordagent-server
```

### 2. WPS 加载项不显示

- 确认加载项目录权限正确
- 检查 WPS 版本（需 12.1+）
- 查看 WPS 日志：`~/.local/share/Kingsoft/wps/log/`

### 3. 无法连接后端

- 确认 `/tmp/gowordagent-port.json` 存在
- 检查防火墙设置
- 验证健康检查接口：`curl http://127.0.0.1:PORT/api/proofread/health`
