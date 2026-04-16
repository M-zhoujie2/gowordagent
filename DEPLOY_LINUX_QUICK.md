# GOWordAgent Linux 快速部署指南

> 适用于银河麒麟 V10 及其他 Linux 发行版

## 🚀 一键部署

### 方式一：使用快速部署脚本（推荐）

```bash
# 1. 克隆或解压项目
cd GOWordAgentAddIn-master

# 2. 运行部署脚本
chmod +x deploy-linux.sh
./deploy-linux.sh
```

### 方式二：手动部署

```bash
# 1. 构建项目
dotnet publish GOWordAgent.WpsService \
  -c Release \
  -r linux-x64 \
  --self-contained true \
  -p:PublishSingleFile=true \
  -o ./backend

# 2. 运行部署脚本
./deploy-linux.sh
```

---

## 📋 部署前准备

### 系统要求

| 项目 | 要求 |
|------|------|
| 操作系统 | 银河麒麟 V10 / Ubuntu / CentOS |
| 架构 | x86_64 |
| .NET | 8.0（如使用自包含版本则无需安装） |
| WPS | 12.1+ for Linux |

### 检查系统

```bash
# 检查架构
uname -m  # 应输出 x86_64

# 检查 WPS
which wps

# 检查 systemd（可选）
systemctl --version
```

---

## 🔧 部署详解

### 1. 自动部署流程

部署脚本会自动完成以下步骤：

1. **检查依赖** - 检查 .NET 运行时（可选）
2. **创建目录** - `~/.local/opt/gowordagent` 和 `~/.config/gowordagent`
3. **复制文件** - 后端可执行文件和 WPS 加载项
4. **创建服务** - Systemd 用户服务或手动启动脚本
5. **启动服务** - 自动启动后端服务
6. **验证安装** - 检查端口文件和健康检查

### 2. 目录结构

```
~/.local/opt/gowordagent/       # 安装目录
├── gowordagent-server          # 主程序
├── start.sh                    # 启动脚本
└── run.sh                      # 手动管理脚本（无systemd时）

~/.config/gowordagent/          # 配置目录
└── config.dat                  # 加密配置文件

~/.local/share/Kingsoft/wps/jsaddons/com.gowordagent.addin/  # WPS插件
├── index.html
├── main.js
└── ...
```

### 3. 服务管理

#### 使用 Systemd（推荐）

```bash
# 查看状态
systemctl --user status gowordagent

# 启动/停止/重启
systemctl --user start gowordagent
systemctl --user stop gowordagent
systemctl --user restart gowordagent

# 查看日志
journalctl --user -u gowordagent -f
```

#### 使用手动脚本

```bash
cd ~/.local/opt/gowordagent

# 启动
./run.sh start

# 停止
./run.sh stop

# 查看状态
./run.sh status

# 查看日志
tail -f /tmp/gowordagent-$USER.log
```

---

## ✅ 验证安装

### 1. 检查服务状态

```bash
# 查看端口文件
PORT_FILE="/tmp/gowordagent-port-$USER.json"
cat "$PORT_FILE"

# 预期输出:
# {"Port":xxxxx,"Pid":xxxxx,"Timestamp":xxxxx}
```

### 2. 测试 API

```bash
PORT=$(cat "/tmp/gowordagent-port-$USER.json" | grep -o '"Port":[0-9]*' | cut -d: -f2)

# 健康检查
curl "http://127.0.0.1:$PORT/api/proofread/health"

# 预期输出:
# {"status":"ok","timestamp":"...","version":"1.0.0"}
```

### 3. 启动 WPS

1. 重启 WPS 文字
2. 在右侧边栏找到 **"智能校对"** 面板
3. 配置 AI 提供商和 API Key
4. 打开文档，点击 **"开始校对"**

---

## 🛠️ 故障排除

### 问题 1：服务无法启动

```bash
# 查看详细错误
~/.local/opt/gowordagent/start.sh

# 或查看日志
journalctl --user -u gowordagent -n 50
cat /tmp/gowordagent-$USER.log
```

### 问题 2：端口被占用

```bash
# 查看端口占用
netstat -tlnp | grep gowordagent

# 停止旧服务
systemctl --user stop gowordagent
# 或
killall gowordagent-server
```

### 问题 3：WPS 插件不显示

```bash
# 检查插件目录
ls -la ~/.local/share/Kingsoft/wps/jsaddons/com.gowordagent.addin/

# 检查 WPS 日志
cat ~/.local/share/Kingsoft/wps/log/wps.log
```

### 问题 4：无法连接后端

```bash
# 检查服务是否运行
systemctl --user status gowordagent

# 检查端口文件
ls -la /tmp/gowordagent-port-$USER.json

# 测试连接
PORT=$(cat "/tmp/gowordagent-port-$USER.json" | grep -o '"Port":[0-9]*' | cut -d: -f2)
curl -v "http://127.0.0.1:$PORT/api/proofread/health"
```

---

## 🔒 安全配置

### API Key 存储

API Key 使用 AES-GCM 加密存储在：
```
~/.config/gowordagent/config.dat
```

### 修改配置

```bash
# 直接编辑配置文件（需重启服务）
# 或通过 WPS 插件界面配置
```

---

## 🗑️ 卸载

### 完全卸载

```bash
# 停止服务
systemctl --user stop gowordagent
systemctl --user disable gowordagent

# 删除文件
rm -rf ~/.local/opt/gowordagent
rm -rf ~/.config/gowordagent
rm -rf ~/.local/share/Kingsoft/wps/jsaddons/com.gowordagent.addin
rm -f ~/.config/systemd/user/gowordagent.service

# 重载 systemd
systemctl --user daemon-reload

echo "卸载完成"
```

---

## 📚 相关文档

- [构建指南](KYLIN_V10_BUILD.md) - 详细构建说明
- [测试指南](TEST_GUIDE.md) - 功能测试步骤
- [项目概述](PROJECT_OVERVIEW.md) - 项目架构说明

---

## 💡 提示

1. **首次启动可能需要几秒钟** 初始化配置
2. **端口文件** 位于 `/tmp/gowordagent-port-$USER.json`
3. **日志文件** 使用 `journalctl` 或 `/tmp/gowordagent-$USER.log`
4. **配置更改** 后需要重启服务生效

---

**部署脚本**: `deploy-linux.sh`  
**最后更新**: 2026-04-13
