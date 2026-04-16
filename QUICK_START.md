# GOWordAgent Linux 快速开始

## 🚀 最简单的一键部署

### 1. 获取发布包

```bash
# 方式一：下载预编译包（推荐）
wget https://github.com/yourusername/GOWordAgent/releases/download/v1.0.0/gowordagent-linux-x64-1.0.0.tar.gz
tar -xzf gowordagent-linux-x64-1.0.0.tar.gz
cd gowordagent-linux-x64-1.0.0

# 方式二：自己构建
./build-release.sh
cd release/gowordagent-linux-x64-1.0.0
```

### 2. 一键部署

```bash
./deploy-linux.sh
```

部署脚本会自动：
- ✓ 检查系统环境
- ✓ 安装到 `~/.local/opt/gowordagent`
- ✓ 创建 systemd 服务（如果可用）
- ✓ 安装 WPS 加载项
- ✓ 启动服务

### 3. 使用

1. 重启 WPS 文字
2. 在右侧边栏找到 **"智能校对"**
3. 配置 API Key（DeepSeek/GLM/Ollama）
4. 打开文档，点击 **"开始校对"**

---

## 📋 常用命令

```bash
# 查看服务状态
systemctl --user status gowordagent

# 查看日志
journalctl --user -u gowordagent -f

# 停止服务
systemctl --user stop gowordagent

# 重启服务
systemctl --user restart gowordagent
```

---

## 🔧 手动管理（无 systemd）

```bash
cd ~/.local/opt/gowordagent

# 启动
./run.sh start

# 停止
./run.sh stop

# 查看状态
./run.sh status
```

---

## 📁 文件位置

| 类型 | 路径 |
|------|------|
| 安装目录 | `~/.local/opt/gowordagent` |
| 配置文件 | `~/.config/gowordagent/config.dat` |
| 端口文件 | `/tmp/gowordagent-port-$USER.json` |
| WPS 插件 | `~/.local/share/Kingsoft/wps/jsaddons/com.gowordagent.addin` |
| 日志文件 | `journalctl` 或 `/tmp/gowordagent-$USER.log` |

---

## 🗑️ 卸载

```bash
# 停止并删除服务
systemctl --user stop gowordagent
systemctl --user disable gowordagent

# 删除文件
rm -rf ~/.local/opt/gowordagent
rm -rf ~/.config/gowordagent
rm -rf ~/.local/share/Kingsoft/wps/jsaddons/com.gowordagent.addin
rm -f ~/.config/systemd/user/gowordagent.service

# 重载 systemd
systemctl --user daemon-reload
```

---

## 📚 详细文档

- [完整部署指南](DEPLOY_LINUX_QUICK.md)
- [构建指南](KYLIN_V10_BUILD.md)
- [测试指南](TEST_GUIDE.md)

---

**提示**: 首次启动可能需要几秒钟初始化配置。
