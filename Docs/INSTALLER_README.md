# GOWordAgent 智能校对系统 - Linux 安装包

## 系统要求

- **操作系统**: 银河麒麟 V10 / Ubuntu 20.04+ / CentOS 8+ (x86_64)
- **办公软件**: WPS Office for Linux 12.1.2+
- **内存**: 4GB+
- **磁盘**: 200MB 可用空间

## 快速安装

### 方法一：图形界面安装（推荐）

1. **双击运行** `GOWordAgent-Install-linux-x64-1.0.0.run`
2. 点击"开始安装"按钮
3. 等待安装完成
4. 重启 WPS 文字

### 方法二：命令行安装

```bash
# 添加执行权限
chmod +x GOWordAgent-Install-linux-x64-1.0.0.run

# 运行安装
./GOWordAgent-Install-linux-x64-1.0.0.run
```

## 首次使用

1. 打开 WPS 文字
2. 在右侧边栏找到 **"智能校对"** 面板
3. 点击 **"保存并连接"** 配置 AI 提供商
4. 输入 API Key 并测试连接
5. 打开文档，点击 **"开始校对"**

## 支持的 AI 提供商

- **DeepSeek**: https://platform.deepseek.com
- **智谱 AI**: https://open.bigmodel.cn
- **Ollama**: 本地部署（http://localhost:11434）

## 服务管理

```bash
# 查看状态
~/.local/opt/gowordagent/run.sh status

# 停止服务
~/.local/opt/gowordagent/run.sh stop

# 启动服务
~/.local/opt/gowordagent/run.sh start

# 查看日志
~/.local/opt/gowordagent/run.sh logs
```

如果使用 Systemd：

```bash
# 查看状态
systemctl --user status gowordagent

# 停止/启动
systemctl --user stop gowordagent
systemctl --user start gowordagent

# 查看日志
journalctl --user -u gowordagent -f
```

## 卸载

```bash
# 停止服务
~/.local/opt/gowordagent/run.sh stop 2>/dev/null
systemctl --user stop gowordagent 2>/dev/null

# 删除文件
rm -rf ~/.local/opt/gowordagent
rm -rf ~/.config/gowordagent
rm -rf ~/.local/share/Kingsoft/wps/jsaddons/com.gowordagent.addin
```

## 故障排除

### 安装后 WPS 中看不到插件

1. 确保已重启 WPS 文字
2. 检查插件文件是否存在：
   ```bash
   ls ~/.local/share/Kingsoft/wps/jsaddons/com.gowordagent.addin/
   ```
3. 尝试重新安装

### 提示"无法连接后端服务"

1. 检查服务状态：
   ```bash
   ~/.local/opt/gowordagent/run.sh status
   ```
2. 手动启动服务：
   ```bash
   ~/.local/opt/gowordagent/run.sh start
   ```
3. 查看错误日志：
   ```bash
   ~/.local/opt/gowordagent/run.sh logs
   ```

### API 连接测试失败

1. 检查网络连接
2. 确认 API Key 正确
3. 检查服务商状态页面

## 技术支持

如有问题，请：

1. 查看日志文件：`/tmp/gowordagent-<用户名>.log`
2. 联系技术支持

---

**版本**: 1.0.0  
**更新日期**: 2026-04-13
