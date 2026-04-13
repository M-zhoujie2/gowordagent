#!/bin/bash
set -e

INSTALL_DIR=/opt/gowordagent
CONFIG_DIR=$HOME/.config/gowordagent
ADDON_DIR=$HOME/.local/share/Kingsoft/wps/jsaddons/com.gowordagent.addin
SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"

echo "=== GOWordAgent 安装脚本 ==="

# 1. 检查架构
ARCH=$(uname -m)
if [ "$ARCH" != "x86_64" ]; then
    echo "错误: 当前仅支持 x86_64 架构，检测到 $ARCH"
    exit 1
fi

# 2. 检查 WPS
if ! command -v wps &> /dev/null; then
    echo "错误: 未检测到 WPS Office"
    exit 1
fi

echo "检测到 WPS: $(wps --version 2>/dev/null || echo '版本未知')"

# 3. 检查后端文件
if [ ! -d "$SCRIPT_DIR/../backend" ]; then
    echo "错误: 找不到后端文件目录 (../backend)"
    echo "请确保已将编译后的后端文件复制到 backend/ 目录"
    exit 1
fi

# 4. 复制后端
echo "正在安装后端服务到 $INSTALL_DIR..."
sudo mkdir -p $INSTALL_DIR
sudo cp -r "$SCRIPT_DIR/../backend/"* $INSTALL_DIR/
sudo chmod +x $INSTALL_DIR/gowordagent-server

# 5. 注册 systemd 用户服务
echo "正在注册系统服务..."
mkdir -p $HOME/.config/systemd/user
cp "$SCRIPT_DIR/gowordagent.service" $HOME/.config/systemd/user/
systemctl --user daemon-reload
systemctl --user enable gowordagent
systemctl --user start gowordagent

# 6. 安装 WPS 加载项
echo "正在安装 WPS 加载项..."
mkdir -p $ADDON_DIR

if [ -d "$SCRIPT_DIR/../addon" ]; then
    cp -r "$SCRIPT_DIR/../addon/"* $ADDON_DIR/
else
    echo "警告: 找不到加载项文件 (../addon)"
fi

# 7. 创建配置目录
mkdir -p $CONFIG_DIR

echo ""
echo "=== 安装完成 ==="
echo ""
echo "后端服务状态:"
systemctl --user status gowordagent --no-pager || true
echo ""
echo "请重启 WPS 文字以加载插件"
echo ""
echo "卸载命令: ./uninstall.sh"
