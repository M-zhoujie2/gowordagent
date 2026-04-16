#!/bin/bash
# GOWordAgent AppImage 构建脚本
# 创建类似 Windows .exe 的单文件可执行程序

set -e

echo "=== GOWordAgent AppImage 构建脚本 ==="
echo ""
echo "注意: 此方法需要 Linux 环境构建"
echo ""

VERSION="${VERSION:-1.0.0}"
APP_NAME="GOWordAgent"
APP_DIR="$APP_NAME.AppDir"

ARCH=$(uname -m)
APP_IMAGE_ARCH=""
RID=""
if [ "$ARCH" = "x86_64" ]; then
    APP_IMAGE_ARCH="x86_64"
    RID="linux-x64"
elif [ "$ARCH" = "aarch64" ] || [ "$ARCH" = "arm64" ]; then
    APP_IMAGE_ARCH="aarch64"
    RID="linux-arm64"
else
    echo "错误: 不支持的架构 $ARCH"
    exit 1
fi

echo "版本: $VERSION"
echo "架构: $ARCH ($APP_IMAGE_ARCH)"
echo ""

# 检查工具
if ! command -v wget &> /dev/null; then
    echo "需要 wget 下载工具"
    exit 1
fi

# 创建工作目录
WORK_DIR=$(mktemp -d)
trap "rm -rf $WORK_DIR" EXIT

echo "创建工作目录: $WORK_DIR"

# 创建 AppDir 结构
mkdir -p "$WORK_DIR/$APP_DIR/usr/bin"
mkdir -p "$WORK_DIR/$APP_DIR/usr/share/applications"
mkdir -p "$WORK_DIR/$APP_DIR/usr/share/icons/hicolor/256x256/apps"

# 复制后端文件
echo "复制后端文件..."
RELEASE_DIR="./release/gowordagent-${RID}-${VERSION}"
if [ -d "$RELEASE_DIR/backend" ]; then
    cp -r "$RELEASE_DIR/backend/"* "$WORK_DIR/$APP_DIR/usr/bin/"
elif [ -d "./GOWordAgent.WpsService/bin/Release/net8.0/$RID/publish" ]; then
    cp -r "./GOWordAgent.WpsService/bin/Release/net8.0/$RID/publish/"* "$WORK_DIR/$APP_DIR/usr/bin/"
else
    echo "错误: 找不到后端文件，请先构建项目"
    exit 1
fi

# 创建启动脚本
cat > "$WORK_DIR/$APP_DIR/AppRun" << 'EOF'
#!/bin/bash
# GOWordAgent AppRun - AppImage 入口点

HERE="$(dirname "$(readlink -f "${0}")")"

# 显示安装对话框
if [ -z "$APPIMAGE" ]; then
    echo "请在终端运行此程序: $0"
    exit 1
fi

# 创建安装向导
cat << 'INSTALL_WIZARD'
================================================================

           GOWordAgent 智能校对 安装向导

================================================================

欢迎使用 GOWordAgent!

此程序将安装:
  后端服务到: ~/.local/opt/gowordagent
  WPS 插件到: ~/.local/share/Kingsoft/wps/jsaddons/
  配置文件到: ~/.config/gowordagent

按回车键继续安装...
INSTALL_WIZARD

read

# 执行安装
export INSTALL_SOURCE="$HERE"
bash "$HERE/usr/bin/install-appimage.sh"

EOF
chmod +x "$WORK_DIR/$APP_DIR/AppRun"

# 创建桌面文件
cat > "$WORK_DIR/$APP_DIR/usr/share/applications/gowordagent.desktop" << EOF
[Desktop Entry]
Name=GOWordAgent
Comment=AI 智能校对插件
Exec=gowordagent-server
Type=Application
Categories=Office;
Icon=gowordagent
EOF

cp "$WORK_DIR/$APP_DIR/usr/share/applications/gowordagent.desktop" "$WORK_DIR/$APP_DIR/"

# 创建安装脚本
cat > "$WORK_DIR/$APP_DIR/usr/bin/install-appimage.sh" << 'INSTALLER_SCRIPT'
#!/bin/bash
# AppImage 内部安装脚本

set -e

RED='\033[0;31m'
GREEN='\033[0;32m'
YELLOW='\033[1;33m'
NC='\033[0m'

INSTALL_DIR="$HOME/.local/opt/gowordagent"
CONFIG_DIR="$HOME/.config/gowordagent"
WPS_ADDON_DIR="$HOME/.local/share/Kingsoft/wps/jsaddons/com.gowordagent.addin"

echo -e "${YELLOW}正在安装 GOWordAgent...${NC}"

# 创建目录
mkdir -p "$INSTALL_DIR"
mkdir -p "$CONFIG_DIR"
mkdir -p "$WPS_ADDON_DIR"

# 复制后端
cp -r "$INSTALL_SOURCE/usr/bin/"* "$INSTALL_DIR/" 2>/dev/null || true
chmod +x "$INSTALL_DIR/gowordagent-server"

# 复制 WPS 插件（从 AppImage 中提取或从网络下载）
if [ -d "$INSTALL_SOURCE/wps-addon" ]; then
    cp -r "$INSTALL_SOURCE/wps-addon/"* "$WPS_ADDON_DIR/"
fi

# 创建启动脚本
cat > "$INSTALL_DIR/start.sh" << 'START_SCRIPT'
#!/bin/bash
cd "$(dirname "$0")"
./gowordagent-server "$@"
START_SCRIPT
chmod +x "$INSTALL_DIR/start.sh"

# 创建 systemd 服务
if systemctl --version &> /dev/null; then
    mkdir -p "$HOME/.config/systemd/user"
    cat > "$HOME/.config/systemd/user/gowordagent.service" << SERVICE_EOF
[Unit]
Description=GOWordAgent Backend Service
After=network.target

[Service]
Type=simple
ExecStart=$INSTALL_DIR/gowordagent-server
Restart=on-failure
RestartSec=3

[Install]
WantedBy=default.target
SERVICE_EOF

    systemctl --user daemon-reload
    systemctl --user enable gowordagent
    systemctl --user start gowordagent

    echo -e "${GREEN}Systemd 服务已创建并启动${NC}"
else
    # 手动启动脚本
    cat > "$INSTALL_DIR/run.sh" << 'RUN_SCRIPT'
#!/bin/bash
PIDFILE="/tmp/gowordagent-$USER.pid"
case "$1" in
    start) nohup "$(dirname "$0")/gowordagent-server" > /tmp/gowordagent-$USER.log 2>&1 & echo $! > "$PIDFILE" ;;
    stop) kill $(cat "$PIDFILE" 2>/dev/null) 2>/dev/null; rm -f "$PIDFILE" ;;
    status) kill -0 $(cat "$PIDFILE" 2>/dev/null) 2>/dev/null && echo "运行中" || echo "未运行" ;;
esac
RUN_SCRIPT
    chmod +x "$INSTALL_DIR/run.sh"
    "$INSTALL_DIR/run.sh" start

    echo -e "${GREEN}服务已启动（手动模式）${NC}"
fi

# 等待服务启动
sleep 2

# 验证安装
PORT=""
RUNTIME_PORT_FILE="${XDG_RUNTIME_DIR:-}/gowordagent-port-$USER.json"
LEGACY_PORT_FILE="/tmp/gowordagent-port-$USER.json"
CONFIG_PORT_FILE="$CONFIG_DIR/service-port.json"

if [ -f "$RUNTIME_PORT_FILE" ]; then
    PORT=$(grep -o '"Port":[0-9]*' "$RUNTIME_PORT_FILE" | cut -d: -f2)
elif [ -f "$LEGACY_PORT_FILE" ]; then
    PORT=$(grep -o '"Port":[0-9]*' "$LEGACY_PORT_FILE" | cut -d: -f2)
elif [ -f "$CONFIG_PORT_FILE" ]; then
    PORT=$(grep -o '"Port":[0-9]*' "$CONFIG_PORT_FILE" | cut -d: -f2)
fi

if [ -n "$PORT" ]; then
    echo -e "${GREEN}安装成功!${NC}"
    echo ""
    echo "请重启 WPS 文字，在右侧边栏找到'智能校对'面板"
else
    echo -e "${YELLOW}服务启动可能需要几秒，请稍后检查${NC}"
fi

echo ""
echo "安装目录: $INSTALL_DIR"
echo "配置目录: $CONFIG_DIR"
INSTALLER_SCRIPT

chmod +x "$WORK_DIR/$APP_DIR/usr/bin/install-appimage.sh"

echo ""
echo "目录结构:"
find "$WORK_DIR/$APP_DIR" -type f | head -20

echo ""
echo -e "${GREEN}AppDir 创建完成${NC}"
echo ""
echo "要创建最终的 AppImage，请下载 appimagetool:"
echo "  wget https://github.com/AppImage/AppImageKit/releases/download/continuous/appimagetool-${APP_IMAGE_ARCH}.AppImage"
echo "  chmod +x appimagetool-${APP_IMAGE_ARCH}.AppImage"
echo "  ./appimagetool-${APP_IMAGE_ARCH}.AppImage $WORK_DIR/$APP_DIR GOWordAgent-${VERSION}-${APP_IMAGE_ARCH}.AppImage"
echo ""

# 如果 appimagetool 存在，直接构建
APP_IMAGE_TOOL="./appimagetool-${APP_IMAGE_ARCH}.AppImage"
if command -v appimagetool &> /dev/null || [ -f "$APP_IMAGE_TOOL" ]; then
    echo "发现 appimagetool，正在构建 AppImage..."
    if [ -f "$APP_IMAGE_TOOL" ]; then
        "$APP_IMAGE_TOOL" "$WORK_DIR/$APP_DIR" "GOWordAgent-${VERSION}-${APP_IMAGE_ARCH}.AppImage"
    else
        appimagetool "$WORK_DIR/$APP_DIR" "GOWordAgent-${VERSION}-${APP_IMAGE_ARCH}.AppImage"
    fi

    echo ""
    echo -e "${GREEN}AppImage 创建成功: GOWordAgent-${VERSION}-${APP_IMAGE_ARCH}.AppImage${NC}"
    echo ""
    echo "使用方式:"
    echo "  1. 双击 GOWordAgent-${VERSION}-${APP_IMAGE_ARCH}.AppImage"
    echo "  2. 或在终端运行: ./GOWordAgent-${VERSION}-${APP_IMAGE_ARCH}.AppImage"
fi
