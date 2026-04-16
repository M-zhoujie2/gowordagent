#!/bin/bash
# ============================================================
# GOWordAgent Linux 一键安装程序
# 将此脚本和发布文件一起分发，双击即可安装
# ============================================================

# 安装程序版本
INSTALLER_VERSION="1.0.0"

# 颜色定义
RED='\033[0;31m'
GREEN='\033[0;32m'
YELLOW='\033[1;33m'
BLUE='\033[0;34m'
CYAN='\033[0;36m'
NC='\033[0m' # No Color

# 清屏（可选）
clear 2>/dev/null || true

# 显示欢迎界面
echo -e "${CYAN}"
echo "================================================================"
echo ""
echo "             GOWordAgent 智能校对系统"
echo "                   Linux 安装程序"
echo ""
echo "================================================================"
echo -e "${NC}"
echo ""
echo -e "版本: ${GREEN}$INSTALLER_VERSION${NC}"
echo -e "支持系统: ${GREEN}银河麒麟 V10 / Ubuntu / CentOS / UOS (x86_64 & arm64)${NC}"
echo ""

# 检查是否支持图形界面
HAVE_ZENITY=false
if command -v zenity &> /dev/null; then
    HAVE_ZENITY=true
fi

# 显示消息函数
show_message() {
    if [ "$HAVE_ZENITY" = true ] && [ -n "$DISPLAY" ]; then
        zenity --info --title="GOWordAgent 安装" --text="$1" 2>/dev/null || echo -e "$1"
    else
        echo -e "$1"
    fi
}

show_error() {
    if [ "$HAVE_ZENITY" = true ] && [ -n "$DISPLAY" ]; then
        zenity --error --title="安装错误" --text="$1" 2>/dev/null || echo -e "${RED}$1${NC}"
    else
        echo -e "${RED}$1${NC}"
    fi
}

# 检查架构
echo -e "${BLUE}[1/5] 检查系统架构...${NC}"
ARCH=$(uname -m)
VALID_ARCH=false
RID=""
if [ "$ARCH" = "x86_64" ]; then
    VALID_ARCH=true
    RID="linux-x64"
elif [ "$ARCH" = "aarch64" ] || [ "$ARCH" = "arm64" ]; then
    VALID_ARCH=true
    RID="linux-arm64"
fi

if [ "$VALID_ARCH" = false ]; then
    show_error "错误: 当前仅支持 x86_64 和 arm64 架构，检测到 $ARCH"
    echo ""
    read -p "按回车键退出..."
    exit 1
fi
echo -e "  ${GREEN}架构检查通过: $ARCH ($RID)${NC}"

# 设置安装路径
echo ""
echo -e "${BLUE}[2/5] 配置安装路径...${NC}"
INSTALL_DIR="${INSTALL_DIR:-$HOME/.local/opt/gowordagent}"
CONFIG_DIR="$HOME/.config/gowordagent"
WPS_ADDON_DIR="$HOME/.local/share/Kingsoft/wps/jsaddons/com.gowordagent.addin"

echo "  安装目录: $INSTALL_DIR"
echo "  配置目录: $CONFIG_DIR"
echo "  WPS插件目录: $WPS_ADDON_DIR"

# 确认安装
echo ""
echo -e "${YELLOW}准备开始安装...${NC}"
echo ""
if [ "$HAVE_ZENITY" = true ] && [ -n "$DISPLAY" ]; then
    if ! zenity --question --title="确认安装" --text="是否开始安装 GOWordAgent?\n\n安装目录: $INSTALL_DIR" --ok-label="开始安装" --cancel-label="取消" 2>/dev/null; then
        echo "安装已取消"
        exit 0
    fi
else
    read -p "是否开始安装? [Y/n] " -n 1 -r
    echo
    if [[ ! $REPLY =~ ^[Yy]$ ]] && [ -n "$REPLY" ]; then
        echo "安装已取消"
        exit 0
    fi
fi

echo ""
echo -e "${BLUE}[3/5] 解压安装文件...${NC}"

# 获取脚本所在目录
SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"

# 检查文件是否存在
BACKEND_DIR="$SCRIPT_DIR/backend"
if [ "$RID" != "linux-x64" ] && [ -d "$SCRIPT_DIR/backend-$RID" ]; then
    BACKEND_DIR="$SCRIPT_DIR/backend-$RID"
fi

if [ ! -d "$BACKEND_DIR" ]; then
    show_error "错误: 找不到后端文件目录 ($BACKEND_DIR)"
    echo "请确保安装文件完整"
    read -p "按回车键退出..."
    exit 1
fi

# 创建目录
echo "  创建安装目录..."
mkdir -p "$INSTALL_DIR"
mkdir -p "$CONFIG_DIR"
mkdir -p "$WPS_ADDON_DIR"

# 复制后端文件
echo "  复制后端文件..."
cp -r "$BACKEND_DIR/"* "$INSTALL_DIR/"
chmod +x "$INSTALL_DIR/gowordagent-server"

# 复制 WPS 插件
echo "  复制 WPS 插件..."
if [ -d "$SCRIPT_DIR/addon" ]; then
    cp -r "$SCRIPT_DIR/addon/"* "$WPS_ADDON_DIR/"
fi

echo -e "  ${GREEN}文件复制完成${NC}"

echo ""
echo -e "${BLUE}[4/5] 配置系统服务...${NC}"

# 创建启动脚本
cat > "$INSTALL_DIR/start.sh" << 'EOF'
#!/bin/bash
cd "$(dirname "$0")"
export DOTNET_CLI_TELEMETRY_OPTOUT=1
export DOTNET_USE_POLLING_FILE_WATCHER=true
./gowordagent-server "$@"
EOF
chmod +x "$INSTALL_DIR/start.sh"

# 尝试创建 systemd 服务
if systemctl --version &> /dev/null; then
    echo "  创建 systemd 用户服务..."
    mkdir -p "$HOME/.config/systemd/user"

    cat > "$HOME/.config/systemd/user/gowordagent.service" << EOF
[Unit]
Description=GOWordAgent Backend Service
After=network.target

[Service]
Type=simple
ExecStart=$INSTALL_DIR/gowordagent-server
Restart=on-failure
RestartSec=3
Environment=DOTNET_CLI_TELEMETRY_OPTOUT=1
Environment=DOTNET_USE_POLLING_FILE_WATCHER=true
Environment=LANG=zh_CN.UTF-8

[Install]
WantedBy=default.target
EOF

    systemctl --user daemon-reload 2>/dev/null || true
    systemctl --user stop gowordagent 2>/dev/null || true
    systemctl --user start gowordagent
    systemctl --user enable gowordagent 2>/dev/null || true

    echo -e "  ${GREEN}Systemd 服务已启用${NC}"
else
    echo "  Systemd 不可用，创建手动启动脚本..."

    cat > "$INSTALL_DIR/run.sh" << 'EOF'
#!/bin/bash
PIDFILE="/tmp/gowordagent-$USER.pid"
case "$1" in
    start)
        if [ -f "$PIDFILE" ] && kill -0 $(cat "$PIDFILE") 2>/dev/null; then
            echo "服务已在运行 (PID: $(cat $PIDFILE))"
            exit 0
        fi
        cd "$(dirname "$0")"
        nohup ./gowordagent-server > /tmp/gowordagent-$USER.log 2>&1 &
        echo $! > "$PIDFILE"
        echo "服务已启动 (PID: $(cat $PIDFILE))"
        ;;
    stop)
        if [ -f "$PIDFILE" ]; then
            kill $(cat "$PIDFILE") 2>/dev/null || true
            rm -f "$PIDFILE"
            echo "服务已停止"
        else
            echo "服务未运行"
        fi
        ;;
    status)
        if [ -f "$PIDFILE" ] && kill -0 $(cat "$PIDFILE") 2>/dev/null; then
            echo "服务正在运行 (PID: $(cat $PIDFILE))"
        else
            echo "服务未运行"
        fi
        ;;
    *)
        echo "用法: $0 {start|stop|status}"
        exit 1
        ;;
esac
EOF
    chmod +x "$INSTALL_DIR/run.sh"
    "$INSTALL_DIR/run.sh" start

    echo -e "  ${GREEN}手动服务已启用${NC}"
fi

echo ""
echo -e "${BLUE}[5/5] 验证安装...${NC}"

# 等待服务启动
sleep 2

PORT=""
RUNTIME_PORT_FILE="${XDG_RUNTIME_DIR:-}/gowordagent-port-$USER.json"
LEGACY_PORT_FILE="/tmp/gowordagent-port-$USER.json"
CONFIG_PORT_FILE="$CONFIG_DIR/service-port.json"

if [ -f "$RUNTIME_PORT_FILE" ]; then
    PORT=$(cat "$RUNTIME_PORT_FILE" | grep -o '"Port":[0-9]*' | cut -d: -f2)
    echo "  发现运行时端口文件: $RUNTIME_PORT_FILE"
elif [ -f "$LEGACY_PORT_FILE" ]; then
    PORT=$(cat "$LEGACY_PORT_FILE" | grep -o '"Port":[0-9]*' | cut -d: -f2)
    echo "  发现旧版端口文件: $LEGACY_PORT_FILE"
elif [ -f "$CONFIG_PORT_FILE" ]; then
    PORT=$(cat "$CONFIG_PORT_FILE" | grep -o '"Port":[0-9]*' | cut -d: -f2)
    echo "  发现配置端口文件: $CONFIG_PORT_FILE"
fi

if [ -n "$PORT" ]; then
    echo -e "  ${GREEN}服务已启动，端口: $PORT${NC}"

    # 测试健康检查
    if curl -s "http://127.0.0.1:$PORT/api/proofread/health" > /dev/null 2>&1; then
        echo -e "  ${GREEN}健康检查通过${NC}"
    else
        echo -e "  ${YELLOW}! 健康检查失败，服务可能仍在启动中${NC}"
    fi
else
    echo -e "  ${YELLOW}! 未检测到端口文件，服务可能启动失败${NC}"
fi

# 安装完成界面
echo ""
echo -e "${GREEN}"
echo "================================================================"
echo ""
echo "                     安装成功!"
echo ""
echo "================================================================"
echo -e "${NC}"
echo ""
echo -e "${CYAN}使用说明:${NC}"
echo "  1. 重启 WPS 文字"
echo "  2. 在右侧边栏找到 ${YELLOW}'智能校对'${NC} 面板"
echo "  3. 点击 ${YELLOW}'保存并连接'${NC} 配置 AI 提供商"
echo "  4. 打开文档，点击 ${YELLOW}'开始校对'${NC}"
echo ""
echo -e "${CYAN}服务管理:${NC}"
if systemctl --version &> /dev/null; then
    echo "  查看状态: systemctl --user status gowordagent"
    echo "  停止服务: systemctl --user stop gowordagent"
    echo "  查看日志: journalctl --user -u gowordagent -f"
else
    echo "  查看状态: $INSTALL_DIR/run.sh status"
    echo "  停止服务: $INSTALL_DIR/run.sh stop"
    echo "  查看日志: tail -f /tmp/gowordagent-$USER.log"
fi
echo ""
echo -e "${CYAN}卸载命令:${NC}"
echo "  rm -rf $INSTALL_DIR $CONFIG_DIR $WPS_ADDON_DIR"
echo ""

# 显示完成对话框
if [ "$HAVE_ZENITY" = true ] && [ -n "$DISPLAY" ]; then
    zenity --info --title="安装完成" --text="GOWordAgent 安装成功!\n\n请重启 WPS 文字，在右侧边栏找到'智能校对'面板。" 2>/dev/null
fi

# 等待用户按键
read -p "按回车键关闭..."

exit 0
