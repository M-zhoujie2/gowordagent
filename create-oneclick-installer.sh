#!/bin/bash
# ============================================================
# GOWordAgent 单文件安装器创建脚本
# 在 Linux 上执行，生成 .run 自解压安装程序
# ============================================================

set -e

INSTALLER_VERSION="1.0.0"
OUTPUT_FILE="GOWordAgent-Install-linux-x64-${INSTALLER_VERSION}.run"
TEMP_DIR=$(mktemp -d)
trap "rm -rf $TEMP_DIR" EXIT

echo "========================================"
echo "GOWordAgent 单文件安装器创建工具"
echo "========================================"
echo ""

# 检查发布文件是否存在
RELEASE_BASE="release/gowordagent-linux-x64-${INSTALLER_VERSION}"
if [ ! -d "$RELEASE_BASE/backend" ]; then
    echo "错误: 发布文件不存在"
    echo "请先运行 ./build-release.sh 构建发布版本"
    exit 1
fi

echo "[1/4] 准备安装文件..."

# 复制文件到临时目录
cp -r "$RELEASE_BASE/backend" "$TEMP_DIR/"
cp -r "$RELEASE_BASE/addon" "$TEMP_DIR/"

echo "[2/4] 打包资源..."
cd "$TEMP_DIR"
tar -czf resources.tar.gz backend addon
cd - > /dev/null

echo "[3/4] 生成安装器脚本..."

# 创建安装器脚本
cat > "$OUTPUT_FILE" << 'INSTALLER_EOF'
#!/bin/bash
# ============================================================
# GOWordAgent Linux 单文件安装器
# 这是一个自解压安装程序
# ============================================================

INSTALLER_VERSION="INSTALLER_VERSION_PLACEHOLDER"
INSTALLER_SIZE=INSTALLER_SIZE_PLACEHOLDER

# 颜色定义
RED='\033[0;31m'
GREEN='\033[0;32m'
YELLOW='\033[1;33m'
BLUE='\033[0;34m'
CYAN='\033[0;36m'
NC='\033[0m'

clear_screen() {
    command clear 2>/dev/null || printf '\033[2J\033[H'
}

# 清屏
clear_screen

# 显示欢迎界面
echo -e "${CYAN}"
echo "================================================================"
echo ""
echo "             GOWordAgent 智能校对系统"
echo "          Linux 单文件安装程序 vINSTALLER_VERSION_PLACEHOLDER"
echo ""
echo "================================================================"
echo -e "${NC}"
echo ""
echo -e "支持系统: ${GREEN}银河麒麟 V10 / Ubuntu / CentOS / UOS (x86_64 & arm64)${NC}"
echo ""

# 检查是否支持图形界面
HAVE_ZENITY=false
if command -v zenity &> /dev/null && [ -n "$DISPLAY" ]; then
    HAVE_ZENITY=true
fi

# 检查架构
echo -e "${BLUE}[1/5] 检查系统架构...${NC}"
ARCH=$(uname -m)
VALID_ARCH=false
if [ "$ARCH" = "x86_64" ] || [ "$ARCH" = "aarch64" ] || [ "$ARCH" = "arm64" ]; then
    VALID_ARCH=true
fi

if [ "$VALID_ARCH" = false ]; then
    echo -e "${RED}错误: 当前仅支持 x86_64 和 arm64 架构，检测到 $ARCH${NC}"
    read -p "按回车键退出..."
    exit 1
fi
echo -e "  ${GREEN}架构检查通过: $ARCH${NC}"

# 检查依赖
echo -e "${BLUE}[2/5] 检查系统依赖...${NC}"
MISSING_DEPS=""
if ! command -v tar &> /dev/null; then
    MISSING_DEPS="$MISSING_DEPS tar"
fi
if [ -n "$MISSING_DEPS" ]; then
    echo -e "${RED}错误: 缺少依赖:$MISSING_DEPS${NC}"
    echo "请使用包管理器安装依赖"
    read -p "按回车键退出..."
    exit 1
fi
echo -e "  ${GREEN}依赖检查通过${NC}"

# 设置安装路径
echo ""
echo -e "${BLUE}[3/5] 配置安装路径...${NC}"
INSTALL_DIR="${INSTALL_DIR:-$HOME/.local/opt/gowordagent}"
CONFIG_DIR="$HOME/.config/gowordagent"
WPS_ADDON_DIR="$HOME/.local/share/Kingsoft/wps/jsaddons/com.gowordagent.addin"

echo "  安装目录: $INSTALL_DIR"
echo "  配置目录: $CONFIG_DIR"

# 确认安装
echo ""
echo -e "${YELLOW}准备开始安装...${NC}"
if [ "$HAVE_ZENITY" = true ]; then
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
echo -e "${BLUE}[4/5] 解压安装文件...${NC}"

# 创建临时目录
TEMP_DIR=$(mktemp -d)
trap "rm -rf $TEMP_DIR" EXIT

# 从脚本中提取资源文件
SCRIPT_SIZE=INSTALLER_SIZE_PLACEHOLDER
tail -c +$(($SCRIPT_SIZE + 1)) "$0" > "$TEMP_DIR/resources.tar.gz"

if [ ! -s "$TEMP_DIR/resources.tar.gz" ]; then
    echo -e "${RED}错误: 无法提取资源文件${NC}"
    read -p "按回车键退出..."
    exit 1
fi

# 解压资源
echo "  解压资源文件..."
cd "$TEMP_DIR"
tar -xzf resources.tar.gz

# 创建安装目录
echo "  创建安装目录..."
mkdir -p "$INSTALL_DIR"
mkdir -p "$CONFIG_DIR"
mkdir -p "$WPS_ADDON_DIR"

# 安装后端
echo "  安装后端服务..."
cp -r "$TEMP_DIR/backend/"* "$INSTALL_DIR/"
chmod +x "$INSTALL_DIR/gowordagent-server"

# 安装 WPS 插件
echo "  安装 WPS 插件..."
if [ -d "$TEMP_DIR/addon" ]; then
    cp -r "$TEMP_DIR/addon/"* "$WPS_ADDON_DIR/"
fi

echo -e "  ${GREEN}文件安装完成${NC}"

echo ""
echo -e "${BLUE}[5/5] 配置系统服务...${NC}"

# 创建启动脚本
cat > "$INSTALL_DIR/start.sh" << 'STARTER_EOF'
#!/bin/bash
cd "$(dirname "$0")"
export DOTNET_CLI_TELEMETRY_OPTOUT=1
export DOTNET_USE_POLLING_FILE_WATCHER=true
./gowordagent-server "$@"
STARTER_EOF
chmod +x "$INSTALL_DIR/start.sh"

# 尝试创建 systemd 服务
if systemctl --version &> /dev/null 2>&1; then
    echo "  创建 systemd 用户服务..."
    mkdir -p "$HOME/.config/systemd/user"

    cat > "$HOME/.config/systemd/user/gowordagent.service" << SERVICEMD_EOF
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
SERVICEMD_EOF

    systemctl --user daemon-reload 2>/dev/null || true
    systemctl --user stop gowordagent 2>/dev/null || true
    systemctl --user start gowordagent 2>/dev/null || true
    systemctl --user enable gowordagent 2>/dev/null || true

    if systemctl --user is-active gowordagent &> /dev/null; then
        echo -e "  ${GREEN}Systemd 服务已启用${NC}"
    else
        echo -e "  ${YELLOW}! Systemd 服务可能未启动，将使用备用方案${NC}"
    fi
else
    echo "  Systemd 不可用，创建手动启动脚本..."
fi

# 备用启动方案（手动脚本）
cat > "$INSTALL_DIR/run.sh" << 'RUNSCRIPT_EOF'
#!/bin/bash
PIDFILE="/tmp/gowordagent-$USER.pid"
LOGFILE="/tmp/gowordagent-$USER.log"

case "$1" in
    start)
        if [ -f "$PIDFILE" ] && kill -0 $(cat "$PIDFILE") 2>/dev/null; then
            echo "服务已在运行 (PID: $(cat $PIDFILE))"
            exit 0
        fi
        cd "$(dirname "$0")"
        nohup ./gowordagent-server > "$LOGFILE" 2>&1 &
        echo $! > "$PIDFILE"
        echo "服务已启动 (PID: $(cat $PIDFILE))"
        sleep 1
        if [ -f "/tmp/gowordagent-port-$USER.json" ]; then
            PORT=$(grep -o '"Port":[0-9]*' "/tmp/gowordagent-port-$USER.json" | cut -d: -f2)
            echo "服务端口: $PORT"
        fi
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
    restart)
        $0 stop
        sleep 1
        $0 start
        ;;
    status)
        if [ -f "$PIDFILE" ] && kill -0 $(cat "$PIDFILE") 2>/dev/null; then
            echo "服务正在运行 (PID: $(cat $PIDFILE))"
            if [ -f "/tmp/gowordagent-port-$USER.json" ]; then
                PORT=$(grep -o '"Port":[0-9]*' "/tmp/gowordagent-port-$USER.json" | cut -d: -f2)
                echo "服务端口: $PORT"
            fi
        else
            echo "服务未运行"
        fi
        ;;
    logs)
        tail -f "$LOGFILE"
        ;;
    *)
        echo "用法: $0 {start|stop|restart|status|logs}"
        exit 1
        ;;
esac
RUNSCRIPT_EOF
chmod +x "$INSTALL_DIR/run.sh"

# 如果没有 systemd 或 systemd 服务未启动，使用手动方式
if ! systemctl --version &> /dev/null 2>&1 || ! systemctl --user is-active gowordagent &> /dev/null; then
    "$INSTALL_DIR/run.sh" start
fi

# 验证安装
echo ""
echo -e "${BLUE}验证安装...${NC}"
sleep 2

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
    echo -e "  ${GREEN}服务已启动，端口: $PORT${NC}"

    if command -v curl &> /dev/null; then
        if curl -s "http://127.0.0.1:$PORT/api/proofread/health" > /dev/null 2>&1; then
            echo -e "  ${GREEN}健康检查通过${NC}"
        else
            echo -e "  ${YELLOW}! 健康检查失败，服务可能仍在启动中${NC}"
        fi
    fi
else
    echo -e "  ${YELLOW}! 未检测到端口文件${NC}"
fi

# 安装完成
echo ""
echo -e "${GREEN}"
echo "================================================================"
echo ""
echo "                   安装成功!"
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
if systemctl --user status gowordagent &> /dev/null 2>&1; then
    echo "  查看状态: ${YELLOW}systemctl --user status gowordagent${NC}"
    echo "  停止服务: ${YELLOW}systemctl --user stop gowordagent${NC}"
    echo "  查看日志: ${YELLOW}journalctl --user -u gowordagent -f${NC}"
else
    echo "  查看状态: ${YELLOW}$INSTALL_DIR/run.sh status${NC}"
    echo "  停止服务: ${YELLOW}$INSTALL_DIR/run.sh stop${NC}"
    echo "  查看日志: ${YELLOW}$INSTALL_DIR/run.sh logs${NC}"
fi
echo ""
echo -e "${CYAN}卸载命令:${NC}"
echo "  ${YELLOW}rm -rf $INSTALL_DIR $CONFIG_DIR $WPS_ADDON_DIR${NC}"
echo ""
echo -e "${CYAN}安装目录:${NC} $INSTALL_DIR"
echo ""

# 显示完成对话框
if [ "$HAVE_ZENITY" = true ]; then
    zenity --info --title="安装完成" --text="GOWordAgent 安装成功!\n\n请重启 WPS 文字，在右侧边栏找到'智能校对'面板。" 2>/dev/null
fi

# 等待用户按键
read -p "按回车键关闭..."

exit 0
INSTALLER_EOF

# 计算安装器脚本大小（不包括资源文件）
INSTALLER_SIZE=$(stat -c%s "$OUTPUT_FILE")

# 替换占位符
sed -i "s/INSTALLER_VERSION_PLACEHOLDER/$INSTALLER_VERSION/g" "$OUTPUT_FILE"
sed -i "s/INSTALLER_SIZE_PLACEHOLDER/$INSTALLER_SIZE/g" "$OUTPUT_FILE"

echo "[4/4] 附加资源文件..."
cat "$TEMP_DIR/resources.tar.gz" >> "$OUTPUT_FILE"

# 设置可执行权限
chmod +x "$OUTPUT_FILE"

# 清理
rm -rf "$TEMP_DIR"

echo ""
echo "========================================"
echo "单文件安装器创建成功!"
echo "========================================"
echo ""
echo "输出文件: $OUTPUT_FILE"
echo "文件大小: $(du -h "$OUTPUT_FILE" | cut -f1)"
echo ""
echo "用户使用方法:"
echo "  1. 将 $OUTPUT_FILE 复制到 Linux 系统"
echo "  2. 双击运行，或在终端执行"
echo "     chmod +x $OUTPUT_FILE"
echo "     ./$OUTPUT_FILE"
echo ""
