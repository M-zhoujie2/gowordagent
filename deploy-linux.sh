#!/bin/bash
# GOWordAgent Linux 快速部署脚本
# 一键部署到银河麒麟 V10 或其他 Linux 发行版

set -e

# 颜色输出
RED='\033[0;31m'
GREEN='\033[0;32m'
YELLOW='\033[1;33m'
NC='\033[0m' # No Color

echo -e "${GREEN}=== GOWordAgent Linux 快速部署脚本 ===${NC}"
echo ""

# 检查架构
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
    echo -e "${RED}错误: 当前仅支持 x86_64 和 arm64 架构，检测到 $ARCH${NC}"
    exit 1
fi

echo -e "${YELLOW}检测到架构: $ARCH ($RID)${NC}"

# 设置安装目录
INSTALL_DIR="${INSTALL_DIR:-$HOME/.local/opt/gowordagent}"
CONFIG_DIR="$HOME/.config/gowordagent"
SERVICE_NAME="gowordagent"

echo -e "${YELLOW}安装目录: $INSTALL_DIR${NC}"
echo -e "${YELLOW}配置目录: $CONFIG_DIR${NC}"
echo ""

# 检查依赖
echo "检查依赖..."
MISSING_DEPS=""

if ! command -v dotnet &> /dev/null; then
    # 检查是否已安装自包含版本
    if [ ! -f "$INSTALL_DIR/gowordagent-server" ]; then
        MISSING_DEPS="dotnet-runtime $MISSING_DEPS"
    fi
fi

if [ -n "$MISSING_DEPS" ]; then
    echo -e "${YELLOW}缺少依赖: $MISSING_DEPS${NC}"
    echo "正在尝试安装..."

    # 检测包管理器并安装
    if command -v apt-get &> /dev/null; then
        # Debian/Ubuntu/银河麒麟
        sudo apt-get update
        sudo apt-get install -y dotnet-runtime-8.0 || {
            echo -e "${YELLOW}.NET 8 运行时安装失败，将使用自包含版本${NC}"
        }
    elif command -v yum &> /dev/null; then
        # RHEL/CentOS
        sudo yum install -y dotnet-runtime-8.0 || {
            echo -e "${YELLOW}.NET 8 运行时安装失败，将使用自包含版本${NC}"
        }
    else
        echo -e "${YELLOW}未知的包管理器，将使用自包含版本${NC}"
    fi
fi

# 创建目录
echo "创建安装目录..."
mkdir -p "$INSTALL_DIR"
mkdir -p "$CONFIG_DIR"

# 复制后端文件
echo "复制后端文件..."
SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"

BACKEND_DIR="$SCRIPT_DIR/backend"
if [ "$RID" != "linux-x64" ] && [ -d "$SCRIPT_DIR/backend-$RID" ]; then
    BACKEND_DIR="$SCRIPT_DIR/backend-$RID"
fi

if [ -d "$BACKEND_DIR" ]; then
    cp -r "$BACKEND_DIR/"* "$INSTALL_DIR/"
elif [ -d "$SCRIPT_DIR/GOWordAgent.WpsService/bin/Release/net8.0/$RID/publish" ]; then
    cp -r "$SCRIPT_DIR/GOWordAgent.WpsService/bin/Release/net8.0/$RID/publish/"* "$INSTALL_DIR/"
else
    echo -e "${RED}错误: 找不到后端文件${NC}"
    echo "请先构建项目: dotnet publish GOWordAgent.WpsService -c Release -r $RID --self-contained"
    exit 1
fi

# 设置执行权限
chmod +x "$INSTALL_DIR/gowordagent-server"

# 创建启动脚本
echo "创建启动脚本..."
cat > "$INSTALL_DIR/start.sh" << 'EOF'
#!/bin/bash
cd "$(dirname "$0")"
export DOTNET_CLI_TELEMETRY_OPTOUT=1
export DOTNET_USE_POLLING_FILE_WATCHER=true
./gowordagent-server "$@"
EOF
chmod +x "$INSTALL_DIR/start.sh"

# 创建 systemd 用户服务文件（如果可用）
if systemctl --version &> /dev/null; then
    echo "创建 systemd 服务..."
    mkdir -p "$HOME/.config/systemd/user"

    cat > "$HOME/.config/systemd/user/$SERVICE_NAME.service" << EOF
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

    # 重载 systemd
    systemctl --user daemon-reload

    # 启动服务
    echo "启动服务..."
    systemctl --user stop "$SERVICE_NAME" 2>/dev/null || true
    systemctl --user start "$SERVICE_NAME"
    systemctl --user enable "$SERVICE_NAME"

    echo ""
    echo -e "${GREEN}服务管理命令:${NC}"
    echo "  启动: systemctl --user start $SERVICE_NAME"
    echo "  停止: systemctl --user stop $SERVICE_NAME"
    echo "  状态: systemctl --user status $SERVICE_NAME"
    echo "  日志: journalctl --user -u $SERVICE_NAME -f"
else
    # 创建简单的启动/停止脚本
    echo "创建手动启动脚本..."

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

    # 启动服务
    "$INSTALL_DIR/run.sh" start
fi

# 探测 WPS 加载项目录
WPS_ADDON_DIR=""
probe_wps_addon_dirs() {
    local dirs=(
        "$HOME/.local/share/Kingsoft/wps/jsaddons/com.gowordagent.addin"
        "$HOME/.config/Kingsoft/wps/jsaddons/com.gowordagent.addin"
        "/usr/share/wps/office/jsaddons/com.gowordagent.addin"
        "/opt/kingsoft/wps-office/office6/jsaddons/com.gowordagent.addin"
    )
    for d in "${dirs[@]}"; do
        if [ -d "$(dirname "$d")" ]; then
            echo "$d"
            return
        fi
    done
    echo "$HOME/.local/share/Kingsoft/wps/jsaddons/com.gowordagent.addin"
}

# 安装 WPS 加载项
echo ""
echo "安装 WPS 加载项..."
WPS_ADDON_DIR=$(probe_wps_addon_dirs)
mkdir -p "$WPS_ADDON_DIR"

if [ -d "$SCRIPT_DIR/addon" ]; then
    cp -r "$SCRIPT_DIR/addon/"* "$WPS_ADDON_DIR/"
    echo -e "${GREEN}加载项已安装到: $WPS_ADDON_DIR${NC}"
elif [ -d "$SCRIPT_DIR/GOWordAgent.WpsAddon" ]; then
    cp -r "$SCRIPT_DIR/GOWordAgent.WpsAddon/"* "$WPS_ADDON_DIR/"
    echo -e "${GREEN}加载项已安装到: $WPS_ADDON_DIR${NC}"
else
    echo -e "${YELLOW}警告: 找不到 WPS 加载项文件${NC}"
fi

# 等待服务启动
echo ""
echo "等待服务启动..."
sleep 2

PORT=""
RUNTIME_PORT_FILE="${XDG_RUNTIME_DIR:-}/gowordagent-port-$USER.json"
LEGACY_PORT_FILE="/tmp/gowordagent-port-$USER.json"
CONFIG_PORT_FILE="$CONFIG_DIR/service-port.json"

if [ -f "$RUNTIME_PORT_FILE" ]; then
    PORT=$(cat "$RUNTIME_PORT_FILE" | grep -o '"Port":[0-9]*' | cut -d: -f2)
elif [ -f "$LEGACY_PORT_FILE" ]; then
    PORT=$(cat "$LEGACY_PORT_FILE" | grep -o '"Port":[0-9]*' | cut -d: -f2)
elif [ -f "$CONFIG_PORT_FILE" ]; then
    PORT=$(cat "$CONFIG_PORT_FILE" | grep -o '"Port":[0-9]*' | cut -d: -f2)
fi

if [ -n "$PORT" ]; then
    echo -e "${GREEN}服务已启动，端口: $PORT${NC}"

    # 测试健康检查
    if curl -s "http://127.0.0.1:$PORT/api/proofread/health" > /dev/null 2>&1; then
        echo -e "${GREEN}健康检查通过${NC}"
    else
        echo -e "${YELLOW}! 健康检查失败，服务可能仍在启动中${NC}"
    fi
else
    echo -e "${YELLOW}! 未检测到端口文件，服务可能启动失败${NC}"
fi

echo ""
echo -e "${GREEN}=== 部署完成 ===${NC}"
echo ""
echo "安装目录: $INSTALL_DIR"
echo "配置目录: $CONFIG_DIR"
echo ""
echo "使用说明:"
echo "  1. 重启 WPS 文字，在右侧边栏找到'智能校对'"
echo "  2. 在插件中配置 AI 提供商和 API Key"
echo "  3. 打开文档，点击'开始校对'"
echo ""
echo "卸载命令:"
echo "  rm -rf $INSTALL_DIR $CONFIG_DIR $WPS_ADDON_DIR"
if systemctl --version &> /dev/null; then
    echo "  systemctl --user disable $SERVICE_NAME"
fi
echo ""
