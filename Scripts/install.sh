#!/bin/bash
set -e

INSTALL_DIR=/opt/gowordagent
CONFIG_DIR=$HOME/.config/gowordagent
ADDON_DIR=$HOME/.local/share/Kingsoft/wps/jsaddons/com.gowordagent.addin
SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"

echo "=== GOWordAgent 安装脚本 ==="

# 1. 检查架构
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
    echo "警告: 当前仅官方支持 x86_64 和 arm64 架构，检测到 $ARCH"
    echo "安装将继续，但可能无法正常运行"
    RID="linux-x64"
fi

# 2. 检测 WPS（尝试多种方式）
WPS_CMD=""
for cmd in wps /usr/bin/wps /opt/kingsoft/wps-office/wps /usr/local/bin/wps; do
    if command -v $cmd &> /dev/null; then
        WPS_CMD=$cmd
        break
    fi
done

if [ -z "$WPS_CMD" ]; then
    echo "警告: 未检测到 WPS Office 命令"
    echo "请确保 WPS Office 已安装"
else
    echo "检测到 WPS: $WPS_CMD"
fi

# 3. 检查后端文件
BACKEND_DIR="$SCRIPT_DIR/../backend"
if [ "$RID" != "linux-x64" ] && [ -d "$SCRIPT_DIR/../backend-$RID" ]; then
    BACKEND_DIR="$SCRIPT_DIR/../backend-$RID"
fi

if [ ! -d "$BACKEND_DIR" ]; then
    echo "错误: 找不到后端文件目录 ($BACKEND_DIR)"
    echo "请确保已将编译后的后端文件复制到 backend/ 目录"
    exit 1
fi

# 4. 检查 systemd 用户服务是否可用
SYSTEMD_USER_AVAILABLE=false
if systemctl --user status &>/dev/null 2>&1; then
    SYSTEMD_USER_AVAILABLE=true
    echo "Systemd 用户服务可用"
else
    echo "警告: Systemd 用户服务不可用，将使用手动启动方式"
fi

# 5. 复制后端
echo "正在安装后端服务到 $INSTALL_DIR..."
if [ -d "$INSTALL_DIR" ]; then
    echo "  清理旧版本..."
    sudo rm -rf $INSTALL_DIR/*
else
    sudo mkdir -p $INSTALL_DIR
fi

sudo cp -r "$BACKEND_DIR/"* $INSTALL_DIR/
sudo chmod +x $INSTALL_DIR/gowordagent-server

# 6. 注册 systemd 用户服务（如果可用）
if [ "$SYSTEMD_USER_AVAILABLE" = true ]; then
    echo "正在注册系统服务..."
    mkdir -p $HOME/.config/systemd/user
    cp "$SCRIPT_DIR/gowordagent.service" $HOME/.config/systemd/user/
    systemctl --user daemon-reload
    systemctl --user enable gowordagent

    # 停止可能正在运行的旧服务
    systemctl --user stop gowordagent 2>/dev/null || true

    # 启动服务
    if systemctl --user start gowordagent; then
        echo "服务启动成功"
        sleep 1
        systemctl --user status gowordagent --no-pager || true
    else
        echo "警告: 服务启动失败，请手动检查"
    fi
else
    echo "跳过服务注册（Systemd 不可用）"
    echo "请手动启动服务: $INSTALL_DIR/gowordagent-server"
fi

# 探测 WPS 加载项目录
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

# 7. 安装 WPS 加载项
echo "正在安装 WPS 加载项..."
ADDON_DIR=$(probe_wps_addon_dirs)
mkdir -p "$ADDON_DIR"

if [ -d "$SCRIPT_DIR/../addon" ]; then
    cp -r "$SCRIPT_DIR/../addon/"* "$ADDON_DIR/"
    echo "加载项已安装到: $ADDON_DIR"
else
    echo "警告: 找不到加载项文件 (../addon)"
fi

# 8. 创建配置目录
mkdir -p $CONFIG_DIR
echo "配置目录: $CONFIG_DIR"

# 9. 等待服务启动并检查
if [ "$SYSTEMD_USER_AVAILABLE" = true ]; then
    echo ""
    echo "等待服务启动..."
    sleep 2

    PORT=""
    RUNTIME_PORT_FILE="${XDG_RUNTIME_DIR:-}/gowordagent-port-$USER.json"
    LEGACY_PORT_FILE="/tmp/gowordagent-port-$USER.json"
    CONFIG_PORT_FILE="$CONFIG_DIR/service-port.json"

    if [ -f "$RUNTIME_PORT_FILE" ]; then
        PORT=$(cat "$RUNTIME_PORT_FILE" 2>/dev/null | grep -o '"Port":[0-9]*' | cut -d: -f2)
    elif [ -f "$LEGACY_PORT_FILE" ]; then
        PORT=$(cat "$LEGACY_PORT_FILE" 2>/dev/null | grep -o '"Port":[0-9]*' | cut -d: -f2)
    elif [ -f "$CONFIG_PORT_FILE" ]; then
        PORT=$(cat "$CONFIG_PORT_FILE" 2>/dev/null | grep -o '"Port":[0-9]*' | cut -d: -f2)
    fi

    if [ -n "$PORT" ]; then
        echo "服务已启动，端口: $PORT"
    else
        echo "警告: 未检测到端口文件，服务可能未正常启动"
    fi
fi

echo ""
echo "=== 安装完成 ==="
echo ""

if [ "$SYSTEMD_USER_AVAILABLE" = true ]; then
    echo "服务状态:"
    systemctl --user status gowordagent --no-pager 2>/dev/null || true
    echo ""
    echo "管理服务命令:"
    echo "  查看状态: systemctl --user status gowordagent"
    echo "  启动服务: systemctl --user start gowordagent"
    echo "  停止服务: systemctl --user stop gowordagent"
    echo "  查看日志: journalctl --user -u gowordagent -f"
else
    echo "手动启动命令:"
    echo "  $INSTALL_DIR/gowordagent-server"
fi

echo ""
echo "请重启 WPS 文字以加载插件"
echo ""
echo "卸载命令: ./uninstall.sh"
