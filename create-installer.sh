#!/bin/bash
# GOWordAgent Linux 单文件安装程序生成脚本
# 生成可直接双击运行的安装程序

set -e

echo "=== GOWordAgent 单文件安装程序生成器 ==="
echo ""

VERSION="${VERSION:-1.0.0}"
INSTALLER_NAME="GOWordAgent-Installer-linux-x64-${VERSION}.run"

echo "版本: $VERSION"
echo "安装程序名称: $INSTALLER_NAME"
echo ""

# 检查是否已构建
if [ ! -d "./release/gowordagent-linux-x64-${VERSION}/backend" ]; then
    echo "未找到构建文件，先执行构建..."
    ./build-release.sh
fi

echo "创建安装程序..."

# 创建临时目录
TMP_DIR=$(mktemp -d)
trap "rm -rf $TMP_DIR" EXIT

# 复制文件到临时目录
cp -r ./release/gowordagent-linux-x64-${VERSION}/* "$TMP_DIR/"

# 创建安装脚本头
cat > "$TMP_DIR/installer-header.sh" << 'INSTALLER_EOF'
#!/bin/bash
# GOWordAgent Linux 安装程序
# 自动解压并执行安装

set -e

# 颜色输出
RED='\033[0;31m'
GREEN='\033[0;32m'
YELLOW='\033[1;33m'
BLUE='\033[0;34m'
NC='\033[0m'

echo -e "${GREEN}============================================================${NC}"
echo -e "${GREEN}     GOWordAgent 智能校对 - Linux 安装程序                 ${NC}"
echo -e "${GREEN}============================================================${NC}"
echo ""

# 检查架构
ARCH=$(uname -m)
if [ "$ARCH" != "x86_64" ] && [ "$ARCH" != "aarch64" ] && [ "$ARCH" != "arm64" ]; then
    echo -e "${RED}错误: 当前仅支持 x86_64 和 arm64 架构，检测到 $ARCH${NC}"
    read -p "按回车键退出..."
    exit 1
fi

# 显示安装信息
echo -e "${BLUE}安装信息:${NC}"
echo "  版本: INSTALLER_VERSION"
echo "  安装目录: ~/.local/opt/gowordagent"
echo "  配置目录: ~/.config/gowordagent"
echo ""

# 确认安装
read -p "是否继续安装? [Y/n] " -n 1 -r
echo
if [[ ! $REPLY =~ ^[Yy]$ ]] && [ -n "$REPLY" ]; then
    echo "安装已取消"
    exit 0
fi

echo ""
echo -e "${YELLOW}正在安装...${NC}"
echo ""

# 创建临时目录解压
TMP_DIR=$(mktemp -d)
trap "rm -rf $TMP_DIR" EXIT

# 解压数据
echo "解压安装文件..."
sed -n '/^__DATA_START__$/,$p' "$0" | tail -n +2 | tar -xz -C "$TMP_DIR" 2>/dev/null || {
    echo -e "${RED}错误: 解压失败${NC}"
    read -p "按回车键退出..."
    exit 1
}

# 执行安装脚本
cd "$TMP_DIR"
bash ./deploy-linux.sh

# 安装完成
echo ""
echo -e "${GREEN}============================================================${NC}"
echo -e "${GREEN}              安装完成!                                    ${NC}"
echo -e "${GREEN}============================================================${NC}"
echo ""
echo "使用说明:"
echo "  1. 重启 WPS 文字"
echo "  2. 在右侧边栏找到'智能校对'面板"
echo "  3. 配置 AI 提供商和 API Key"
echo "  4. 打开文档，点击'开始校对'"
echo ""
read -p "按回车键退出..."

exit 0

__DATA_START__
INSTALLER_EOF

# 替换版本号
sed -i "s/INSTALLER_VERSION/$VERSION/g" "$TMP_DIR/installer-header.sh"

# 创建最终的安装程序
echo "打包安装程序..."
cat "$TMP_DIR/installer-header.sh" > "$INSTALLER_NAME"
(cd "$TMP_DIR" && tar -czf - . >> "$OLDPWD/$INSTALLER_NAME")
chmod +x "$INSTALLER_NAME"

echo ""
echo -e "${GREEN}安装程序已生成: $INSTALLER_NAME${NC}"
echo ""
echo "使用方式:"
echo "  1. 在文件管理器中双击 $INSTALLER_NAME"
echo "  2. 或在终端运行: ./$INSTALLER_NAME"
echo ""
echo "文件大小: $(du -h "$INSTALLER_NAME" | cut -f1)"
echo ""
