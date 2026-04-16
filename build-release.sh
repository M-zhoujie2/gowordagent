#!/bin/bash
# GOWordAgent Linux 发布包构建脚本
# 构建可直接部署的发布包

set -e

echo "=== GOWordAgent Linux 发布包构建 ==="
echo ""

# 版本号
VERSION="${VERSION:-1.0.0}"

echo "版本: $VERSION"
echo ""

# 构建函数
build_for_rid() {
    local RID=$1
    local RELEASE_NAME="gowordagent-${RID}-${VERSION}"

    echo "--- 构建 $RID ---"

    # 清理旧构建
    rm -rf "./release/$RELEASE_NAME"
    mkdir -p "./release/$RELEASE_NAME"

    # 构建后端
    echo "构建后端服务 ($RID)..."
    dotnet publish GOWordAgent.WpsService \
      -c Release \
      -r "$RID" \
      --self-contained true \
      -p:PublishSingleFile=true \
      -p:PublishTrimmed=false \
      -o "./release/$RELEASE_NAME/backend"

    echo "后端构建完成 ($RID)"

    # 复制 WPS 加载项
    echo "复制 WPS 加载项..."
    mkdir -p "./release/$RELEASE_NAME/addon"
    cp -r GOWordAgent.WpsAddon/* "./release/$RELEASE_NAME/addon/"

    # 复制脚本和文档
    echo "复制部署脚本..."
    cp deploy-linux.sh "./release/$RELEASE_NAME/"
    cp DEPLOY_LINUX_QUICK.md "./release/$RELEASE_NAME/README.md"
    cp KYLIN_V10_BUILD.md "./release/$RELEASE_NAME/"
    cp TEST_GUIDE.md "./release/$RELEASE_NAME/"
    chmod +x "./release/$RELEASE_NAME/deploy-linux.sh"

    echo "脚本复制完成"

    # 打包
    echo "创建发布包..."
    cd ./release
    tar -czf "$RELEASE_NAME.tar.gz" "$RELEASE_NAME"
    zip -rq "$RELEASE_NAME.zip" "$RELEASE_NAME"
    cd - > /dev/null

    echo "发布包位置:"
    echo "  - release/$RELEASE_NAME.tar.gz"
    echo "  - release/$RELEASE_NAME.zip"
    echo ""
}

# 默认构建当前架构
HOST_ARCH=$(uname -m)
if [ "$HOST_ARCH" = "x86_64" ]; then
    build_for_rid "linux-x64"
    # 如果安装了 arm64 交叉编译工具链，可一并构建
    if dotnet --list-runtimes | grep -q "linux-arm64" 2>/dev/null || dotnet --info | grep -q "arm64"; then
        echo "检测到 arm64 支持，尝试交叉编译..."
        build_for_rid "linux-arm64" || echo "arm64 交叉编译失败，已跳过"
    fi
elif [ "$HOST_ARCH" = "aarch64" ] || [ "$HOST_ARCH" = "arm64" ]; then
    build_for_rid "linux-arm64"
else
    echo "警告: 未识别的架构 $HOST_ARCH，尝试构建 linux-x64"
    build_for_rid "linux-x64"
fi

echo "=== 构建完成 ==="
echo ""
echo "使用说明:"
echo "  1. 解压发布包"
echo "  2. 运行 ./deploy-linux.sh"
echo ""
