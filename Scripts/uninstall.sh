#!/bin/bash

echo "=== GOWordAgent 卸载脚本 ==="

# 停止并禁用服务
systemctl --user stop gowordagent 2>/dev/null || true
systemctl --user disable gowordagent 2>/dev/null || true
rm -f $HOME/.config/systemd/user/gowordagent.service
systemctl --user daemon-reload 2>/dev/null || true

# 删除文件
sudo rm -rf /opt/gowordagent
rm -rf $HOME/.local/share/Kingsoft/wps/jsaddons/com.gowordagent.addin

# 询问是否删除配置
echo ""
read -p "是否删除配置文件? [y/N] " -n 1 -r
echo
if [[ $REPLY =~ ^[Yy]$ ]]; then
    rm -rf $HOME/.config/gowordagent
    echo "配置文件已删除"
else
    echo "配置文件保留在: $HOME/.config/gowordagent/"
fi

echo ""
echo "=== 卸载完成 ==="
