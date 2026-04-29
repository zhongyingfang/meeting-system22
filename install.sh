#!/bin/bash

# 会议座位管理系统 - Linux安装脚本
# 版本: v1.2.0

set -e

echo "=========================================="
echo "  会议座位管理系统 - 自动安装脚本"
echo "  版本: v1.2.0"
echo "=========================================="
echo ""

# 检查是否为root用户
if [ "$EUID" -ne 0 ]; then 
    echo "请使用 root 权限运行此脚本"
    echo "使用: sudo $0"
    exit 1
fi

# 检测系统发行版
if [ -f /etc/os-release ]; then
    . /etc/os-release
    OS=$ID
else
    echo "无法检测操作系统类型"
    exit 1
fi

echo "检测到操作系统: $OS"
echo ""

# 安装系统依赖
echo "正在安装系统依赖..."
if [ "$OS" = "ubuntu" ] || [ "$OS" = "debian" ]; then
    apt update
    apt install -y nodejs npm python3 python3-pip \
        fonts-noto-cjk fonts-wqy-zenhei fonts-wqy-microhei \
        poppler-utils supervisor
elif [ "$OS" = "centos" ] || [ "$OS" = "rhel" ]; then
    yum install -y nodejs npm python3 python3-pip supervisor
    # 安装中文字体
    yum install -y google-noto-cjk-fonts wqy-zenhei-fonts wqy-microhei-fonts || true
else
    echo "不支持的操作系统: $OS"
    exit 1
fi

echo "系统依赖安装完成！"
echo ""

# 获取当前目录
INSTALL_DIR=$(pwd)
echo "安装目录: $INSTALL_DIR"

# 创建必要的目录
echo "正在创建目录..."
mkdir -p "$INSTALL_DIR/data"
mkdir -p "$INSTALL_DIR/output"
mkdir -p "$INSTALL_DIR/uploads"
mkdir -p "$INSTALL_DIR/fonts"
mkdir -p "$INSTALL_DIR/templates"
mkdir -p "$INSTALL_DIR/data/backups"

# 设置权限
echo "正在设置目录权限..."
chmod -R 755 "$INSTALL_DIR"
chmod -R 777 "$INSTALL_DIR/data"
chmod -R 777 "$INSTALL_DIR/output"
chmod -R 777 "$INSTALL_DIR/uploads"
if [ -d "$INSTALL_DIR/fonts" ]; then
    chmod -R 644 "$INSTALL_DIR/fonts"/* 2>/dev/null || true
fi

echo "目录权限设置完成！"
echo ""

# 安装Node.js依赖
echo "正在安装 Node.js 依赖..."
cd "$INSTALL_DIR"
npm install --production
echo "Node.js 依赖安装完成！"
echo ""

# 安装Python依赖
echo "正在安装 Python 依赖..."
pip3 install -r requirements.txt
echo "Python 依赖安装完成！"
echo ""

# 安装supervisor
echo "正在配置服务管理..."
if ! command -v supervisord &> /dev/null; then
    pip3 install supervisor
fi

echo ""
echo "=========================================="
echo "  安装完成！"
echo "=========================================="
echo ""
echo "下一步操作："
echo ""
echo "1. 启动服务："
echo "   cd $INSTALL_DIR"
echo "   supervisord -c supervisord.conf"
echo ""
echo "2. 或使用 systemd 服务（推荐）："
echo "   查看 DEPLOY.md 了解如何配置 systemd"
echo ""
echo "3. 访问系统："
echo "   - 座位查询：http://$(hostname -I | awk '{print $1}'):3000"
echo "   - 管理后台：http://$(hostname -I | awk '{print $1}'):3000/admin.html"
echo "   - 座位牌生成：http://$(hostname -I | awk '{print $1}'):8505"
echo ""
echo "4. 默认账号："
echo "   - 管理员：admin / admin888"
echo "   - 会议举办方：organizer / organizer888"
echo ""
echo "⚠️  重要：首次登录后请立即修改默认密码！"
echo ""
echo "详细文档请查看："
echo "  - README.md  (使用说明)"
echo "  - DEPLOY.md  (部署指南)"
echo ""
