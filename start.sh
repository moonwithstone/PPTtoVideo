#!/bin/bash
# PPT 转视频 - 一键启动脚本

echo "=============================="
echo "  PPT 转视频工具 - Web 版"
echo "=============================="

# 进入脚本所在目录
cd "$(dirname "$0")"

# 检查 Python
if ! command -v python3 &> /dev/null; then
    echo "错误：未找到 python3，请先安装 Python 3"
    exit 1
fi

# 创建/激活虚拟环境
if [ ! -d "venv" ]; then
    echo "创建虚拟环境..."
    python3 -m venv venv
fi
source venv/bin/activate

# 检查并安装依赖
echo "检查 Python 依赖..."
pip install -q -r requirements.txt

# 检查系统依赖
echo "检查系统依赖..."
if ! command -v soffice &> /dev/null; then
    echo "⚠ 警告：未找到 LibreOffice"
    echo "  请安装：brew install --cask libreoffice"
fi

if ! command -v pdftoppm &> /dev/null; then
    echo "⚠ 警告：未找到 poppler"
    echo "  请安装：brew install poppler"
fi

echo ""
echo "启动 Web 服务..."
echo "浏览器访问 http://127.0.0.1:5000"
echo ""

python app.py
