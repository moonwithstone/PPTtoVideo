#!/bin/bash
# ==============================================
#  macOS 打包脚本
#  生成 PPTtoVideo.app 和 PPTtoVideo.dmg
# ==============================================

set -e

echo "=============================="
echo "  PPT转视频 - macOS 打包"
echo "=============================="

cd "$(dirname "$0")"

# 激活虚拟环境
if [ ! -d "venv" ]; then
    echo "创建虚拟环境..."
    python3 -m venv venv
fi
source venv/bin/activate

# 安装依赖
echo "安装依赖..."
pip install -q -r requirements.txt
pip install -q pywebview pyinstaller

# 清理旧的构建
echo "清理旧构建..."
rm -rf build dist

# 执行打包
echo "开始打包..."
pyinstaller ppttovideo.spec --noconfirm

echo ""

# 检查结果
if [ -d "dist/PPTtoVideo.app" ]; then
    echo "✅ 打包成功！"
    echo "   应用位置: dist/PPTtoVideo.app"
    
    # 生成 DMG
    echo ""
    echo "正在生成 DMG 安装包..."
    
    DMG_NAME="PPTtoVideo-macOS.dmg"
    DMG_DIR="dist/dmg_temp"
    
    rm -rf "$DMG_DIR"
    mkdir -p "$DMG_DIR"
    cp -R "dist/PPTtoVideo.app" "$DMG_DIR/"
    
    # 创建 Applications 快捷方式
    ln -s /Applications "$DMG_DIR/Applications"
    
    # 生成 DMG
    rm -f "dist/$DMG_NAME"
    hdiutil create -volname "PPT转视频" \
        -srcfolder "$DMG_DIR" \
        -ov -format UDZO \
        "dist/$DMG_NAME"
    
    rm -rf "$DMG_DIR"
    
    echo ""
    echo "✅ DMG 生成成功！"
    echo "   文件: dist/$DMG_NAME"
    echo ""
    echo "安装方式：打开 DMG，将 PPTtoVideo 拖入 Applications 文件夹"
else
    echo "❌ 打包失败，请检查错误信息"
    exit 1
fi
