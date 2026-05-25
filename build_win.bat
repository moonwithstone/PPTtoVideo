@echo off
REM ==============================================
REM  Windows 打包脚本
REM  生成 PPTtoVideo.exe 安装包
REM ==============================================

echo ==============================
echo   PPT转视频 - Windows 打包
echo ==============================

cd /d "%~dp0"

REM 创建虚拟环境
if not exist "venv" (
    echo 创建虚拟环境...
    python -m venv venv
)
call venv\Scripts\activate.bat

REM 安装依赖
echo 安装依赖...
pip install -q -r requirements.txt
pip install -q pywebview pyinstaller

REM 清理旧的构建
echo 清理旧构建...
if exist "build" rmdir /s /q build
if exist "dist" rmdir /s /q dist

REM 执行打包
echo 开始打包...
pyinstaller ppttovideo.spec --noconfirm

echo.

REM 检查结果
if exist "dist\PPTtoVideo\PPTtoVideo.exe" (
    echo ✅ 打包成功！
    echo    应用位置: dist\PPTtoVideo\PPTtoVideo.exe
    echo.
    echo 你可以将 dist\PPTtoVideo 文件夹整个复制给用户
    echo 或使用 Inno Setup 等工具制作安装包
) else (
    echo ❌ 打包失败，请检查错误信息
    exit /b 1
)

pause
