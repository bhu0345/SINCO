@echo off
REM 生产交期预测系统打包脚本 (Windows)

echo ================================
echo 生产交期预测系统 - 应用打包工具
echo ================================
echo.

REM 检查是否安装了 pyinstaller
pip show pyinstaller >nul 2>&1
if errorlevel 1 (
    echo 正在安装 PyInstaller...
    pip install pyinstaller
    echo.
)

echo 开始打包应用...
echo.

REM 打包为单个应用程序
pyinstaller --name="生产交期预测系统" --windowed --onedir --clean MVP.py

echo.
echo ================================
echo 打包完成！
echo ================================
echo.
echo 应用位置: dist\生产交期预测系统\
echo.
echo 你可以：
echo   1. 直接双击 生产交期预测系统.exe 运行
echo   2. 将整个文件夹分享给其他人
echo   3. 其他人不需要安装Python就能使用
echo.
pause
