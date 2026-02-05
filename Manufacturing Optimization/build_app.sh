#!/bin/bash
# 生产交期预测系统打包脚本

echo "================================"
echo "生产交期预测系统 - 应用打包工具"
echo "================================"
echo ""

# 检查是否安装了 pyinstaller
if ! command -v pyinstaller &> /dev/null
then
    echo "正在安装 PyInstaller..."
    pip install pyinstaller
    echo ""
fi

echo "开始打包应用..."
echo ""

# 打包为单个应用程序
# --name: 应用名称
# --windowed: GUI模式（不显示终端窗口）
# --onefile: 打包成单个文件（可选，也可以用 --onedir）
# --icon: 应用图标（如果有的话）
# --add-data: 包含数据文件

pyinstaller --name="生产交期预测系统" \
    --windowed \
    --onedir \
    --clean \
    main.py

echo ""
echo "================================"
echo "打包完成！"
echo "================================"
echo ""
echo "应用位置："
echo "  Mac: dist/生产交期预测系统.app"
echo "  或者: dist/生产交期预测系统/"
echo ""
echo "你可以："
echo "  1. 直接双击运行"
echo "  2. 拖到应用程序文件夹"
echo "  3. 分享给其他人（他们不需要安装Python）"
echo ""
