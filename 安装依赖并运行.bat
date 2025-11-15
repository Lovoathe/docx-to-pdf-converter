@echo off
chcp 65001 > nul

echo ===================================================
echo          Doc/Docx 转 PDF 工具 - 一键启动脚本          
echo ===================================================
echo.
echo 此脚本将：
echo 1. 安装必要的Python依赖
 echo 2. 运行测试程序，验证功能是否正常
echo.

:: 检查Python是否安装
python --version > nul 2>&1
if %errorlevel% neq 0 (
    echo 错误: 未找到Python。请先从 https://www.python.org/downloads/ 安装Python
    echo 安装时请勾选 "Add Python to PATH" 选项
    echo.
    pause
    exit /b 1
)

echo 正在安装必要的Python依赖...
echo.
pip install pywin32 --user

if %errorlevel% neq 0 (
    echo.
    echo 警告: 依赖安装可能失败。请尝试以管理员身份运行此脚本。
    echo 或者手动在命令提示符中运行: pip install pywin32
    echo.
    pause
)

echo.
echo 依赖安装完成，准备启动测试程序...
echo 注意: 请确保您的计算机已安装Microsoft Word
echo.
pause
echo 正在启动测试程序...
python test_conversion.py

:: 等待用户按任意键退出
if %errorlevel% neq 0 (
    echo.
    echo 程序运行过程中遇到错误，请查看上面的输出信息
    echo.
    pause
)