@echo off
title 小型氧化加工厂财务管理系统 - Web版
echo ==========================================
echo   小型氧化加工厂财务管理系统 V18.0
echo   Web版 - 浏览器访问
echo ==========================================
echo.
echo 正在启动Web服务...
echo.

python --version >nul 2>&1
if errorlevel 1 (
    echo 错误：未检测到Python环境！
    echo 请先安装Python 3.8+
    pause
    exit /b 1
)

pip install flask -q 2>nul

start http://127.0.0.1:5000
python web_app.py
pause
