@echo off
chcp 65001 >nul
title 小型氧化加工厂财务管理系统
echo ==========================================
echo   小型氧化加工厂财务管理系统 V16.0
echo   Python独立运行版
echo ==========================================
echo.
echo 正在启动...
echo.

REM 检查Python是否安装
python --version >nul 2>&1
if errorlevel 1 (
    echo 错误：未检测到Python环境！
    echo.
    echo 请先安装Python：
    echo 1. 访问 https://www.python.org/downloads/
    echo 2. 下载并安装Python 3.8或更高版本
    echo 3. 安装时勾选 "Add Python to PATH"
    echo.
    pause
    exit /b 1
)

REM 检查依赖库
python -c "import openpyxl" >nul 2>&1
if errorlevel 1 (
    echo 正在安装依赖库...
    pip install openpyxl -q
)

REM 运行程序
python "%~dp0氧化加工厂财务系统.py"

pause
