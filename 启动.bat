@echo off
chcp 65001 >nul
title 电力收益分析系统

echo.
echo ========================================
echo           电力收益分析系统
echo ========================================
echo.
echo 请选择要运行的程序：
echo.
echo 1. 日前结算收益分析器 (interactive_analyzer.py)
echo 2. 原合约优化分析器 (contract_optimizer.py)
echo 3. 安装依赖包
echo 0. 退出
echo.
echo ========================================

set /p choice=请输入选项 (0-3): 

if "%choice%"=="1" (
    echo.
    echo 启动日前结算收益分析器...
    py -3 interactive_analyzer.py
    pause
) else if "%choice%"=="2" (
    echo.
    echo 启动原合约优化分析器...
    py -3 contract_optimizer.py
    pause
) else if "%choice%"=="3" (
    echo.
    echo 正在安装依赖包...
    py -3 -m pip install --user -r requirements.txt
    echo.
    echo 依赖包安装完成！
    pause
) else if "%choice%"=="0" (
    echo.
    echo 谢谢使用！
    exit
) else (
    echo.
    echo 无效选项，请重新运行程序！
    pause
)