@echo off
title 邮件群发助手启动脚本
echo 正在启动邮件群发助手...

REM 检查Python是否安装
python --version > nul 2>&1
if %errorlevel% equ 0 (
    echo 检测到Python环境，使用Python启动...
    
    REM 检查是否需要安装依赖
    if not exist "venv\" (
        echo 首次运行，正在创建虚拟环境...
        python -m venv venv
        echo 虚拟环境创建完成
        
        echo 正在安装依赖...
        call venv\Scripts\activate
        python -m pip install -r requirements.txt
        echo 依赖安装完成
    ) else (
        echo 使用现有虚拟环境...
        call venv\Scripts\activate
    )
    
    REM 启动Python版本的程序
    python main.py
) else (
    echo 未检测到Python环境，使用绿色版EXE启动...
    
    REM 检查是否存在EXE文件
    if exist "邮件群发助手.exe" (
        start "" "邮件群发助手.exe"
    ) else (
        echo 错误: 未找到邮件群发助手.exe，请确保程序文件完整!
        pause
        exit /b 1
    )
)

exit /b 0 