@echo off
title 邮件群发助手
echo 正在启动邮件群发助手...

REM 直接启动EXE文件
if exist "邮件群发助手.exe" (
    start "" "邮件群发助手.exe"
) else (
    echo 未找到邮件群发助手.exe，尝试运行setup.bat...
    if exist "setup.bat" (
        call setup.bat
    ) else (
        echo 错误: 未找到启动文件，请确保程序完整!
        pause
    )
)

exit 