@echo off
chcp 65001 >nul
echo ==========================================
echo   会议座位管理系统 - 打包脚本
echo   版本: v1.2.0
echo ==========================================
echo.

set VERSION=v1.3.0
set PACKAGE_NAME=meeting-system-%VERSION%

echo [准备] 清理临时文件...
if exist "%PACKAGE_NAME%.zip" del /f /q "%PACKAGE_NAME%.zip"
if exist "%PACKAGE_NAME%" rmdir /s /q "%PACKAGE_NAME%"

echo.
echo [1/3] 创建打包目录...
mkdir "%PACKAGE_NAME%"

echo.
echo [2/3] 复制项目文件...

REM 核心文件
xcopy /y /i app.py "%PACKAGE_NAME%\" >nul
xcopy /y /i server.js "%PACKAGE_NAME%\" >nul
xcopy /y /i parse-excel.js "%PACKAGE_NAME%\" >nul
xcopy /y /i package.json "%PACKAGE_NAME%\" >nul
xcopy /y /i package-lock.json "%PACKAGE_NAME%\" >nul
xcopy /y /i requirements.txt "%PACKAGE_NAME%\" >nul
xcopy /y /i config.json "%PACKAGE_NAME%\" >nul
xcopy /y /i Dockerfile "%PACKAGE_NAME%\" >nul
xcopy /y /i docker-compose.yml "%PACKAGE_NAME%\" >nul
xcopy /y /i supervisord.conf "%PACKAGE_NAME%\" >nul
xcopy /y /i nginx.conf "%PACKAGE_NAME%\" >nul
xcopy /y /i .dockerignore "%PACKAGE_NAME%\" >nul
xcopy /y /i .gitignore "%PACKAGE_NAME%\" >nul

REM 文档
xcopy /y /i README.md "%PACKAGE_NAME%\" >nul
xcopy /y /i DEPLOY.md "%PACKAGE_NAME%\" >nul
xcopy /y /i INSTALL.md "%PACKAGE_NAME%\" >nul

REM 安装脚本
xcopy /y /i install.sh "%PACKAGE_NAME%\" >nul
xcopy /y /i install.bat "%PACKAGE_NAME%\" >nul

REM 前端文件
xcopy /y /e /i public "%PACKAGE_NAME%\public\" >nul

REM Streamlit配置
xcopy /y /e /i .streamlit "%PACKAGE_NAME%\.streamlit\" >nul

REM 字体文件
xcopy /y /e /i fonts "%PACKAGE_NAME%\fonts\" >nul

REM 创建空目录
mkdir "%PACKAGE_NAME%\data" >nul 2>&1
mkdir "%PACKAGE_NAME%\data\backups" >nul 2>&1
mkdir "%PACKAGE_NAME%\output" >nul 2>&1
mkdir "%PACKAGE_NAME%\uploads" >nul 2>&1
mkdir "%PACKAGE_NAME%\templates" >nul 2>&1

REM 创建空的占位文件
echo. > "%PACKAGE_NAME%\data\.gitkeep"
echo. > "%PACKAGE_NAME%\output\.gitkeep"
echo. > "%PACKAGE_NAME%\uploads\.gitkeep"
echo. > "%PACKAGE_NAME%\templates\.gitkeep"

echo.
echo [3/3] 创建压缩包...
powershell -Command "Compress-Archive -Path '%PACKAGE_NAME%\*' -DestinationPath '%PACKAGE_NAME%.zip' -Force"

echo.
echo ==========================================
echo   打包完成！
echo ==========================================
echo.
echo 生成的文件：
echo   - %PACKAGE_NAME%.zip  (Windows/Linux通用)
echo.
echo 文件大小：
for %%I in ("%PACKAGE_NAME%.zip") do echo   - %%~zI 字节
echo.
echo 清理临时目录...
rmdir /s /q "%PACKAGE_NAME%"
echo.
echo 打包任务完成！
pause
