@echo off
chcp 65001 >nul
echo ==========================================
echo   会议座位管理系统 - Windows安装脚本
echo   版本: v1.2.0
echo ==========================================
echo.

REM 检查管理员权限
net session >nul 2>&1
if %errorLevel% neq 0 (
    echo [提示] 建议以管理员身份运行此脚本
    echo.
)

REM 检查Node.js
echo [1/4] 检查 Node.js...
node --version >nul 2>&1
if %errorLevel% neq 0 (
    echo [错误] 未检测到 Node.js
    echo 请从 https://nodejs.org/ 下载并安装 Node.js
    pause
    exit /b 1
)
echo [OK] Node.js 已安装
node --version
echo.

REM 检查Python
echo [2/4] 检查 Python...
python --version >nul 2>&1
if %errorLevel% neq 0 (
    python3 --version >nul 2>&1
    if %errorLevel% neq 0 (
        echo [错误] 未检测到 Python
        echo 请从 https://www.python.org/downloads/ 下载并安装 Python
        pause
        exit /b 1
    )
)
echo [OK] Python 已安装
python --version 2>nul || python3 --version
echo.

REM 创建目录
echo [3/4] 创建必要目录...
if not exist "data" mkdir data
if not exist "data\backups" mkdir data\backups
if not exist "output" mkdir output
if not exist "uploads" mkdir uploads
if not exist "templates" mkdir templates
if not exist "fonts" mkdir fonts
echo [OK] 目录创建完成
echo.

REM 安装依赖
echo [4/4] 安装项目依赖...
echo.
echo 正在安装 Node.js 依赖...
call npm install
if %errorLevel% neq 0 (
    echo [警告] Node.js 依赖安装可能遇到问题，请手动运行 npm install
)
echo.
echo 正在安装 Python 依赖...
pip install -r requirements.txt
if %errorLevel% neq 0 (
    pip3 install -r requirements.txt
)
echo.

echo ==========================================
echo   安装完成！
echo ==========================================
echo.
echo 下一步操作：
echo.
echo 1. 启动服务（需要两个命令行窗口）：
echo.
echo    窗口1 - 座位查询系统：
echo    node server.js
echo.
echo    窗口2 - 座位牌生成系统：
echo    streamlit run app.py --server.port 8505 --server.address 0.0.0.0
echo.
echo 2. 访问系统：
echo    - 座位查询：http://localhost:3000
echo    - 管理后台：http://localhost:3000/admin.html
echo    - 座位牌生成：http://localhost:8505
echo.
echo 3. 默认账号：
echo    - 管理员：admin / admin888
echo    - 会议举办方：organizer / organizer888
echo.
echo ⚠️  重要：首次登录后请立即修改默认密码！
echo.
echo 详细文档请查看：
echo   - README.md  使用说明
echo   - DEPLOY.md  部署指南
echo.
pause
