# 会议座位管理系统 - 打包脚本
# 版本: v1.2.0

$VERSION = "v1.2.0"
$PACKAGE_NAME = "meeting-system-$VERSION"

Write-Host "==========================================" -ForegroundColor Cyan
Write-Host "  会议座位管理系统 - 打包脚本" -ForegroundColor Cyan
Write-Host "  版本: $VERSION" -ForegroundColor Cyan
Write-Host "==========================================" -ForegroundColor Cyan
Write-Host ""

# 清理旧文件
Write-Host "[1/5] 清理旧文件..." -ForegroundColor Yellow
if (Test-Path $PACKAGE_NAME) {
    Remove-Item -Recurse -Force $PACKAGE_NAME -ErrorAction SilentlyContinue
}
if (Test-Path "$PACKAGE_NAME.zip") {
    Remove-Item -Force "$PACKAGE_NAME.zip" -ErrorAction SilentlyContinue
}
Write-Host "      清理完成" -ForegroundColor Green

# 创建打包目录
Write-Host "[2/5] 创建打包目录..." -ForegroundColor Yellow
New-Item -ItemType Directory -Path $PACKAGE_NAME | Out-Null
Write-Host "      目录创建完成" -ForegroundColor Green

# 复制核心文件
Write-Host "[3/5] 复制项目文件..." -ForegroundColor Yellow

$coreFiles = @(
    "app.py",
    "server.js",
    "parse-excel.js",
    "package.json",
    "package-lock.json",
    "requirements.txt",
    "config.json",
    "Dockerfile",
    "docker-compose.yml",
    "supervisord.conf",
    "nginx.conf",
    ".dockerignore",
    "README.md",
    "DEPLOY.md",
    "INSTALL.md",
    "install.sh",
    "install.bat"
)

foreach ($file in $coreFiles) {
    if (Test-Path $file) {
        Copy-Item $file -Destination "$PACKAGE_NAME\" -Force
        Write-Host "      已复制: $file" -ForegroundColor Gray
    }
}

# 复制目录
$directories = @("public", ".streamlit", "fonts")
foreach ($dir in $directories) {
    if (Test-Path $dir) {
        Copy-Item $dir -Recurse -Destination "$PACKAGE_NAME\" -Force
        Write-Host "      已复制: $dir/" -ForegroundColor Gray
    }
}

# 创建空目录
Write-Host "[4/5] 创建空目录结构..." -ForegroundColor Yellow
$emptyDirs = @("data", "data/backups", "output", "uploads", "templates")
foreach ($dir in $emptyDirs) {
    $fullPath = Join-Path $PACKAGE_NAME $dir
    New-Item -ItemType Directory -Path $fullPath -Force | Out-Null
    $gitkeep = Join-Path $fullPath ".gitkeep"
    "" | Out-File -FilePath $gitkeep -Encoding utf8
}
Write-Host "      目录结构创建完成" -ForegroundColor Green

# 创建压缩包
Write-Host "[5/5] 创建压缩包..." -ForegroundColor Yellow
Compress-Archive -Path "$PACKAGE_NAME/*" -DestinationPath "$PACKAGE_NAME.zip" -Force
Write-Host "      压缩包创建完成" -ForegroundColor Green

# 清理临时目录
Remove-Item -Recurse -Force $PACKAGE_NAME

Write-Host ""
Write-Host "==========================================" -ForegroundColor Cyan
Write-Host "  打包完成！" -ForegroundColor Green
Write-Host "==========================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "生成的文件：" -ForegroundColor White
Write-Host "  - $PACKAGE_NAME.zip" -ForegroundColor Cyan

if (Test-Path "$PACKAGE_NAME.zip") {
    $size = (Get-Item "$PACKAGE_NAME.zip").Length
    $sizeKB = [math]::Round($size / 1KB, 2)
    $sizeMB = [math]::Round($size / 1MB, 2)
    Write-Host ""
    Write-Host "文件大小：" -ForegroundColor White
    Write-Host "  - $sizeKB KB ($sizeMB MB)" -ForegroundColor Gray
}

Write-Host ""
Write-Host "部署说明：" -ForegroundColor White
Write-Host "  - Windows: 解压后运行 install.bat" -ForegroundColor Gray
Write-Host "  - Linux:   解压后运行 sudo ./install.sh" -ForegroundColor Gray
Write-Host "  - Docker:  解压后运行 docker compose up -d --build" -ForegroundColor Gray
Write-Host ""
Write-Host "详细文档请查看：" -ForegroundColor White
Write-Host "  - README.md  (使用说明)" -ForegroundColor Gray
Write-Host "  - DEPLOY.md  (部署指南)" -ForegroundColor Gray
Write-Host ""
