const fs = require('fs');
const path = require('path');
const { execSync } = require('child_process');

const VERSION = 'v1.7.6';
const PACKAGE_NAME = `meeting-system-${VERSION}`;

console.log('==========================================');
console.log('  会议座位管理系统 - 打包脚本');
console.log(`  版本: ${VERSION}`);
console.log('==========================================');
console.log('');

// 复制目录函数
function copyDir(src, dest) {
    fs.mkdirSync(dest, { recursive: true });
    const entries = fs.readdirSync(src, { withFileTypes: true });
    for (const entry of entries) {
        const srcPath = path.join(src, entry.name);
        const destPath = path.join(dest, entry.name);
        if (entry.isDirectory()) {
            copyDir(srcPath, destPath);
        } else {
            fs.copyFileSync(srcPath, destPath);
        }
    }
}

// 清理旧文件
console.log('[1/5] 清理旧文件...');
if (fs.existsSync(PACKAGE_NAME)) {
    fs.rmSync(PACKAGE_NAME, { recursive: true, force: true });
}
if (fs.existsSync(`${PACKAGE_NAME}.zip`)) {
    fs.unlinkSync(`${PACKAGE_NAME}.zip`);
}
console.log('      清理完成');

// 创建打包目录
console.log('[2/5] 创建打包目录...');
fs.mkdirSync(PACKAGE_NAME, { recursive: true });
console.log('      目录创建完成');

// 复制核心文件
console.log('[3/5] 复制项目文件...');

const coreFiles = [
    'app.py',
    'server.js',
    'parse-excel.js',
    'package.json',
    'package-lock.json',
    'requirements.txt',
    'config.json',
    'Dockerfile',
    'docker-compose.yml',
    'supervisord.conf',
    'nginx.conf',
    '.dockerignore',
    'README.md',
    'DEPLOY.md',
    'INSTALL.md',
    'install.sh',
    'install.bat'
];

coreFiles.forEach(file => {
    if (fs.existsSync(file)) {
        fs.copyFileSync(file, path.join(PACKAGE_NAME, file));
        console.log(`      已复制: ${file}`);
    }
});

// 复制目录
const directories = ['public', '.streamlit', 'fonts'];
directories.forEach(dir => {
    if (fs.existsSync(dir)) {
        copyDir(dir, path.join(PACKAGE_NAME, dir));
        console.log(`      已复制: ${dir}/`);
    }
});

// 创建空目录
console.log('[4/5] 创建空目录结构...');
const emptyDirs = ['data', 'data/backups', 'output', 'uploads', 'templates'];
emptyDirs.forEach(dir => {
    const fullPath = path.join(PACKAGE_NAME, dir);
    fs.mkdirSync(fullPath, { recursive: true });
    fs.writeFileSync(path.join(fullPath, '.gitkeep'), '');
});
console.log('      目录结构创建完成');

// 创建压缩包
console.log('[5/5] 创建压缩包...');

try {
    // 使用PowerShell的Compress-Archive
    const zipCommand = `powershell -Command "Compress-Archive -Path '${PACKAGE_NAME}/*' -DestinationPath '${PACKAGE_NAME}.zip' -Force"`;
    execSync(zipCommand, { stdio: 'ignore' });
    console.log('      压缩包创建完成');
} catch (error) {
    console.log('      PowerShell压缩失败，尝试使用tar...');
    try {
        const tarCommand = `tar -czf ${PACKAGE_NAME}.tar.gz -C ${PACKAGE_NAME} .`;
        execSync(tarCommand, { stdio: 'ignore' });
        console.log('      tar.gz压缩包创建完成');
    } catch (e) {
        console.log('      压缩失败，请手动打包');
    }
}

// 清理临时目录
fs.rmSync(PACKAGE_NAME, { recursive: true, force: true });

console.log('');
console.log('==========================================');
console.log('  打包完成！');
console.log('==========================================');
console.log('');

if (fs.existsSync(`${PACKAGE_NAME}.zip`)) {
    const stats = fs.statSync(`${PACKAGE_NAME}.zip`);
    const sizeKB = (stats.size / 1024).toFixed(2);
    const sizeMB = (stats.size / (1024 * 1024)).toFixed(2);
    
    console.log('生成的文件：');
    console.log(`  - ${PACKAGE_NAME}.zip`);
    console.log('');
    console.log('文件大小：');
    console.log(`  - ${sizeKB} KB (${sizeMB} MB)`);
} else if (fs.existsSync(`${PACKAGE_NAME}.tar.gz`)) {
    const stats = fs.statSync(`${PACKAGE_NAME}.tar.gz`);
    const sizeKB = (stats.size / 1024).toFixed(2);
    const sizeMB = (stats.size / (1024 * 1024)).toFixed(2);
    
    console.log('生成的文件：');
    console.log(`  - ${PACKAGE_NAME}.tar.gz`);
    console.log('');
    console.log('文件大小：');
    console.log(`  - ${sizeKB} KB (${sizeMB} MB)`);
}

console.log('');
console.log('部署说明：');
console.log('  - Windows: 解压后运行 install.bat');
console.log('  - Linux:   解压后运行 sudo ./install.sh');
console.log('  - Docker:  解压后运行 docker compose up -d --build');
console.log('');
console.log('详细文档请查看：');
console.log('  - README.md  (使用说明)');
console.log('  - DEPLOY.md  (部署指南)');
console.log('');
