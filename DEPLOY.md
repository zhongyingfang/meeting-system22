# 部署指南

## 最新版本：v1.2.0

### 版本更新内容

**v1.2.0** (2026-04-29)
- 🔧 **修复Linux服务器字体加载问题**
  - 增强 `load_chinese_font()` 函数容错能力
  - 添加本地 `fonts/` 目录支持
  - 改进TTC字体文件处理
  - 多级字体回退机制
  - 确保程序在字体加载失败时仍能正常运行

**v1.1.0**
- 新增会议举办方角色
- 新增座位编排功能
- 新增待安排参会者姓名搜索
- 支持中文输入和模糊搜索

---

## 快速部署

### 方式一：Docker部署（推荐）

```bash
# 1. 解压部署包
tar xzf meeting-system-v1.2.0.tar.gz -C /opt/
cd /opt/meeting-system

# 2. 构建并启动
docker compose up -d --build

# 3. 查看日志
docker compose logs -f
```

访问地址：
- 座位查询与管理：http://服务器IP
- 座位牌生成：http://服务器IP:8506

---

### 方式二：Linux原生部署

#### 1. 安装系统依赖

```bash
# Ubuntu/Debian
sudo apt update
sudo apt install -y nodejs npm python3 python3-pip \
    fonts-noto-cjk fonts-wqy-zenhei fonts-wqy-microhei poppler-utils

# CentOS/RHEL
sudo yum install -y nodejs npm python3 python3-pip
```

#### 2. 解压并安装依赖

```bash
# 解压部署包
tar xzf meeting-system-v1.2.0.tar.gz -C /opt/
cd /opt/meeting-system

# 安装Node.js依赖
npm install --production

# 安装Python依赖
pip3 install -r requirements.txt
```

#### 3. 启动服务

```bash
# 使用supervisord管理（推荐）
# 先安装supervisor
pip3 install supervisor

# 启动所有服务
supervisord -c supervisord.conf
```

或者手动启动：

```bash
# 启动座位查询系统（后台运行）
node server.js &

# 启动座位牌生成系统（后台运行）
streamlit run app.py --server.port 8505 --server.address 0.0.0.0 --server.headless true &
```

#### 4. 配置开机自启

创建systemd服务文件 `/etc/systemd/system/meeting-system.service`：

```ini
[Unit]
Description=Meeting Seat Management System
After=network.target

[Service]
Type=simple
User=www-data
WorkingDirectory=/opt/meeting-system
ExecStart=/usr/bin/supervisord -n -c /opt/meeting-system/supervisord.conf
Restart=always

[Install]
WantedBy=multi-user.target
```

启动服务：

```bash
sudo systemctl daemon-reload
sudo systemctl enable meeting-system
sudo systemctl start meeting-system
```

---

### 方式三：Windows部署

#### 1. 安装环境
- Node.js (14.x+)：https://nodejs.org/
- Python (3.8+)：https://www.python.org/downloads/

#### 2. 解压部署包
将 `meeting-system-v1.2.0.zip` 解压到 `D:\meeting-system`

#### 3. 安装依赖

```cmd
cd D:\meeting-system
npm install
pip install -r requirements.txt
```

#### 4. 启动服务

```cmd
# 窗口1：启动座位查询系统
node server.js

# 窗口2：启动座位牌生成系统
streamlit run app.py --server.port 8505 --server.address 0.0.0.0
```

---

## 字体配置（重要！）

### Linux服务器字体问题修复

**问题现象：**
```
OSError: [Errno 5] Input/output error
```

**解决方案：**

#### 方案1：使用本地fonts目录（推荐）

部署包已包含常用中文字体，直接使用即可：

```
fonts/
├── msyh.ttc       # 微软雅黑
├── simhei.ttf     # 黑体
├── simsun.ttc     # 宋体
├── FZXBSJW.ttf    # 方正小标宋
└── 榜书字体.ttf   # 榜书字体
```

#### 方案2：安装系统字体

```bash
# Ubuntu/Debian
sudo apt install -y fonts-noto-cjk fonts-wqy-zenhei fonts-wqy-microhei

# CentOS/RHEL
sudo yum install -y google-noto-cjk-fonts wqy-zenhei-fonts wqy-microhei-fonts
```

#### 方案3：手动上传字体

将需要的字体文件（.ttf/.otf/.ttc）上传到项目的 `fonts/` 目录即可。

---

## 默认配置

| 项目 | 配置值 |
|------|--------|
| 管理员账号 | admin / admin888 |
| 会议举办方账号 | organizer / organizer888 |
| 座位查询端口 | 3000 (Docker: 80) |
| 座位牌生成端口 | 8505 (Docker: 8506) |

---

## 安全建议

1. **修改默认密码**：首次登录后立即修改
2. **配置防火墙**：仅开放必要端口
3. **启用HTTPS**：生产环境配置SSL证书
4. **定期备份**：备份 `data/` 目录
5. **更新依赖**：定期更新 npm 和 pip 包

---

## 故障排查

### 字体加载失败

**症状：** 座位牌生成时报错或中文显示方框

**解决：**
1. 确认 `fonts/` 目录存在且有字体文件
2. 检查文件权限：`chmod 644 fonts/*`
3. 重启服务

### 服务启动失败

**检查：**
```bash
# 查看端口占用
ss -tlnp | grep -E '3000|8505'

# 查看日志
docker compose logs -f
# 或
tail -f access.log
```

---

## 更新日志

### v1.2.0
- 修复Linux服务器字体加载问题
- 增强容错能力
- 添加本地字体目录支持

### v1.1.0
- 新增会议举办方角色
- 新增座位编排功能
- 新增姓名搜索功能

### v1.0.0
- 初始版本
