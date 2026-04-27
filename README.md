# 会议座位管理系统 — 安装部署指南

## 一、系统概述

本系统用于各类会议的会场座位管理，整合三大核心功能：

| 模块 | 说明 | 技术栈 |
|------|------|--------|
| 座位查询系统 | Excel上传、布局解析、座位搜索、预览图生成 | Node.js + Express |
| SVG座次图导出 | 按布局类型渲染会场座位图，支持打印喷绘 | Node.js + 纯SVG |
| 座位牌生成 | 批量生成可打印的PDF座位牌 | Python + Streamlit |

### 核心特性

- **多Sheet多类型会场**：一个Excel文件含多个Sheet，系统自动检测每个Sheet的布局类型（剧院式 / 普通 / U型）并正确解析
- **座位数据预览**：上传后即时生成会场布局预览图和完整SVG座次图
- **智能布局检测**：自动识别沙发排、过道、空座位，按座位数分组，仅在每组首排显示座位号
- **数据对比**：支持新旧Excel对比，快速识别人员变动

---

## 二、运行环境要求

| 项目 | 最低要求 | 推荐配置 |
|------|----------|----------|
| 操作系统 | Linux (Ubuntu 20.04+) | Ubuntu 22.04 |
| CPU | 2 核 | 4 核 |
| 内存 | 2 GB | 4 GB |
| 磁盘 | 10 GB | 20 GB SSD |
| Docker | 20.10+ | 24.x |
| Docker Compose | 2.0+ | 2.x |

---

## 三、Docker 部署（推荐）

### 3.1 安装 Docker

```bash
# Ubuntu
curl -fsSL https://get.docker.com | sudo sh
sudo systemctl enable docker
sudo systemctl start docker
```

### 3.2 部署项目

```bash
# 1. 解压部署包
tar xzf meeting-system.tar.gz -C /opt/
cd /opt/meeting-system

# 2. 构建并启动
docker compose up -d --build

# 3. 查看日志确认启动
docker compose logs -f
```

服务端口：

| 服务 | 端口 | 访问地址 |
|------|------|----------|
| 座位查询与管理后台 | 80 | `http://服务器IP` |
| 座位牌生成 | 8506 | `http://服务器IP:8506` |

### 3.3 首次配置

1. 打开浏览器访问管理后台：`http://服务器IP/admin.html`
2. 默认密码：`admin888`（请立即修改）
3. 上传 Excel 文件 → 系统自动检测布局类型 → 生成预览
4. 如需修改密码，进入管理后台右上角"系统配置"

---

## 四、裸机部署（无 Docker）

### 4.1 安装依赖

```bash
# 系统级依赖
sudo apt update
sudo apt install -y nodejs npm python3 python3-pip \
    fonts-noto-cjk fonts-wqy-zenhei fonts-wqy-microhei poppler-utils nginx

# Node.js 依赖
cd /opt/meeting-system
npm install --production

# Python 依赖
pip3 install -r requirements.txt
```

### 4.2 配置 Nginx

```bash
sudo cp nginx.conf /etc/nginx/nginx.conf
sudo nginx -t && sudo systemctl restart nginx
```

### 4.3 启动服务

```bash
# 启动 Node.js 座位查询系统
node server.js &

# 启动 Streamlit 座位牌生成（后台运行）
streamlit run app.py --server.port 8505 --server.address 0.0.0.0 --server.headless true &
```

### 4.4 进程管理（systemd）

创建 `/etc/systemd/system/meeting-system.service`：

```ini
[Unit]
Description=会议座位管理系统
After=network.target

[Service]
Type=simple
WorkingDirectory=/opt/meeting-system
ExecStart=/usr/bin/node /opt/meeting-system/server.js
Restart=on-failure
RestartSec=5
Environment=NODE_ENV=production PORT=3000

[Install]
WantedBy=multi-user.target
```

```bash
sudo systemctl daemon-reload
sudo systemctl enable --now meeting-system
```

---

## 五、目录结构

```
meeting-system/
├── server.js              # 座位查询系统主程序（Node.js）
├── parse-excel.js         # Excel 解析引擎（布局检测 + 人名提取）
├── app.py                 # 座位牌生成系统主程序（Streamlit）
├── package.json           # Node.js 依赖描述
├── requirements.txt       # Python 依赖描述
├── config.json            # 系统配置文件（标题、密码）
├── nginx.conf             # Nginx 反向代理配置
├── Dockerfile             # Docker 镜像构建文件
├── docker-compose.yml     # Docker Compose 编排文件
├── supervisord.conf       # 容器内进程管理器配置
├── .dockerignore          # Docker 构建忽略规则
├── INSTALL.md             # 本部署指南
│
├── public/                # 前端静态页面
│   ├── index.html         # 用户查询首页
│   ├── admin.html         # 管理后台
│   └── login.html         # 登录页
│
├── data/                  # 数据存储（挂载卷）
│   ├── data.json          # 座位 / 参会者数据
│   ├── config.json        # 运行期配置
│   └── backups/           # 自动备份
│
├── output/                # 生成的 PDF / 预览文件
├── uploads/               # 上传的 Excel 文件
│
├── fonts/                 # 自定义字体（可选）
│   └── README.txt
│
└── .streamlit/            # Streamlit 配置
    └── config.toml
```

---

## 六、Excel 文件格式规范

### 6.1 基本规则

- 每个 Sheet = 一个会场，Sheet 名称 = 会场名称
- 系统支持三种布局类型，自动检测或手动指定

### 6.2 普通会场（standard）

适合大多数标准排座布局。

```
|      | A    | B  | C   | D     | E   | F    | G  | H   | I     | J   |
|------|------|----|-----|-------|-----|------|----|-----|-------|-----|
| 第1排|      | 1  |     | 2     |     | 3    |    | 4   |       |     |
|      | 张三 |    | 李四 |       |     | 王五 |    |     | 赵六  |     |
| 第2排|      | 1  |     | 2     |     | 3    |    | 4   |       |     |
|      | 甲   |    | 乙   |       |     | 丙   |    |     | 丁    |     |
```

- 座位号行和排标签可以在同行，也可以分开
- 列间距 > 1 的位置视为过道

### 6.3 剧院会场（theater）

适合有楼层结构的大型剧院。

```
|      | A  | B | C | D | E | F | G  |
|------|----|---|---|---|---|---|----|
| 一楼 |    |   |   |   |   |   |    |
| 一楼第1排 |  | 1 | 2 | 3 | 4 | 5 |  |
|      | 甲 | 乙 | 丙 | 丁 | 戊 | 己 |  |
```

- 必须有楼层标签（一楼/二楼 或 1F/2F）
- 不同楼层之间自动插入过道

### 6.4 U型会场（u-shape）

适合圆桌/U型会议布局。

```
|     | A      | B    | C      | D    | E      | F    | G      | H    |
|-----|--------|------|--------|------|--------|------|--------|------|
|     | 第二列 | 座位号 | 第一列 | 座位号 | 第三列 | 座位号 | 第四列 | 座位号 |
|     | 张三   | 1    | 李四   | 1    | 王五   | 1    | 赵六   | 1    |
|     | 钱七   | 2    | 孙八   | 2    | 周九   | 2    | 吴十   | 2    |
|     |        |      |        |      | 4     | 5    | 6     | 7    | ← 底部行座位号
|     |        |      |        |      | 郑一   | 冯二 | 陈三  | 褚四 | ← 底部行人名
```

- 必须有 `第X列` 标签（≥2列）
- 两个最内侧列之间的数字行 = 底部行座位

---

## 七、管理后台操作指南

### 7.1 上传座位表

1. 登录管理后台 `/admin.html`
2. 拖拽或点击上传 Excel 文件
3. 系统自动检测每个 Sheet 的布局类型，多 Sheet 时可手动调整
4. 选择"解析模式"或保持"自动检测"
5. 点击"上传表格"

### 7.2 生成预览图

点击"生成预览图"按钮，查看所有会场的缩略布局。

### 7.3 导出座次图

- **导出座位安排表 Excel**：带人名的可视化 Excel
- **导出座位安排表 SVG**：带人名的高清矢量图（适合喷绘打印）

### 7.4 对比 Excel

上传新版 Excel → 点击"开始对比" → 查看人员变动（新增/移除/换座）→ 选择应用。

---

## 八、数据备份与恢复

### 自动备份

系统在每次数据写入时自动备份，保留最近 10 份备份在 `data/backups/`。

### 手动备份（Docker）

```bash
# 备份数据卷
docker run --rm -v meeting-system_meeting-data:/data \
  -v $(pwd):/backup alpine \
  tar czf /backup/meeting-data-$(date +%Y%m%d).tar.gz /data
```

### 恢复备份（Docker）

```bash
docker run --rm -v meeting-system_meeting-data:/data \
  -v $(pwd):/backup alpine \
  tar xzf /backup/meeting-data-YYYYMMDD.tar.gz -C /
```

### 管理后台恢复

也可以通过管理后台 → "数据备份" 面板直接下载/恢复备份。

---

## 九、常用维护命令

```bash
# 查看服务状态
docker compose ps

# 查看实时日志
docker compose logs -f

# 重启服务
docker compose restart

# 停止服务
docker compose down

# 更新并重建
docker compose up -d --build

# 进入容器
docker compose exec meeting-system bash

# 只启动座位查询系统（开发调试）
cd /opt/meeting-system && node server.js
```

---

## 十、安全建议

1. **修改默认密码**：首次部署后立即通过管理后台修改 `admin888`
2. **配置防火墙**：仅开放 80 和 8506 端口
3. **HTTPS**：生产环境建议配置 SSL 证书（nginx.conf 中取消 443 注释）
4. **定期更新**：保持系统依赖和 Docker 镜像更新
5. **日志审计**：所有数据操作均有审计日志（`audit.log`）

---

## 十一、故障排查

| 现象 | 排查步骤 |
|------|----------|
| 容器启动失败 | `docker compose logs` 查看错误日志 |
| 端口被占用 | `ss -tlnp \| grep -E '80\|8506'` |
| Excel 上传失败 | 检查文件格式（.xlsx），确认无密码保护 |
| 布局解析错误 | 检查 Excel 是否符合格式规范（见第六章） |
| 预览图错位 | 确保选择了正确的布局模式（标准/剧院/U型） |
| 字体缺失 | 将字体文件放入 `fonts/` 目录，重启容器 |

---

## 十二、许可证

仅供内部使用。
