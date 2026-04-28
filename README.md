# 会议座位管理系统

一个功能完整的会议座位管理系统，支持座位编排、参会者管理、区域规划和座位牌生成等功能。

## 系统功能概述

### 核心模块

| 模块 | 说明 | 技术栈 |
|------|------|--------|
| 座位查询系统 | Excel上传、布局解析、座位搜索、预览图生成 | Node.js + Express |
| 管理后台 | 会场管理、参会者管理、座位编排、权限控制 | Node.js + Express |
| 座位牌生成 | 批量生成可打印的PDF座位牌 | Python + Streamlit |

### 主要功能特性

#### 1. 登录与权限控制
- **双角色登录**：支持管理员和会议举办方两种角色
- **头像选择**：可视化的角色选择界面
- **密码加密**：使用bcrypt加密存储密码
- **权限控制**：根据角色显示/隐藏不同功能

#### 2. 会场管理
- **Excel导入**：支持从Excel文件导入会场布局
- **多会场支持**：一个Excel文件可包含多个Sheet（每个Sheet一个会场）
- **布局类型**：自动检测或手动指定三种布局类型（普通/剧院/U型）
- **智能布局检测**：自动识别沙发排、过道、空座位

#### 3. 参会者管理
- **Excel导入**：支持批量导入参会者信息
- **手动添加**：支持单个添加参会者
- **编辑删除**：支持修改和删除参会者信息
- **分类筛选**：支持按单位、职级筛选参会者
- **多选操作**：支持批量选择参会者

#### 4. 座位编排
- **拖拽分配**：拖拽参会者到指定座位
- **区域规划**：鼠标框选座位区域，保存区域名称
- **批量分配**：将选中的参会者批量分配到指定区域
- **随机分配**：支持随机分配参会者到区域
- **座位清理**：支持清空座位上的参会者
- **实时统计**：显示已安排和待安排人数

#### 5. 数据导出
- **SVG座次图**：高清矢量图，适合喷绘打印
- **Excel导出**：导出座位安排表
- **PDF座位牌**：批量生成座位牌，支持自定义样式

#### 6. 数据安全
- **自动备份**：每次数据写入自动备份，保留最近10份
- **手动备份/恢复**：管理后台支持备份和恢复操作
- **审计日志**：所有数据操作均有日志记录

## 快速开始

### 环境要求

| 项目 | 最低要求 | 推荐配置 |
|------|----------|----------|
| 操作系统 | Windows 10+ / Linux (Ubuntu 20.04+) | Windows 11 / Ubuntu 22.04 |
| Node.js | 14.x | 18.x |
| Python | 3.8 | 3.10 |
| CPU | 2 核 | 4 核 |
| 内存 | 2 GB | 4 GB |
| 磁盘 | 10 GB | 20 GB SSD |

### Windows快速部署

#### 1. 安装Node.js
下载并安装Node.js: https://nodejs.org/

#### 2. 安装Python
下载并安装Python: https://www.python.org/downloads/
安装时勾选"Add Python to PATH"

#### 3. 解压部署包
将 `meeting-system-v1.1.0.zip` 解压到 `D:\meeting-system`

#### 4. 安装依赖
```cmd
cd D:\meeting-system

# 安装Node.js依赖
npm install

# 安装Python依赖
pip install -r requirements.txt
```

#### 5. 启动服务
```cmd
# 启动座位查询与管理后台（端口3000）
node server.js

# 新开一个命令行窗口，启动座位牌生成（端口8505）
streamlit run app.py --server.port 8505 --server.address 0.0.0.0
```

#### 6. 访问系统
- 用户查询页面：http://localhost:3000/
- 管理后台：http://localhost:3000/admin.html
- 座位牌生成：http://localhost:8505/

### Linux快速部署

#### 1. 安装依赖
```bash
# Ubuntu/Debian
sudo apt update
sudo apt install -y nodejs npm python3 python3-pip \
    fonts-noto-cjk fonts-wqy-zenhei fonts-wqy-microhei poppler-utils

# CentOS/RHEL
sudo yum install -y nodejs npm python3 python3-pip
```

#### 2. 解压部署包
```bash
tar xzf meeting-system-v1.1.0.tar.gz -C /opt/
cd /opt/meeting-system
```

#### 3. 安装依赖
```bash
# Node.js依赖
npm install --production

# Python依赖
pip3 install -r requirements.txt
```

#### 4. 启动服务
```bash
# 启动座位查询与管理后台
node server.js &

# 启动座位牌生成
streamlit run app.py --server.port 8505 --server.address 0.0.0.0 --server.headless true &
```

### Docker部署（推荐）

#### 1. 安装Docker
```bash
# Ubuntu
curl -fsSL https://get.docker.com | sudo sh
sudo systemctl enable docker
sudo systemctl start docker
```

#### 2. 部署项目
```bash
# 解压部署包
tar xzf meeting-system-v1.1.0.tar.gz -C /opt/
cd /opt/meeting-system

# 构建并启动
docker compose up -d --build

# 查看日志确认启动
docker compose logs -f
```

#### 3. 访问服务
| 服务 | 端口 | 访问地址 |
|------|------|----------|
| 座位查询与管理后台 | 80 | `http://服务器IP` |
| 座位牌生成 | 8506 | `http://服务器IP:8506` |

## 配置说明

### 默认账号

| 角色 | 用户名 | 密码 | 说明 |
|------|--------|------|------|
| 管理员 | admin | admin888 | 拥有所有权限 |
| 会议举办方 | organizer | organizer888 | 仅限座位编排功能 |

### 首次配置

1. 打开浏览器访问管理后台：`http://服务器IP/admin.html`
2. 选择角色（管理员/会议举办方）并输入对应密码
3. 首次登录后请立即修改密码！

### 系统配置文件

`config.json` - 系统配置文件

```json
{
  "title": "会议座位查询系统",
  "adminPassword": "admin888",
  "organizerPassword": "organizer888"
}
```

## Excel文件格式规范

### 会场布局Excel

#### 基本规则
- 每个Sheet = 一个会场，Sheet 名称 = 会场名称
- 系统支持三种布局类型，自动检测或手动指定

#### 普通会场（standard）
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

#### 剧院会场（theater）
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

#### U型会场（u-shape）
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

### 参会者Excel

参会者列表文件应包含以下列：
- 姓名（必填）
- 单位（可选）
- 职级（可选）
- 备注（可选）

## 管理后台操作指南

### 登录系统
1. 访问 `/admin.html`
2. 选择角色（管理员/会议举办方）
3. 输入密码登录

### 会场管理
1. 点击"上传会场Excel"
2. 选择Excel文件
3. 系统自动检测每个Sheet的布局类型
4. 确认后上传

### 参会者管理
1. 点击"上传参会者Excel"导入参会者
2. 或点击"添加参会者"手动添加
3. 使用筛选功能按单位/职级筛选
4. 支持编辑和删除参会者

### 座位编排
1. 选择会场
2. **方法一（拖拽）**：从待安排列表拖拽参会者到座位
3. **方法二（区域分配）**：
   - 鼠标框选座位区域
   - 输入区域名称并保存
   - 选择参会者（可多选）
   - 选择目标区域，点击"分配到区域"
4. 点击座位上的"×"可以清除该座位

### 数据导出
- **导出SVG座次图**：高清矢量图，适合喷绘打印
- **导出Excel**：导出座位安排表
- **生成座位牌**：跳转到座位牌生成系统

### 系统配置
1. 点击右上角"系统配置"
2. 修改管理员和会议举办方密码
3. 点击"保存"

## 座位牌生成系统

### 功能特点
- 多种座位牌样式（普通座位牌、三角立式台卡）
- 支持自定义字体、颜色、大小
- 支持添加背景图片
- 支持立体、阴影、勾边等文字效果
- 实时预览功能

### 数据来源
1. 上传Excel文件（姓名列表或座位布局）
2. 从座位查询系统导入（推荐）

### 生成步骤
1. 选择数据来源
2. 选择要生成座位牌的参会者
3. 调整样式设置（字体、颜色、大小等）
4. 查看实时预览
5. 点击"生成PDF文件"
6. 下载生成的PDF

## 目录结构

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
├── README.md              # 本说明文档
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
│   ├── msyh.ttc           # 微软雅黑
│   ├── simhei.ttf         # 黑体
│   └── simsun.ttc         # 宋体
│
└── .streamlit/            # Streamlit 配置
    └── config.toml
```

## 数据备份与恢复

### 自动备份
系统在每次数据写入时自动备份，保留最近 10 份备份在 `data/backups/`。

### 手动备份
1. 登录管理后台
2. 进入"数据备份"面板
3. 点击"下载备份"

### 恢复备份
1. 登录管理后台
2. 进入"数据备份"面板
3. 点击"上传备份"
4. 选择备份文件并恢复

## 常用维护命令

### Windows
```cmd
# 查看端口占用
netstat -ano | findstr ":3000"
netstat -ano | findstr ":8505"

# 结束进程
taskkill /F /PID <进程ID>
```

### Linux
```bash
# 查看服务状态（Docker）
docker compose ps

# 查看实时日志（Docker）
docker compose logs -f

# 重启服务（Docker）
docker compose restart

# 停止服务（Docker）
docker compose down

# 更新并重建（Docker）
docker compose up -d --build

# 查看端口占用
ss -tlnp | grep -E '3000|8505'
```

## 安全建议

1. **修改默认密码**：首次部署后立即修改默认密码
2. **配置防火墙**：仅开放必要端口（3000/80、8505/8506）
3. **HTTPS**：生产环境建议配置SSL证书
4. **定期更新**：保持系统依赖和Docker镜像更新
5. **日志审计**：定期检查审计日志（`audit.log`）
6. **数据备份**：定期备份重要数据

## 故障排查

| 现象 | 排查步骤 |
|------|----------|
| 服务启动失败 | 检查端口是否被占用，查看错误日志 |
| 端口被占用 | Windows: `netstat -ano` \| Linux: `ss -tlnp` |
| Excel 上传失败 | 检查文件格式（.xlsx），确认无密码保护 |
| 布局解析错误 | 检查Excel是否符合格式规范 |
| 预览图错位 | 确保选择了正确的布局模式（标准/剧院/U型） |
| 座位牌生成失败 | 检查Python依赖是否完整安装 |
| 字体显示异常 | 将字体文件放入 `fonts/` 目录，重启服务 |
| 登录失败 | 确认选择了正确的角色，检查密码是否正确 |
| 会议举办方看不到功能 | 会议举办方仅有座位编排权限 |

## 技术支持

如遇到问题，请检查：
1. 日志文件（`access.log`、`audit.log`）
2. 端口占用情况
3. 依赖是否完整安装

## 版本历史

### v1.1.0
- 新增会议举办方角色
- 新增座位编排功能（拖拽、区域规划）
- 新增参会者分类筛选
- 新增参会者编辑删除
- 修复统计信息实时更新
- 修复会场显示问题

### v1.0.0
- 初始版本
- 支持Excel导入会场布局
- 支持座位查询
- 支持SVG座次图导出
- 支持PDF座位牌生成

## 许可证

仅供内部使用。
