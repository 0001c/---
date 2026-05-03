<p align="center">
  <img src="https://img.shields.io/badge/Python-3.12-blue?logo=python" alt="Python">
  <img src="https://img.shields.io/badge/Flask-3.1+-lightgrey?logo=flask" alt="Flask">
  <img src="https://img.shields.io/badge/Docker-✓-2496ED?logo=docker" alt="Docker">
  <img src="https://img.shields.io/badge/license-MIT-green" alt="License">
</p>

<h1 align="center">xx调剂定制班 — 智能调剂规划生成工具</h1>

<p align="center">
读取考研调剂意向调查统计表 → 调用大模型 AI 分析 → 自动生成每位考生专属的 <strong>1v1 个人评估表 Excel</strong>
</p>

---

## 📋 目录

- [项目介绍](#-项目介绍)
- [核心功能](#-核心功能)
- [快速开始（Web 版）](#-快速开始web-版)
- [快速开始（命令行版）](#-快速开始命令行版)
- [Docker 部署](#-docker-部署)
- [项目结构](#-项目结构)
- [技术栈](#-技术栈)
- [数据格式说明](#-数据格式说明)
- [常见问题](#-常见问题)

---

## 🎯 项目介绍

考研调剂窗口期短、信息量大。传统模式下，调剂顾问需要逐一翻查考生问卷、查找往年录取数据、评估匹配度，再撰写个性化规划，效率极低且主观性强。

本工具将这一流程自动化：

1. **读取** — 从问卷系统导出的 `统计表.xlsx` 中提取每位考生的 20 个字段（本科院校、初试分数、调剂意向等）
2. **分析** — 调用火山引擎大模型（豆包）分两轮进行深度分析：
   - 第一轮：6 维度调剂可能性分析（分数、本科、一志愿、语种、作品集、时间精力）
   - 第二轮：生成"冲/稳/保"三档 × "A区/B区"两地共 6 格院校推荐矩阵
   - 每次分析附带联网搜索，确保引用 2026 年最新调剂数据
3. **生成** — 利用 Excel 模板，自动填充每位考生的评估报告

---

## ✨ 核心功能

| 功能 | 说明 |
|------|------|
| 📂 **文件上传** | 上传统计表和模板（支持拖拽），模板可选（使用内置默认） |
| 👥 **考生预览** | 自动解析考生列表，支持姓名/学校/专业搜索过滤 |
| ✅ **灵活选择** | 全选 / 全不选 / 逐位勾选，可只处理部分考生 |
| 🤖 **AI 双轮分析** | 调剂可能性分析 + 院校推荐，附带联网搜索确保数据时效性 |
| 📊 **实时进度** | WebSocket 推送，每位考生经历"准备数据→AI分析→生成表格"三步可视化 |
| 📥 **结果下载** | 单文件下载或一键批量打包 ZIP |
| 🖥️ **双模式** | Web 端（推荐） + 命令行版（备选） |
| 🐳 **容器化** | Docker 镜像即拉即用，NAS / VPS 均可部署 |

---

## 🚀 快速开始（Web 版）

### 前置条件
- Python 3.12+
- 火山引擎 API 密钥（环境变量 `ARK_API_KEY`）

### 安装 & 运行

```bash
# 1. 克隆仓库
git clone <your-repo-url>
cd 数据表

# 2. 创建虚拟环境并安装依赖
python -m venv .venv
.venv\Scripts\activate      # Windows
# source .venv/bin/activate  # Linux/macOS

pip install flask flask-socketio pandas openpyxl openai

# 3. 设置 API 密钥
set ARK_API_KEY=your_api_key_here

# 4. 启动（双击 start_web.bat，或手动执行）
cd web
python app.py
```

### 使用流程

```
1️⃣ 浏览器打开 http://localhost:5000
2️⃣ 上传最新的 统计表.xlsx（必需）+ 模板.xlsx（可选）
3️⃣ 浏览考生列表，勾选需要处理的考生
4️⃣ 点击"开始 AI 分析"，实时观看处理进度
5️⃣ 处理完成后，下载单份或批量打包 ZIP
```

---

## 📟 快速开始（命令行版）

适合不需要 Web 界面的快速处理场景：

```bash
set ARK_API_KEY=your_api_key_here
cd 数据表
python edit.py

# 按提示输入考生序号（如 "1,3,5" 或 "all"）
```

生成的 Excel 保存在 `output/` 目录。

---

## 🐳 Docker 部署

### 构建 & 启动

```bash
# docker-compose（推荐）
ARK_API_KEY=your_key docker-compose up -d

# 或 docker run
docker run -d --name ximeng-tiaoji \
  -p 5000:5000 \
  -e ARK_API_KEY=your_key \
  -v ./docker-data/uploads:/app/uploads \
  -v ./docker-data/output:/app/output \
  --restart unless-stopped \
  ximeng-tiaoji-web:latest
```

### 在 NAS 上部署

```bash
# 1. 在本机导出镜像
docker save ximeng-tiaoji-web:latest -o ximeng-tiaoji-web.tar

# 2. 将 .tar 文件传输到 NAS
# 3. 在 NAS 上加载并启动
docker load -i ximeng-tiaoji-web.tar
docker run -d --name ximeng-tiaoji \
  -p 5000:5000 \
  -e ARK_API_KEY=your_key \
  -v /path/to/uploads:/app/uploads \
  -v /path/to/output:/app/output \
  ximeng-tiaoji-web:latest
```

> 镜像基于 `python:3.12-slim`，内容大小仅 **98.6MB**（tar 包约 94MB），自带健康检查。

---

## 📁 项目结构

```
数据表/
├── ai_tool.py                # AI 大模型调用（火山引擎豆包 API）
├── edit.py                   # 命令行版主程序
├── read.py                   # 数据读取模块（pandas 解析统计表）
├── 模板.xlsx                 # 输出 Excel 模板（A1/A3/A5/A7/A9 预留单元格）
│
├── web/                      # Web 应用
│   ├── app.py                # Flask 后端（上传/解析/AI调度/生成/下载）
│   ├── requirements.txt      # Python 依赖清单
│   ├── templates/
│   │   └── index.html        # 前端页面（Tailwind + Socket.IO）
│   ├── uploads/              # 上传文件目录（运行时）
│   └── output/               # 输出文件目录（运行时）
│
├── start.bat                 # 命令行版启动脚本
├── start_web.bat             # Web 版启动脚本
│
├── Dockerfile                # Docker 镜像构建
├── docker-compose.yml        # Docker Compose 编排
├── .dockerignore             # Docker 构建忽略
├── .gitignore                # Git 忽略规则
│
├── output/                   # 命令行版输出目录（运行时）
└── docker-data/              # Docker 数据持久化（运行时）
```

---

## 🛠️ 技术栈

| 层级 | 技术 |
|------|------|
| 后端框架 | Python **Flask** + **Flask-SocketIO** |
| 数据读写 | **pandas**（统计表解析）、**openpyxl**（Excel 生成） |
| AI 大模型 | 火山引擎方舟平台 — **doubao-seed-2-0-lite**（分析）+ **doubao-seed-2-0-mini**（推荐） |
| 前端 | **Tailwind CSS** + 原生 JavaScript + **Socket.IO** |
| 容器化 | **Docker** + **Docker Compose** |
| 部署 | 任意支持 Docker 的环境（NAS / VPS / 云服务器） |

### 依赖清单

```txt
flask>=3.0
flask-socketio>=5.3
pandas>=2.0
openpyxl>=3.1
openai>=1.0
```

---

## 📊 数据格式说明

### 统计表（输入）

从问卷系统导出的 `.xlsx` 文件，格式要求：

| 列 | 字段名 | 内容 |
|----|--------|------|
| 第1-6列 | （元数据） | 问卷系统自动附加的提交时间、来源等 |
| **第7列起** | data1 ~ data20 | 有效数据（共 20 个字段） |

核心字段映射：

| 字段 | 含义 | 字段 | 含义 |
|------|------|------|------|
| data1 | 姓名 | data8 | 初试总分 |
| data2 | 目前学历 | data9 | 外语科目和分数 |
| data3 | 本科院校及专业 | data10 | 政治分数 |
| data4 | 一志愿报考院校 | data11~14 | 专业课 / 专项计划 / 手机号 |
| data5 | 一志愿专业代码和全称 | data16~19 | 调剂意向（学习方式 / 专硕学硕 / 区域 / 等级） |
| data6 | 学硕/专硕 | data20 | 艺术大类调剂意向（含 `〖...〗` 正则提取） |
| data7 | 学习方式 | | |

> ⚠️ **统计表含姓名、手机号等个人隐私，已在 .gitignore 中排除，请勿提交到 GitHub。**

### 个人评估表（输出）

以 `模板.xlsx` 为基准，在 5 个固定单元格写入内容：

| 单元格 | 内容 |
|--------|------|
| A1 | `{姓名}调剂整体规划` |
| A3 | 基础信息（15 项：学历、院校、成绩等） |
| A5 | 调剂意向（5 项：方式、区域、等级等） |
| A7 | AI 调剂可能性分析（6 维度） |
| A9 | AI 推荐院校（冲/稳/保 × A区/B区） |

---

## ❓ 常见问题

**Q: 需要什么 API 密钥？**
A: 需要火山引擎方舟平台的 API Key（环境变量 `ARK_API_KEY`），获取方式见 [火山引擎官方文档](https://www.volcengine.com/docs/82379/1399008)。

**Q: 统计表的格式变了怎么处理？**
A: 第 7 列起的数据映射在 `read.py`（命令行版）和 `web/app.py` 的 `parse_stats_file()` 函数中定义，问卷格式变更时修改 `row[6:]` 的起始位置即可。

**Q: AI 分析要多久？**
A: 每位考生约 1-3 分钟（取决于 API 响应速度），Web 端实时显示进度。

**Q: 可以在 NAS 上跑吗？**
A: 可以。导出 Docker 镜像（`docker save`）传到 NAS，`docker load` + `docker run` 即可运行。

**Q: 同时处理多位考生会冲突吗？**
A: 当前为单用户设计，后台线程逐位串行处理，如需并发可调整 `app.py` 中的分析线程逻辑。

---

## 📄 License

MIT
