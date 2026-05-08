# 动态血糖记录生成器 — 程序开发手册

## 项目概述

本项目是一个基于 Flask 的 Web 应用，运行于 Docker 容器中，提供以下功能：

- 通过浏览器上传 **Excel 血糖数据文件**
- 从 Excel 中筛选各时间点（如 06:00、08:30、10:30 等）的血糖值
- 按照内置标准 Word 模板（`血糖数据-01.docx`）的格式，生成新的 Word 文档并支持下载

---

## 目录结构

```
bloodsugar/
├── app.py                  # Flask 主应用
├── requirements.txt        # Python 依赖
├── Dockerfile              # Docker 镜像构建文件
├── docker-compose.yml      # Docker Compose 配置
├── templates/
│   └── index.html          # 前端上传界面
├── uploads/                # 临时存放上传文件（自动创建）
├── outputs/                # 备用输出目录（自动创建）
├── 血糖数据-01.docx         # 示例 Word 模板
└── 血糖数据-ZP1OO08785505.xlsx  # 示例 Excel 数据
```

---

## 技术栈

| 层次 | 技术 |
|------|------|
| Web 框架 | Flask 3.1 |
| Excel 解析 | openpyxl 3.1 |
| Word 生成 | python-docx 1.1 |
| XML 处理 | lxml 6.1 |
| 生产服务器 | Gunicorn 23 |
| 容器化 | Docker / Docker Compose |

---

## 核心逻辑说明

### 1. Excel 解析（`parse_excel`）

- 读取 `.xlsx` / `.xls` 文件的活跃工作表
- 自动识别"血糖时间"列（含"时间"、"日期"等关键字）和"血糖值"列（含"血糖"、"mmol"等关键字）
- 将每行解析为 `(datetime, float)` 元组列表

### 2. Word 模板解析（`parse_word_template`）

- 读取 Word 文件第一个表格的表头行
- 跳过第一列（"日期"），提取其余列的时间点字符串（如 `06:00`）

### 3. 血糖值匹配（`match_readings`）

- 将记录按日期分组
- 对每个日期和每个时间点，在 `±window_minutes`（默认 4 分钟）窗口内查找时间距离最近的血糖读数
- 返回按日期和时间点组织的字典

### 4. Word 文档生成（`generate_word`）

- 以内置标准模板（`血糖数据-01.docx`）为基础
- 将标题段落更新为 `YYYY-MM-DD至YYYY-MM-DD动态血糖记录`
- 删除模板中的原有数据行，保留表头
- 按日期添加新数据行，复制模板中的单元格/行格式
- 将结果写入内存缓冲区并作为文件下载返回

---

## 快速启动

### 方式一：Docker Compose（推荐）

```bash
docker-compose up --build -d
```

访问 [http://localhost:5002](http://localhost:5002)

### 方式二：直接运行（开发环境）

```bash
pip install -r requirements.txt
python app.py
```

访问 [http://localhost:5002](http://localhost:5002)

---

## 使用流程

1. 打开浏览器，访问 `http://<服务器IP>:5002`
2. 上传 **血糖数据 Excel 文件**（格式：`.xlsx`，包含时间列和血糖值列）
3. 可选调整**时间匹配窗口**（分钟），默认 4 分钟
4. 点击"生成 Word 文档并下载"，浏览器自动下载生成的文档

---

## Excel 文件格式要求

| 血糖时间 | 血糖值 mmol/L | 说明 |
|----------|--------------|------|
| 2026-04-24 06:03 | 5.4 | |
| 2026-04-24 08:33 | 6.4 | |
| ... | ... | ... |

- 时间格式：`YYYY-MM-DD HH:MM` 或 `YYYY-MM-DD HH:MM:SS`
- 血糖值：数字（mmol/L）

## Word 模板格式要求

表格结构：

| 日期 | 06:00 | 08:30 | 10:30 | 13:00 | 16:30 | 19:00 | 21:00 |
|------|-------|-------|-------|-------|-------|-------|-------|
| 2026-01-04 | ... | ... | ... | ... | ... | ... | ... |

- 第一列必须为"日期"
- 其余列为 `HH:MM` 格式的时间点
- 时间点数量和名称可自定义

---

## 环境变量

| 变量 | 说明 | 默认值 |
|------|------|--------|
| `SECRET_KEY` | Flask session 密钥（生产环境请修改） | `bloodsugar-secret-2026` |

---

## API 端点

| 路径 | 方法 | 说明 |
|------|------|------|
| `/` | GET | 上传/下载界面 |
| `/generate` | POST | 处理上传文件，返回生成的 Word 文档 |
| `/health` | GET | 健康检查，返回 `{"status": "ok"}` |

---

## 注意事项

- 上传文件在处理完成后会立即删除，不会持久保存
- 生产部署时请修改 `docker-compose.yml` 中的 `SECRET_KEY`
- 如需支持超大 Excel 文件，可在 `docker-compose.yml` 中增加 Gunicorn 超时时间
