# Everything 全盘文件搜索 Skill

本 skill 让智能体（Kimi / Claude / Codex 等）能够通过 [Voidtools Everything](https://www.voidtools.com/) 的命令行接口 `es.exe`，在 Windows 电脑上执行全硬盘范围的高速文件搜索。

## 目录

- [Everything 全盘文件搜索 Skill](#everything-全盘文件搜索-skill)
  - [目录](#目录)
  - [适用场景](#适用场景)
  - [前提条件](#前提条件)
  - [文件结构](#文件结构)
  - [使用方法](#使用方法)
    - [通过封装脚本搜索（推荐）](#通过封装脚本搜索推荐)
    - [直接使用 es.exe](#直接使用-esexe)
  - [常用参数](#常用参数)
  - [Everything 查询语法速查](#everything-查询语法速查)
  - [示例](#示例)
  - [故障排查](#故障排查)
  - [更新 es.exe](#更新-esexe)

## 适用场景

当需要**在整个电脑硬盘范围内**查找文件、文档或文件夹时，即可触发本 skill：

- 查找所有 PDF、Word、Excel、PPT 等文档
- 按文件名、大小、修改时间排序筛选
- 搜索特定路径或父目录下的文件
- 导出搜索结果为文本、CSV、EFU 等格式
- 通过正则表达式进行高级搜索

## 前提条件

1. **Everything 客户端已安装并正在运行**（窗口程序 `Everything.exe`）。
2. **命令行工具 `es.exe` 可访问**。

> 本 skill 已在 `bin/es.exe` 中内置 x64 版本的 `es.exe`。如果使用 `scripts/search.py`，会自动优先调用内置版本，无需额外配置。

## 文件结构

```
everything-file-search/
├── SKILL.md              # 智能体读取的 skill 说明
├── README.md             # 人类可读的说明文档（本文件）
├── bin/
│   └── es.exe            # Everything 命令行工具
└── scripts/
    └── search.py         # 搜索封装脚本
```

## 使用方法

### 通过封装脚本搜索（推荐）

```bash
python .agents/skills/everything-file-search/scripts/search.py "*.pdf" -n 10
```

脚本会自动查找 `es.exe`，构造命令行参数，并格式化输出结果。

### 直接使用 es.exe

也可以直接调用内置的 `es.exe`：

```bash
".agents/skills/everything-file-search/bin/es.exe" "*.pdf" -n 10
```

或系统中其他位置的 `es.exe`：

```bash
"C:/Program Files/Everything/es.exe" "*.pdf" -n 10
```

## 常用参数

| 参数 | 说明 |
|------|------|
| `-q, --query` | 搜索查询（Everything 查询语法） |
| `-n, --max-results` | 最大结果数，默认 100 |
| `-r, --regex` | 使用正则表达式 |
| `-i, --case` | 区分大小写 |
| `-w, --whole-word` | 全字匹配 |
| `-p, --match-path` | 匹配完整路径 |
| `--path <path>` | 限制搜索路径 |
| `--parent <path>` | 搜索指定父目录下的子文件 |
| `--sort <field>` | 排序字段：`name`、`path`、`size`、`date-modified` 等 |
| `--sort-ascending` / `--sort-descending` | 排序方向 |
| `--show-size` | 显示文件大小 |
| `--show-date-modified` | 显示修改日期 |
| `--format text/json/list` | 输出格式 |
| `-o, --output <file>` | 导出到文件 |

## Everything 查询语法速查

| 查询示例 | 含义 |
|----------|------|
| `*.pdf` | 所有 PDF 文件 |
| `report docx` | 文件名同时包含 report 和 docx |
| `*.xlsx size:>1mb` | 大于 1MB 的 Excel 文件 |
| `*.pdf dm:lastweek` | 上周修改的 PDF 文件 |
| `folder:projects` | 名为 projects 的文件夹 |
| `C:\Users\*.txt` | 限制在 `C:\Users` 下的 txt 文件 |
| `regex:^2024.*\.docx$` | 配合 `-r` 使用正则 |

更多语法参阅：https://www.voidtools.com/zh-cn/support/everything/searching/

## 示例

```bash
# 搜索包含 "预算" 的 Excel 文件
python .agents/skills/everything-file-search/scripts/search.py "预算 *.xlsx" -n 50

# 最近 3 天修改的 PDF，按修改时间降序
python .agents/skills/everything-file-search/scripts/search.py "*.pdf dm:last3days" --sort date-modified --sort-descending

# 最大的 10 个视频文件
python .agents/skills/everything-file-search/scripts/search.py "ext:mp4;mkv;avi;mov" --sort size --sort-descending -n 10

# 搜索名为 project 的文件夹
python .agents/skills/everything-file-search/scripts/search.py "folder:project"

# 导出所有 MP3 的完整路径列表
python .agents/skills/everything-file-search/scripts/search.py "*.mp3" --format list -o "D:/mp3_list.txt"
```

## 故障排查

| 错误 | 原因与处理 |
|------|-----------|
| `es.exe not found` | `es.exe` 未找到。检查 `bin/es.exe` 是否存在，或从官网下载后放置到 PATH 中。 |
| `Everything IPC window not found` | Everything 客户端未运行。请启动 `Everything.exe`。 |
| 搜索结果为空 | 查询语法可能有误，或 Everything 尚未完成索引。 |
| 结果过多 | 增加筛选条件，或使用 `-n` 限制结果数量。 |

## 更新 es.exe

内置的 `es.exe` 版本为 `ES-1.1.0.30 x64`。如需更新：

1. 访问 https://www.voidtools.com/zh-cn/downloads/
2. 下载 **Command-line Interface (ES)** 的 x64 版本
3. 解压并覆盖 `.agents/skills/everything-file-search/bin/es.exe`
