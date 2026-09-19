---
name: everything-file-search
description: 当用户需要在整个电脑硬盘范围内搜索文件、文档、文件夹或任意本地文件系统项目时，使用本 skill 通过 Everything 命令行接口（es.exe）执行搜索并返回结果。触发场景包括但不限于："在整个电脑里找..."、"搜索硬盘上的..."、"帮我定位文件..."、"查找包含...的文件"、"列出所有...文件"，以及任何涉及全盘/全硬盘范围文件检索的请求。
---

# Everything 全盘文件搜索

本 skill 指导智能体通过 Voidtools Everything 的命令行接口 `es.exe`，在 Windows 电脑上执行全硬盘范围的高速文件搜索。

## 前提条件

1. **Everything 客户端必须已安装并正在运行**（窗口程序 `Everything.exe`）。
2. **命令行工具 `es.exe` 必须可访问**。

本 skill 已在 `bin/es.exe` 中附带 `es.exe`（x64 版本），`scripts/search.py` 会自动优先使用它。

如果未使用封装脚本，也可以将 `es.exe` 放在以下位置之一：
   - 本 skill 的 `bin/` 目录
   - `C:\Program Files\Everything\es.exe`
   - `C:\Program Files (x86)\Everything\es.exe`
   - 任意已加入系统 `PATH` 的目录

> 如需单独下载，可前往 https://www.voidtools.com/zh-cn/downloads/ 下载 **Command-line Interface (ES)**。

## 何时使用本 skill

只要用户请求涉及**在整个电脑/硬盘/本地文件系统中查找文件或文件夹**，即应使用本 skill，例如：

- "帮我找一下电脑上所有叫 report 的 Word 文档"
- "搜索硬盘里最近一周修改过的 PDF"
- "查找包含 '预算' 两个字的 Excel 文件"
- "列出 C 盘所有大于 100MB 的视频文件"
- "在整个电脑范围内搜索名为 project 的文件夹"

## 使用方法

### 1. 定位 es.exe

优先使用本 skill 附带的 `scripts/search.py` 脚本，它会自动查找 `es.exe` 并执行搜索。

如果脚本不可用，手动调用 `Bash` 执行 `es.exe`，并尝试以下常见路径：

```bash
"/c/Program Files/Everything/es.exe" <query> [options]
"/c/Program Files (x86)/Everything/es.exe" <query> [options]
```

### 2. 执行搜索

推荐通过 `scripts/search.py` 调用，避免手动拼接命令行：

```bash
python .agents/skills/everything-file-search/scripts/search.py "*.pdf"
```

常用参数：

| 参数 | 说明 |
|------|------|
| `-q, --query` | 搜索查询（Everything 查询语法） |
| `-n, --max-results` | 最大结果数，默认 100 |
| `-r, --regex` | 使用正则表达式 |
| `-i, --case` | 区分大小写 |
| `-w, --whole-word` | 全字匹配 |
| `-p, --match-path` | 匹配完整路径 |
| `--sort` | 排序字段，如 `name`, `path`, `size`, `date-modified` 等 |
| `--sort-ascending` / `--sort-descending` | 排序方向 |
| `-path` | 限制在某个路径下搜索 |
| `-parent` | 搜索指定父目录下的子文件 |
| `-export-txt` | 导出到文本文件 |
| `--format` | 输出格式：`text`（默认）、`json`、`list` |

### 3. 返回结果

搜索完成后，将结果以清晰、结构化的方式返回给用户：

- 如果结果较少（< 20 条），直接列出完整路径。
- 如果结果较多，给出摘要统计，并说明保存的导出文件路径。
- 如果搜索失败，说明错误原因（如 Everything 未运行、es.exe 未找到）。

## Everything 查询语法速查

| 查询示例 | 含义 |
|----------|------|
| `*.pdf` | 所有 PDF 文件 |
| `report docx` | 文件名同时包含 report 和 docx |
| `*.xlsx size:>1mb` | 大于 1MB 的 Excel 文件 |
| `*.pdf dm:lastweek` | 上周修改的 PDF 文件 |
| `folder:projects` | 名为 projects 的文件夹 |
| `C:\Users\*.txt` | 限制在 C:\Users 下的 txt 文件 |
| `regex:^2024.*\.docx$` | 配合 `-r` 使用正则 |

更多语法参阅：https://www.voidtools.com/zh-cn/support/everything/searching/

## 示例

### 搜索所有包含 "预算" 的 Excel 文件

```bash
python .agents/skills/everything-file-search/scripts/search.py "预算 *.xlsx" -n 50
```

### 查找最近 3 天修改的所有 PDF

```bash
python .agents/skills/everything-file-search/scripts/search.py "*.pdf dm:last3days" --sort date-modified --sort-descending
```

### 查找电脑上最大的 10 个视频文件

```bash
python .agents/skills/everything-file-search/scripts/search.py "ext:mp4;mkv;avi;mov" --sort size --sort-descending -n 10
```

### 搜索名为 "project" 的文件夹

```bash
python .agents/skills/everything-file-search/scripts/search.py "folder:project"
```

### 导出所有 MP3 列表到文件

```bash
python .agents/skills/everything-file-search/scripts/search.py "*.mp3" --format list -o "D:/mp3_list.txt"
```

## 错误处理

| 错误 | 原因与处理 |
|------|-----------|
| `es.exe not found` | 未安装或不在 PATH 中，引导用户下载并放置 es.exe |
| `Everything IPC window not found` | Everything 客户端未运行，提示用户启动 Everything |
| 搜索结果为空 | 确认查询语法，或检查 Everything 是否已建立索引 |
| 结果过多 | 增加筛选条件或限制 `-n` 结果数 |

## 安全与隐私提示

- 全盘搜索可能返回敏感路径，展示结果前注意上下文。
- 不要对搜索结果执行未经用户确认的删除、修改操作。
- 避免在查询中暴露密码、密钥等敏感信息。
