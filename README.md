<div align="center">

# 临床医疗与效率智能体技能集

*专为临床医生、住培师资与日常高频电脑工作者量身打造的 AI Agent 技能库与本地命令行效率工具。*

[![Python Version](https://img.shields.io/badge/Python-3.8%2B-blue?style=flat-square&logo=python&logoColor=white)](https://www.python.org/)
[![Platform Support](https://img.shields.io/badge/Platform-Windows%20%7C%20Linux%20%7C%20macOS-lightgrey?style=flat-square)]()
[![Agent Compatible](https://img.shields.io/badge/Agent-Antigravity%20%7C%20Claude%20Code%20%7C%20Cursor-orange?style=flat-square)]()
[![Pipeline](https://img.shields.io/badge/Pipeline-零幻觉本地库检索-success?style=flat-square)]()

⭐ 如果本项目对您的临床工作、住培带教或办公效率有所帮助，请在 GitHub 上点亮 Star 支持！

[项目概述](#项目概述) • [核心特性](#核心特性) • [技能矩阵](#技能矩阵) • [技能详解与用法](#技能详解与用法) • [快速上手](#快速上手) • [目录结构](#目录结构) • [重要注意事项](#重要注意事项)

</div>

---

## 项目概述

本项目是一套围绕临床实际业务场景与工作效率深度研发的智能体（AI Agent）专业技能库，由临床一线医师兼住培指导教师独立编写与维护。

仓库涵盖了临床工作与教学中的高频痛点：
- **湖北 2.0 版 ICD 编码与手术智能查询**：毫秒级精准定位疾病诊断与手术操作编码，全量标注手术级别；
- **医保 DIP 入组推荐与分值优化**：基于官方核心/综合病组目录测算最高分值入组方案及预估支付费用，严守病情事实红线；
- **住培教学查房教案自动生成**：严格依据国家 2021 版住培指南，一键生成 8,000–12,000 字符合国家公文标准排版的 DOCX 教案；
- **住培教学病例讨论教案撰写**：结构化启发式提问与指南权威解析，支持 Markdown 与 DOCX 导出；
- **Windows 全盘极速文件检索**：整合 Everything 引擎，实现毫秒级全盘文件瞬时定位与筛选。

所有技能均原生兼容现代 AI 智能体规范（Google Antigravity、Claude Code、Cursor 等），各技能目录下附带纯 Python 标准库编写的独立 CLI 脚本，支持在无 AI 环境的终端下直接秒级运行。

---

## 核心特性

- **双阶段防幻觉机制**：医疗编码、手术分级与 DIP 分值均通过本地预编译数据库和规则引擎硬匹配返回，大模型仅负责语义理解与分流，严守数据真实性。
- **零额外依赖负担**：核心查询脚本完全基于 Python 3 标准库（`sqlite3`、`re`、`json`、`argparse`、`pathlib`）构建，开箱即用，无需复杂的 `pip` 包配置。
- **国家指南与公文规范排版**：住培教学教案严格遵循《住院医师规范化培训教学指南（2021 年版）》，导出的 Word 文件内置符合公文标准的版心边距、仿宋/黑体/楷体字号与 28 磅固定行距。
- **双模灵活驱动**：既支持在 AI 对话框中以自然语言交互触发，也可作为高频命令行工具在终端中秒级执行。

---

## 技能矩阵

| 技能目录 | 技能名称 | 核心应用场景 | 技术方案与引擎 | 调用方式 |
| :--- | :--- | :--- | :--- | :--- |
| [`skill/icd-search-skill`](#1-icd-search-skill-湖北-20-版-icd-编码查询) | **湖北 2.0 版 ICD 编码查询** | 疾病诊断、手术术式编码精准查询与手术分级核验 | 本地 SQLite（约 6.7 MB，近5万条），纯标准库 | AI 对话 / CLI |
| [`skill/dip-advisor`](#2-dip-advisor-医保-dip-入组推荐顾问) | **医保 DIP 入组推荐顾问** | DIP 最优入组方案推荐、主次诊断调优、医保支付费用测算 | 核心/综合病组目录索引，纯标准库 | AI 对话 / CLI |
| [`skill/教学查房教案自动生成技能`](#3-教学查房教案自动生成技能-住培教学查房教案自动生成) | **住培教学查房教案自动生成** | 自动生成 8,000–12,000 字符合公文排版规范的住培教学查房教案 | 规则引擎 + `python-docx` 公文排版 | AI 对话 / Python |
| [`skill/teaching-case-discussion`](#4-teaching-case-discussion-住培教学病例讨论教案撰写) | **住培教学病例讨论教案撰写** | 启发式教学病例讨论教案编写、自查清单与质量控制 | Markdown 模板 + `python-docx` | AI 对话 / Python |
| [`skill/everything-file-search`](#5-everything-file-search-windows-全盘极速文件检索) | **Windows 全盘极速文件检索** | 硬盘全盘文件秒级定位、按大小/时间/类型高级过滤 | Voidtools Everything `es.exe`，Python 封装 | AI 对话 / CLI |

---

## 技能详解与用法

### 1. `icd-search-skill`：湖北 2.0 版 ICD 编码查询

面向临床医师和病案质控人员的本地化精准检索工具。

- **全量内置官方数据**：高压缩 SQLite 数据库（`icd_hubei2.db`，约 6.7 MB），收录湖北 2.0 版 ICD-10 疾病诊断 35,725 条与 ICD-9-CM-3 手术操作 13,664 条。
- **混合输入智能拦截**：自动拦截同时混杂“疾病诊断”与“手术操作”的模糊输入（如“高血压和阑尾切除”），引导用户明确意图，防止编码交叉混淆。
- **手术级别自动关联**：每项手术均附带国家法定手术级别（`[1级]` 至 `[4级]`），便于术前权限审核与分级质控。
- **关联度排序机制**：支持字词精准匹配、首字加权与长短惩罚；模糊检索时默认返回关联度最高的 Top 5 候选。

#### 使用示例

**AI 对话模式：**
> “查一下手术编码：单侧乳房改良根治术”  
> ➜ `单侧乳腺改良根治术（85.4301） [3级]`

**终端 CLI 模式：**
```bash
# 查询手术（附带手术级别）
python skill/icd-search-skill/scripts/query.py -t surgery "腹腔镜下胆囊切除术"

# 查询疾病诊断
python skill/icd-search-skill/scripts/query.py -t disease "原发性高血压"

# 依据编码反查名称
python skill/icd-search-skill/scripts/query.py "85.4301"

# 输出为 JSON 结构化数据
python skill/icd-search-skill/scripts/query.py "前哨淋巴结" --json
```

---

### 2. `dip-advisor`：医保 DIP 入组推荐顾问

根据患者病案实际情况，智能匹配推荐分值最高的 DIP 入组方案，并计算医保支付预估金额。

- **全要素病组目录**：内置 ICD-10 诊断库（35,723 条）、核心病组目录（5,669 条）、综合病组目录（2,417 条）以及 DIP 相关操作库（1,511 条）。
- **多维度对比推荐**：
  - 支持主次诊断互换及同族上级类目微调；
  - 对比同诊断下不同操作（保守治疗、穿刺引流、腔镜手术等）的分值差异；
  - 严格按 DIP 分值降序输出前 3~5 个方案。
- **支付标准实时测算**：支持按地区系数动态测算结算标准（默认内嵌：职工 = 分值 × 5.69，居民 = 分值 × 4.36）。
- **合规底线硬约束**：严禁捏造病历未提及的诊断；严禁推荐与临床病情脱节的无效手术冲分。

#### 使用示例

**AI 对话模式：**
> “患者女性，诊断：左侧乳腺浸润性癌、高血压病2级，拟行左乳腺癌改良根治术。请推荐 DIP 入组方案并对比分值。”

**终端 CLI 模式：**
```bash
# 查询诊断关键词与编码
python skill/dip-advisor/scripts/search.py diag "胆囊结石"

# 查看该诊断编码可入的所有病组与分值
python skill/dip-advisor/scripts/search.py group "K80.200"

# 查看同类目下各 4 位编码核心病组的最高分值
python skill/dip-advisor/scripts/search.py family "K80"
```

---

### 3. `教学查房教案自动生成技能`：住培教学查房教案自动生成

依据《住院医师规范化培训教学查房指南（2021 年版）》，根据临床病案自动生成 8,000–12,000 字高质量教学查房教案，并直接排版输出符合国家公文格式规范的 Word 文档。

- **查房“三部曲”全流程闭环**：
  1. *示教室准备阶段（5–10 分钟）*：宣布查房主题、明确三维教学目标、进行病例特点归纳。
  2. *床旁采集阶段（15–20 分钟）*：住院医师脱稿汇报病史、带教老师示范规范问诊与体格检查、人文关怀与医患沟通。
  3. *示教室讨论阶段（40–60 分钟）*：启发式互动提问、临床诊断思维训练、结合最新指南讨论诊疗方案与文献拓展。
- **SMART 原则三维目标**：细分为知识目标、能力目标与素养目标。
- **三层次启发式问题设计**：涵盖事实性问题、分析性问题与评价性问题，均附带临床指南权威参考答案。
- **公文规范排版标准**：生成标准 A4 规格文档（版心 156mm × 225mm，上/下边距 36mm，左/右边距 27mm，每页 22 行每行 28 字，正文三号仿宋_GB2312，固定行距 28 磅，四级标题层次清晰规范）。

#### 使用示例

**AI 对话模式：**
> “根据这例急性坏疽性阑尾炎的病例资料，帮我生成一份符合住培指南和公文排版规范的教学查房教案 Word 文档。”

**Python 代码调用：**
```python
import sys
sys.path.insert(0, r'skill/教学查房教案自动生成技能')
from generate_jiaoan import generate_from_case

case_data = {
    '教学查房标题': '急性阑尾炎教学查房',
    '患者基本信息': '男性，28岁',
    '主诉': '转移性右下腹痛伴恶心发热1天',
    '现病史': '...',
    '体格检查': '右下腹麦氏点明显压痛、反跳痛...',
    '辅助检查': '血常规 WBC 14.5×10^9/L, B超示阑尾明显增粗...',
    '初步诊断': '急性化脓性阑尾炎',
    '诊疗经过': '拟急诊行腹腔镜阑尾切除术...'
}

# 一键生成符合公文格式规范的 DOCX
generate_from_case(case_data, '急性阑尾炎教学查房教案.docx', '急性阑尾炎')
```

---

### 4. `teaching-case-discussion`：住培教学病例讨论教案撰写

协助临床带教老师高效设计与撰写住培教学病例讨论教案。

- **规范化教学设计**：围绕典型病例展开，涵盖教学目标制定、讨论问题层层设疑、时间分配掌控（标准 60 分钟）以及课后评价反馈。
- **指南循证答案**：每个核心讨论问题均结合临床最新指南与专家共识，生成 100–300 字精要参考答案。
- **双格式支持**：提供用于课件演示与快速浏览的 Markdown 格式，以及用于正式归档的公文规范 DOCX 格式。
- **配套质控清单**：技能内含质量检查清单（`checklists/`）、撰写指南（`references/`）与标准化模板（`templates/`）。

#### 使用示例

**AI 对话模式：**
> “我需要一份关于‘下肢深静脉血栓形成’的住培教学病例讨论教案，时长 60 分钟，请提供三维教学目标与递进式讨论问题。”

**Python 代码调用：**
```python
from skill.teaching_case_discussion.scripts.generate_teaching_plan import generate_teaching_plan

plan_data = {
    '基本信息': {
        '培训基地': '十堰市人民医院普外科基地',
        '专业基地/科室': '甲乳血管外科一病区',
        '指导医师': '王巍（主治医师）',
        '教学时长': '60 分钟'
    },
    '教学主题': '下肢深静脉血栓形成的抗凝与手术指征评估',
    # ... 详细教案字典数据
}

generate_teaching_plan(plan_data, '深静脉血栓教学病例讨论教案.docx')
```

---

### 5. `everything-file-search`：Windows 全盘极速文件检索

基于 Voidtools Everything 搜索引擎，赋予 AI 智能体与命令行在 Windows 全盘百万文件中毫秒级检索定位的能力。

- **极速响应**：基于 NTFS USN 日志机制，实现全盘文件瞬时搜索。
- **自包含绿色组件**：技能 `bin/` 目录下已预置 64 位便携版 `es.exe`，免除手动配置环境变量的步骤。
- **丰富语法支持**：支持通配符、多关键词、文件大小筛选（`size:>500MB`）、修改时间筛选（`dm:last3days`）、正则表达式匹配以及结果导出。

> [!NOTE]
> 运行该技能前，请确保 Windows 系统后台已启动 Everything 客户端（`Everything.exe`）。

#### 使用示例

**终端 CLI 模式：**
```bash
# 搜索包含特定关键词的文档
python skill/everything-file-search/scripts/search.py "住培教案 *.docx" -n 20

# 检索最近 3 天修改过的所有 PDF 文件
python skill/everything-file-search/scripts/search.py "*.pdf dm:last3days" --sort date-modified --sort-descending

# 查找全盘排名前 10 的大视频文件
python skill/everything-file-search/scripts/search.py "ext:mp4;mkv;avi" --sort size --sort-descending -n 10
```

---

## 快速上手

### 场景一：作为 AI Agent 技能载入（Antigravity、Claude Code、Cursor）

1. **克隆本仓库到本地：**
   ```bash
   git clone https://github.com/wwsky7787/skill.git
   ```
2. **载入技能：**
   - **工作空间级挂载**：直接在智能体工具中打开本仓库目录，支持 `SKILL.md` 的 Agent 会自动扫描并识别各子目录下的技能；
   - **全局挂载**：将 `skill/` 目录下的各技能子文件夹软链接或复制到全局 Agent 技能目录中（如 Windows 下为 `C:\Users\<用户名>\.gemini\config\skills\` 或 `~/.agents/skills/`）。

### 场景二：作为独立本地 Python 脚本使用

核心查询与分析工具无需复杂环境配置：
- **Python 版本**：Python 3.8 或更高版本；
- **基础运行依赖**：Python 原生标准库（自带 `sqlite3`, `re`, `json`, `argparse`, `pathlib`）；
- **Word 导出依赖**（仅教学教案生成 DOCX 时需要）：
  ```bash
  pip install python-docx
  ```

---

## 目录结构

```text
skill/
├── README.md                            # 项目总览与使用说明文档
└── skill/
    ├── icd-search-skill/                # 湖北 2.0 版 ICD 编码查询技能
    │   ├── SKILL.md                     # Agent 技能定义
    │   ├── README.md                    # 技能独立详尽文档
    │   ├── icd_hubei2.db                # 湖北 2.0 版 SQLite 完整数据库（约 6.7 MB）
    │   └── scripts/
    │       ├── query.py                 # 零依赖本地检索引擎
    │       └── build_db.py              # 数据重构与编译脚本
    │
    ├── dip-advisor/                     # 医保 DIP 入组推荐与分值顾问
    │   ├── SKILL.md                     # Agent 技能定义
    │   ├── data/                        # 诊断、核心病组、综合病组与操作库 JSON
    │   └── scripts/
    │       └── search.py                # 零依赖 DIP 与 ICD 检索计算脚本
    │
    ├── 教学查房教案自动生成技能/        # 住培教学查房教案生成（8000-12000字公文排版 DOCX）
    │   ├── SKILL.md                     # Agent 技能定义
    │   ├── 使用说明.md                  # 详细使用手册与格式说明
    │   ├── generate_jiaoan.py           # 简化生成接口
    │   └── scripts/
    │       └── generate_teaching_rounds_plan.py # 公文格式排版与生成引擎
    │
    ├── teaching-case-discussion/        # 住培教学病例讨论教案撰写
    │   ├── SKILL.md                     # Agent 技能定义
    │   ├── checklists/                  # 教学质量自查清单
    │   ├── references/                  # 住培指南规范参考
    │   ├── templates/                   # 讨论教案标准模板
    │   └── scripts/
    │       └── generate_teaching_plan.py# 公文 DOCX 生成脚本
    │
    └── everything-file-search/          # Windows Everything 全盘检索
        ├── SKILL.md                     # Agent 技能定义
        ├── README.md                    # 独立说明文档
        ├── bin/
        │   └── es.exe                   # 预置 64 位 Everything 命令行工具
        └── scripts/
            └── search.py                # 带参数解析的 Python 调用封装
```

---

## 重要注意事项

> [!IMPORTANT]
> **患者隐私保护（数据脱敏）**
> 在通过 AI 智能体对话处理真实病例数据时，请务必提前脱敏并隐去患者姓名、身份证号、住院号、电话等直接标识患者个人身份的信息，严格遵守医疗数据伦理规范。

> [!CAUTION]
> **临床辅助决策定位与免责声明**
> 本仓库中所有技能生成的内容（包括 ICD 编码建议、DIP 入组方案、教学教案等）仅供**临床诊疗辅助参考、病案质控自查与教学研讨**使用，不作为正式的临床处方、医疗鉴定结论或医保结算承诺。实际临床诊疗与病案归档决策始终以具备执业资质的医务人员判断为准。
