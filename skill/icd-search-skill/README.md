# 湖北2.0版 ICD-10/ICD-9-CM-3 编码智能查询技能 (Portable ICD Search)

本技能专为临床医生和病案质控设计，支持根据自然语言临床描述，毫秒级精准检索**湖北2.0版ICD-10疾病诊断编码**与**ICD-9-CM-3手术操作编码**。

---

## 🌟 核心特性

1. **两阶段防幻觉架构（LLM + SQLite）**：
   - 大模型负责理解临床自然语言（同义词、缩写、别名），转化为规范医学实体；
   - 底层 Python 检索脚本直接查询本地官方标准库，确保编码、名称、手术级别 **100% 真实准确，零臆造**。
2. **极速轻量与零外部依赖**：
   - 原始 2.38MB Excel 经清洗、重构，压缩为仅 **6.7 MB** 的单文件 SQLite 数据库；
   - 核心检索引擎 `scripts/query.py` **完全基于 Python 3 标准库**（仅用 `sqlite3`, `re`, `argparse`, `json`, `pathlib`），**无需安装任何 pip 依赖包**。
3. **跨平台全自包含（Windows & Ubuntu）**：
   - 所有路径均为相对动态探测，无绝对路径硬编码；
   - 随拷随用，在 Windows PowerShell / CMD、Ubuntu / Debian Linux、macOS 终端均可原生无感运行。
4. **混合输入拦截机制**：
   - 严格识别并拒绝同时混杂疾病与手术的输入（如“高血压和阑尾切除术”），避免编码混乱，明确临床指代。
5. **智能关联度排序**：
   - 支持全字精准命中、前后缀包含、字词重合度与长短惩罚；无法唯一确定时自动按相关度降序输出 **Top 5 候选**。
6. **规范输出合同**：
   - **手术格式**：`手术名称（手术编码） [X级]`（例如：`淋巴结扩大性区域性切除术（40.3x00x001） [3级]`）
   - **疾病格式**：`疾病名称（疾病编码）`（例如：`高血压（I10.x00x002）`）

---

## 📁 目录与文件清单

```text
skill-ICD编码查询-2026.09.15/
├── SKILL.md                          # Antigravity 技能主定义（定义意图路由与输出合同）
├── README.md                         # 详细使用指南与运维文档（本文档）
├── icd-search-skill.zip              # 便携免安装发布包 (~2.9 MB，解压即用)
├── icd_hubei2.db                     # 优化编译后的 SQLite 数据库 (~6.7 MB，含近5万条记录与全文索引)
├── 湖北2.0版ICD编码库...副本.xlsx      # 原始官方 Excel 底表 (备份留底)
└── scripts/
    ├── query.py                      # 核心检索引擎（生产运行脚本，无依赖）
    └── build_db.py                   # 数据库编译维护脚本（后续 Excel 更新时使用）
```

---

## 🚀 快速上手使用指南

### 方式一：在 AI 对话中自然语言调用（Antigravity Skill）

本技能已部署在当前环境。今后你在 AI 聊天框中直接输入任意临床自然语言即可：

- **查手术**：
  > “查一下手术编码：单侧乳房改良根治术”  
  > ➜ 输出：`单侧乳腺改良根治术（85.4301） [3级]`
- **模糊查术式**：
  > “淋巴结扩大切除”  
  > ➜ 输出 Top 5 相关手术选项供选择。
- **查疾病诊断**：
  > “ICD编码：2型糖尿病并发周围神经病变”  
  > ➜ 输出：`2型糖尿病性周围神经病变（E11.401+G63.5*）`
- **混合输入自动拦截**：
  > “查一下阑尾炎和切除术”  
  > ➜ 拦截提示：“拒绝查询：检测到您的输入同时包含疾病与手术（或意图不明确），请明确本次是查询【疾病诊断】还是【手术操作】，或分开发送查询。”

---

### 方式二：终端命令行独立调用 (CLI - Windows / Ubuntu)

无需启动 AI，在系统终端中可秒级直查。

#### 1. 查询手术（自动提取并附带手术级别）：
```bash
# 精准术式
python scripts/query.py "淋巴结扩大性区域性切除术"
# 输出：淋巴结扩大性区域性切除术（40.3x00x001） [3级]

# 模糊/习惯用语（输出 Top 5 候选）
python scripts/query.py "乳腺癌改良根治"
python scripts/query.py "大隐静脉高位结扎剥脱"
```

#### 2. 查询疾病诊断：
```bash
python scripts/query.py "原发性高血压"
# 输出：高血压（I10.x00x002）

python scripts/query.py "2型糖尿病"
# 输出：2型糖尿病（E11.900）
```

#### 3. 根据编码反查名称：
```bash
python scripts/query.py "40.3x00x001"
# 输出：淋巴结扩大性区域性切除术（40.3x00x001） [3级]

python scripts/query.py "I10.x00x002"
# 输出：高血压（I10.x00x002）
```

#### 4. 强制指定类型查询（参数 `-t surgery` 或 `-t disease`）：
```bash
# 强制只查手术库
python scripts/query.py -t surgery "阑尾切除术"

# 强制只查疾病库
python scripts/query.py -t disease "高血压"
```

#### 5. 机器结构化 JSON 输出（便于其他程序集成）：
```bash
python scripts/query.py "前哨淋巴结" --json
```

---

## 🐧 Ubuntu / Linux 离线部署步骤

1. 将 `icd-search-skill.zip` 复制到 Ubuntu 机器的目标文件夹。
2. 解压并运行：
   ```bash
   unzip icd-search-skill.zip -d icd-search
   cd icd-search
   python3 scripts/query.py "腹腔镜下胆囊切除术"
   ```
3. **环境要求**：系统自带的 `python3`（Python 3.8+）即可，无需 `pip install` 任何库。

---

## 🛠️ 数据更新与维护（Excel 重新生成 SQLite）

如果湖北省卫健委或医保局后续下发了新版 Excel 编码表：
1. 将新的 Excel 文件放于根目录下（保留或替换原有 `.xlsx`）。
2. 安装 openpyxl 并运行构建脚本：
   ```bash
   pip install openpyxl
   python scripts/build_db.py
   ```
3. 脚本会自动重新提取疾病与手术工作表，构建高效的 SQLite 索引并完成压缩整理。
