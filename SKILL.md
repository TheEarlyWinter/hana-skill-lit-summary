---
name: lit-summary
description: "文献总结生成与 Word 排版技能（/lit-summary）。当用户发送文献 PDF、DOI、标题、摘要或正文全文，要求写文献总结、文献汇报、组会文献、每周文献总结、读文献并总结、提炼某篇文献的创新点/可借鉴点时必须使用。固定四大模块骨架（文献名片、主要内容概括、创新点两条三段式、可借鉴到本课题两到三条），影响因子与分区自动联网查证并注明口径，沿用实验总结（/summary）的两阶段交付：先在对话框输出 Markdown 草案，用户确认后再调用内置脚本导出排版规范的 Word 文档（宋体+Times New Roman、正文小四、标准三线表、图题），并自动维护文献台账去重。不负责替用户检索文献本身，只处理用户提供的文献。"
default-enabled: false
---

# 文献总结生成技能 (/lit-summary)

面向化学、化工、材料方向的每周文献总结。与 `/summary`（实验工作总结）共用同一套排版契约与审阅流程，只更换内容骨架。

**边界**：不替用户检索、挑选文献。文献由用户提供（PDF / DOI / 标题 / 摘要 / 正文）。技能负责读懂、概括、提炼创新点、落到本课题、导出 Word。唯一需要联网的环节是影响因子与分区的查证，见阶段 0 第 4 步。

**不署名**：文档标题与正文均不出现署名行、姓名、学号、"报告人"字样。

---

## 交付流程

```
[用户发送文献 PDF / DOI / 标题]
            ↓
【阶段 0：文献解析】
  1. PDF → 抽取正文文本（判断能否读全；缺页/乱码要如实说）
  2. PDF → 抽取关键图（机理图、性能/结构图），默认最多 3 张
  3. 元信息核对：题目、作者、期刊、年卷页、DOI，以文献首页为准
  4. 联网查证影响因子与分区（必做，见下）
  5. 台账查重：这篇是否写过
            ↓
【阶段 1：四大模块草案】—— 对话框输出 Markdown
            ↓
【阶段 2：用户审阅确认】—— 等明确确认（"同意""确认""生成word"）
            ↓
【阶段 3：导出 Word + 写入台账 + 交付】
```

**准则**：用户未明确确认前，不提前生成 Word 文件。

---

## 阶段 0：文献解析、查证与查重

### 1. 获取本地路径

- PDF 由用户上传时：用 `current_status` 看 session_files，再用 `materialize` 换成绝对路径。
- 用户只给 DOI 或标题时：不自作主张去找替代文献，直接说明"把 PDF 或正文贴给我"。

### 2. 抽取正文

用 PyMuPDF 读文本，确认能否读全正文（而不仅是摘要页）：

```powershell
& "C:\Users\WIT_User\AppData\Local\Programs\Python\Python312\python.exe" -c "import fitz,sys; d=fitz.open(sys.argv[1]); [print(f'<<<PAGE {i+1}>>>\n'+p.get_text()) for i,p in enumerate(d)]" "<pdf_path>"
```

> 本机 `python` 可能命中 WindowsApps 占位程序（退出码 9009），统一用上面的完整路径调用。

读不到正文（扫描版、纯图 PDF）时，明确说明情况并请用户贴文字，不凭摘要硬写。

### 3. 抽取关键图

```powershell
& "C:\Users\WIT_User\AppData\Local\Programs\Python\Python312\python.exe" "C:\Users\WIT_User\.hanako\skills\lit-summary\scripts\extract_figures.py" `
  --pdf "<pdf_path>" `
  --output-dir "C:\Users\WIT_User\Desktop\OH-WorkSpace\文献总结\.lit_cache\<短标题>"
```

- 默认抽嵌入位图；图片是矢量时加 `--pages 2,3` 只把这几页渲染成 300 DPI，再选图。
- 选图原则：**机理/设计示意图 1 张 + 关键性能或结构证据图 1 张**，最多 3 张。
- 图题格式：`图1 原文 Fig. 2 反应路径示意（引自原文）`，注明引自原文。

### 4. 影响因子与分区（必做，不许空着）

文献页面上没有，就联网查。用 `web_search` 检索，优先看期刊官网、Clarivate JCR 页面或权威汇总：

- 检索式示例：`<期刊全名> impact factor 2025`、`<期刊名> 中科院分区 2025`
- **分开写，不混成一句**：`影响因子 14.4（2024 JCR，Clarivate）`；`中科院分区 化学 1 区 Top（2025 年升级版）`
- 一定注明采用哪一年的口径，影响因子与分区逐年变动。
- 两个数据来源不一致时，取权威度更高的一侧，并写明取自何处。
- 万一检索不到（如新刊、会议论文集），才写"待核实"，并说明查过哪里。**任何情况下不凭记忆填数字。**

### 5. 台账查重

```powershell
& "C:\Users\WIT_User\AppData\Local\Programs\Python\Python312\python.exe" "C:\Users\WIT_User\.hanako\skills\lit-summary\scripts\ledger.py" check --doi "<doi>"
```

命中已有记录时先问用户"这篇之前写过（某日期），这次按什么角度重写：更详细 / 换侧重点 / 补新内容"，不默认重写一份。

---

## 阶段 1：四大模块草案

详细写法、正反例、学术化改写对照见 `references/structure_template.md`；课题背景与常用表征见 `references/project_profile.md`，可借鉴点必须咬住它写。

| # | 模块 | 一句话要求 |
|---|------|-----------|
| 一 | 文献名片 | 三线表：题目（原文英文）、作者、期刊卷页年、DOI、文献类型、影响因子与分区（含口径）；不附参考文献条目 |
| 二 | 主要内容概括 | 150～250 字，说清体系、方法、关键结果、结论，不夹评价 |
| 三 | 创新点（固定 2 条） | 每条三段式：相对于谁 · 新在哪儿 · 靠什么证据 |
| 四 | 可借鉴到本课题（2～3 条） | 每条四要素：用在哪个样品/实验 · 解决什么问题 · 具体怎么做 · 照搬还是改造；重点在"怎么做" |

**四个模块是全部**。不写证据链表、不写局限与可质疑处、不写与本课题已有工作的对照、不写本周可动手、末尾不附参考文献条目。需要判断的地方融进创新点的"证据"一行里说。

---

## 阶段 2：审阅与确认

草案输出后，末尾提示：

> 以上是这篇文献的总结草案。你核对一下创新点和可借鉴点有没有偏差，确认后回复**「同意」**或**「生成word」**，我导出 Word 文档。

用户提出修改时只改对应模块，并在回复里点明改了哪几处，不重发整篇。

---

## 阶段 3：导出 Word、写台账、交付

### 1. 组装 JSON

脚本接受通用结构（`title` / `subtitle` / `sections[]`，每个 section 含 `title`、`level`、`content[]`），内容块支持：

```json
{"type": "paragraph", "text": "...", "bold": false}
{"type": "list_item", "text": "..."}
{"type": "table", "headers": ["列1", "列2"], "rows": [["值1", "值2"]], "title": "表1 文献名片"}
{"type": "image", "image_path": "C:\\...\\fig.png", "caption": "图1 原文 Fig. 2 反应路径示意（引自原文）"}
```

- **层级只用 section 的 `level` 表达**（1/2/3），引擎没有 heading 内容块。要写"创新点1"这类三级标题，就单开一个 `level: 3` 的 section，条文放进它的 `content`。
- `list_item` 是缩进段落，不自动加项目符号，序号自己写进文本。
- 完整样例见 `references/example_data.json`，照抄骨架换内容即可。

JSON 存到 `C:\Users\WIT_User\Desktop\OH-WorkSpace\文献总结\.build\<日期>_<短标题>.json`，生成前先校验：

```powershell
& "C:\Users\WIT_User\AppData\Local\Programs\Python\Python312\python.exe" -c "import json,sys;json.load(open(sys.argv[1],encoding='utf-8'));print('json ok')" "<json_path>"
```

### 2. 生成 docx

```powershell
& "C:\Users\WIT_User\AppData\Local\Programs\Python\Python312\python.exe" "C:\Users\WIT_User\.hanako\skills\lit-summary\scripts\generate_docx.py" `
  --input-json "<json_path>" `
  --output "C:\Users\WIT_User\Desktop\OH-WorkSpace\文献总结\文献总结_YYYYMMDD_<短标题>.docx"
```

单篇：`文献总结_YYYYMMDD_<第一作者或关键词>.docx`。
一周多篇合成一份：`文献总结_YYYYMMDD.docx`，每篇一级标题 `文献一：<中文短标题>`，四个模块降为二级标题。

### 3. 排版契约（与 /summary 完全一致）

- 中西文分离：中文 `宋体`，英文/数字 `Times New Roman`
- 大标题小二（18 pt）加粗居中；副标题五号灰
- 一级标题小三（15 pt）加粗；二级四号（14 pt）加粗；三级小四（12 pt）加粗
- 正文小四（12 pt），两端对齐，1.25 倍行距
- 三线表：顶底线 1.5 pt，栏目线 0.75 pt，无竖线，表头加粗居中，表内五号
- 图：最大宽 13.5 cm 居中，图题五号居中

### 4. 写台账

```powershell
& "C:\Users\WIT_User\AppData\Local\Programs\Python\Python312\python.exe" "C:\Users\WIT_User\.hanako\skills\lit-summary\scripts\ledger.py" add `
  --doi "<doi>" --title "<原文题目>" --journal "<期刊>" --year <年> `
  --innovation "<一句话创新点>" --borrowable "<一句话可借鉴点>" --docx "<docx 绝对路径>"
```

台账固定落在 `C:\Users\WIT_User\Desktop\OH-WorkSpace\文献总结\文献台账.csv`，按 DOI 去重。

### 5. 交付

`stage_files` 交付 `.docx`。回复里只给一句结论 + 文件，不复述全文。

---

## 硬约束（写不出来也不许编）

1. **创新点固定两条**，每条三段式。`作者采用 Ni 基催化剂研究了…` 是工作不是创新，必须有"相对于谁、哪里不同、被什么证据支持"。
2. **创新点的"证据"一行要带判断**：说清是直接证据还是仅相关、对照全不全。这是全篇唯一保留判断力的地方，写软了整份总结就空。
3. **可借鉴点许实不许虚**：`可借鉴其表征方法` 属于没说。必须指到具体样品、具体环节、具体动作（投料、条件、观察指标、时间点）。
4. **主要内容概括只陈述不评价**：不写"该文创新性强""结果令人满意"，也不用堆数据的办法凑字数。
5. **影响因子与分区必须查证**，注明口径年份与来源，查不到才写"待核实"并说明查过哪里。
6. **不编造**：数字、页码、DOI、作者原话、期刊等级一律不造；原文没写的写"原文未明确"。
7. **不署名**，不写"报告人""总结人"。

---

## 文件结构

```
lit-summary/
├── SKILL.md
├── references/
│   ├── structure_template.md   四大模块写法、正反例、学术化改写对照
│   ├── project_profile.md      本课题背景与常用表征（可借鉴点据此撰写）
│   └── example_data.json       可直接套用的 JSON 骨架
└── scripts/
    ├── generate_docx.py        通用 Word 排版引擎（与 /summary 同一份）
    ├── extract_figures.py      文献 PDF 抽图 / 指定页渲染
    └── ledger.py               文献台账查重、写入、列出
```
