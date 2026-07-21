---
name: word-report-format
description: "对中文Word 文档（.doc/.docx）进行【格式】审查与自动修正。触发场景：用户要求检查/规范/修正报告的字体、字号、行距、缩进、页边距、标题层级与图表编号，并在不改动原文的前提下生成带批注的格式化版本。严格按内置固定规范执行，禁止模型自行发挥或引用其它规范。仅做格式，不做文字润色或逻辑一致性审查。"
license: Proprietary
---

# 格式审查与自动修正 (word-report-format)

本 skill 对一份中文自评价报告 Word 文档做**格式合规检查**并在**新文档**上自动修正，
给每处修改/提示加批注（批注内容即所违反的规范条目）。

## 严格边界（务必遵守）

1. **只做格式**：字体/字号/行距/缩进/页面设置/标题层级序号/图表序号连续性。不做文字润色、不做逻辑审查。
2. **规范是固定的**：一切判定阈值只来自 `spec/format_spec.json`。**禁止**依据模型自身训练知识发挥，**禁止**引用其它 skill 的规范。
3. **绝不改原文**：所有改动都在解包出的副本上进行，原文件只读。
4. **先判断后修改**：脚本只对「当前生效格式 ≠ 规范」的项产出修改；符合的不动。
5. **不要让模型写脚本 / 手写 JSON**：所有脚本已预置，运行时只调用。`fixes.json` 等中间文件全部由脚本生成，模型不得手工拼 JSON（避免引号未转义、字段名漏 `set_` 前缀等）。
6. **输出命名**：`原文件名_格式化版本_时间戳.原扩展名`（由 `50_finalize.py` 负责，勿自行改名）。
7. **子 agent 委派是硬性要求**：若运行环境支持子 agent，必须走 Path B，把每个分片交给独立子 agent 处理（用于隔离上下文，避免 128k 窗口被 3000 页文档塞爆）。

## 目录结构

```
word-report-format/
  SKILL.md                  ← 本文件
  spec/format_spec.json     ← 固定规范（唯一判定依据）
  scripts/
    00_check_env.py         ← 探测 python/lxml/soffice，判断能否转换 .doc
    10_prepare_input.py     ← 生成 working.docx（.doc 需转换）+ meta.json
    20_extract_structure.py ← 解析生效格式 → structure.json + shards/
    30_check_format.py      ← 格式检查（全量 或 --shard 单片）
    31_check_global.py      ← 全局项：页边距 + 序号连续性 + 内容提示
    35_merge_fixes.py       ← 合并各分片结果为 fixes.json
    40_apply_fixes.py       ← 在新文档上应用 + 加 XAgent 批注 → formatted.docx
    50_finalize.py          ← 按命名规范输出（.doc 输入则转回 .doc）
    lib/                    ← 自包含依赖（lxml，无需 python-docx / defusedxml）
```

## 第 0 步：环境检查与分支（每次必做）

```bash
python scripts/00_check_env.py
```

输出单行 JSON：`{python, lxml, soffice, can_convert_doc, ok}`。

- `ok=false` → 缺关键依赖（python 或 lxml），停止并告知用户环境不满足。
- 用户给的是 **.doc** 且 `can_convert_doc=false` → **停止**，通知用户：当前环境无法转换 .doc，
  无法保证输出类型一致，请改传 .docx 或在支持 LibreOffice 的环境运行。**不要**擅自改用别的格式硬跑。

## 第 1 步：准备输入

```bash
python scripts/10_prepare_input.py <输入文件> <workdir>
```

- .docx → 复制为 `<workdir>/working.docx`
- .doc → 用 LibreOffice 转成 `working.docx`；若不可转，退出码 2 并给出中文提示（见上）。
- 生成 `<workdir>/meta.json`（记录原名、原扩展名、working.docx 路径）。原文件始终只读。

## 第 2 步：抽取结构（解析「生效格式」）

```bash
python scripts/20_extract_structure.py <workdir> [--shard-size 400]
```

关键点（脚本已实现，模型无需关心细节）：
- 通过 `styles.xml` + `basedOn` 继承链解析**生效格式**（含继承来的 `outlineLvl`），不只看段落直接属性。
- 识别封面/目录/正文区域；**目录段落被排除**在标题/图表判定与连续性计数之外，避免假问题与污染。
- 产出：`structure.json`（完整，**留在磁盘、不要读进模型上下文**）与 `shards/shard_NNN.json`（分片切片）。
- 终端只打印计数摘要，供模型判断规模、决定走 Path A 还是 Path B。

## 第 3 步：检查 + 应用（二选一）

### Path A —— 环境不支持子 agent（小文档亦可）

```bash
python scripts/30_check_format.py <workdir>      # 全量：逐段格式 + 页边距 + 连续性 + 内容提示 → fixes.json
python scripts/40_apply_fixes.py <workdir>       # 应用到新文档 + XAgent 批注 → formatted.docx
python scripts/50_finalize.py  <workdir> <输出目录>
```

### Path B —— 支持子 agent（**优先/硬性要求**，适用于 3000 页大文档）

1. 主 agent 先跑全局项（连续性必须跨全文，不能按片算）：
   ```bash
   python scripts/31_check_global.py <workdir>   # → fixes_parts/_global.json
   ```
2. **对 `shards/` 下每个分片，委派一个独立子 agent** 执行（互不共享上下文）：
   ```bash
   python scripts/30_check_format.py <workdir> --shard shards/shard_NNN.json
   # → fixes_parts/<片名>.json（子 agent 只需回传该片修改条数，不要回传全文）
   ```
3. 全部分片回来后，主 agent 合并并应用：
   ```bash
   python scripts/35_merge_fixes.py <workdir>    # 各片 + _global 合并去重 → fixes.json
   python scripts/40_apply_fixes.py <workdir>    # → formatted.docx（含 XAgent 批注）
   python scripts/50_finalize.py  <workdir> <输出目录>
   ```

> 分片检查只做「逐段格式」；页边距与序号连续性是全局的，统一由 `31_check_global.py` 负责，
> 避免跨片时序号计数被切断。两条路径产出的 `fixes.json` 完全一致。

## 产物说明

- `formatted.docx`：已修正 + 每处改动/提示带 `XAgent` 批注；已设 `updateFields`，用户打开时目录会自动刷新。
- `50_finalize.py` 输出：`<原名>_格式化版本_<时间戳>.<原扩展名>`；原为 .doc 时自动转回 .doc。
- 无法自动修的项（封面缺字段、封面绿色底、图表缺内容说明）只以**批注提示**形式给出，不臆改。

## 模型在本流程中的职责（很重要）

模型只做**编排**：按上面顺序调脚本、读计数摘要、决定 Path A/B、委派子 agent、把最终文件交付用户。
**不要**：改写脚本、手写任何 JSON、凭训练知识判断格式对错、引入本 spec 之外的规则。
所有格式判定都是脚本里的确定性逻辑，不依赖模型推理。
