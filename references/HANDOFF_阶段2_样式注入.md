# Handoff：方案C 阶段2（全角色 canonical 样式注入）

> **用途**：给**新 session 冷启动做阶段2**用。冷启动读本文件 + `make_canonical_reference.py` 即可动手。
> 背景/契约见 `HANDOFF_方案C实施.md`（不重复）；Word 保真真值的"为什么"见 `已知陷阱…md` #10/#11/#12。
>
> **一句话**：阶段0/1 + 阶段2首块（甲法）已完成并经用户 Word 验收；剩阶段2 主体——把
> `_patch_toc_styles` 推广成**全角色 canonical 命名样式注入 + 指派 pStyle + 清直接覆盖**，并删掉**非编号
> 情形**的绝对伴随值 hack。canonical 样式的**精确参数已由用户逐像素验收**，权威真源是
> `scripts/make_canonical_reference.py`。

---

## 0. 现状（分支 `claude/scheme-c-codegen-wf5tw6`，79 passing）

`pip install lxml` 后 `python3 -m unittest discover -s tests -p "test_*.py"` → 79 passing。已落：

- **阶段1**：`docxcommon.StyleResolver.resolve_cascade()`（含编号层的统一 cascade）+ `scripts/lib/cascade.py`
  （provenance 量尺）+ `tests/test_cascade.py`（对抗 fixtures）。
- **阶段2 首块（甲法克隆钳）**：`40_apply_fixes.py::_clamp_numbering_indent` —— 自动编号且缩进违规的段落，
  **克隆 abstractNum**（不改共享，根治 #17）、中和级别缩进、段落 numPr 改指克隆；自动编号段写纯字符单位、
  不补绝对伴随值。见陷阱 #11。
- **检测接入**：`20_extract_structure.py` 改用 `resolve_cascade`（正确优先级折编号层）。
- **阶段3 甲法部分**：`45_validate_output.py::_check_numbering_clamp` —— 拿到缩进修复的自动编号段，有效
  hanging 仍由编号层供给＝硬失败退2；样式层继承 hanging＝非致命 note。
- **目录只认 leftChars**（阶段2首块）：`_patch_toc_styles` 删了写死制表位（后被阶段0 推翻，见下 §3）。
- **阶段0 canonical 参考件 + Word 验收（本轮重点）**：`scripts/make_canonical_reference.py` 生成"每角色一段
  套 canonical 样式"的 docx，用户在 Word 里**逐项验收通过**——甲法 2字符首行、目录制表位、文档网格、
  封面西文、单元格边距等全部对齐规范文档。**该生成器就是全角色 canonical 样式的可执行权威规格。**

---

## 1. 铁律 / 沙箱约束（先记牢）

- **红线**：只做格式、**不改原文**、判定值**只来自 spec**、**不手写 `fixes.json`**。
- **沙箱无 Word/soffice**：Word 渲染是唯一判据，只能靠用户实测（阶段0 已建立参考件验收闭环）。
- **角色判错=套错样式**，故**保留"存疑只批注不指派"安全阀**（陷阱#5：outline/style 可信，pattern 只批注）。

---

## 2. 阶段2 主体要做什么

把 `40_apply_fixes.py::_patch_toc_styles`（目前只注入 TOC 样式）**推广成"注入全部 canonical 命名样式 +
给可信识别段落指派 pStyle + 清直接覆盖"**，并**删掉非编号情形的绝对伴随值 hack**（`_set_first_line_and_clear_left`
的 `char_only=False` 分支，见陷阱#10）。落地后：

- 溶解 #12/#14 的"**样式层**继承 hanging"这半、#1(0.85cm)；
- `45` 的 `_check_numbering_clamp` 可升级成**全坍缩不变量**（任一 canonical 属性仍由 `direct`/`numbering`/
  非规范 `style` 供给即硬失败——用 `cascade.py` 的 provenance，已就绪）。

**别推倒既有成果**：甲法克隆钳、检测接入 cascade、45 编号层不变量、判定层 checks/spec —— 都是地基，复用。

---

## 3. canonical 样式精确参数（**权威真源 = `make_canonical_reference.py`**，照它落进 spec/apply）

> 全部经用户 Word 逐项验收。字号单位：sz 半点（三号=32，小三=30，四号=28，五号=21，小一=40）。
> **⚠️ 落地前务必读陷阱 #12**——下列每条背后都有"反直觉、别改回去"的 Word 行为。

### 3.1 文档级（`settings.xml`）—— 网格能否 15.6磅/41行 的关键
- `<w:compat>` 整块照抄：`spaceForUL`/`balanceSingleByteDoubleByteWidth`/`doNotLeaveBackslashAlone`/
  `ulTrailSpace`/`doNotExpandShiftReturn`/`adjustLineHeightInTable`/**`useFELayout`** + 6 个 compatSetting
  （**`compatibilityMode=15`**、overrideTableStyleFontSizeAndJustification=1、enableOpenTypeFeatures=1、
  doNotFlipMirrorIndents=1、differentiateMultirowTableHeaders=1、useWord2013TrackBottomHyphenation=1）。
- `defaultTabStop=420`（**2字符**；标题自动编号后的制表符落这里，别用 640）。
- `drawingGridHorizontalSpacing=105`、`drawingGridVerticalSpacing=156`、`display{H,V}orizontalDrawingGridEvery=2`。
- `noPunctuationKerning`、`characterSpacingControl=compressPunctuation`、`themeFontLang eastAsia=zh-CN`。

### 3.2 节（`sectPr`）
- `pgMar`：上1984/下1814/左1616/右1616/页眉850/页脚992（来自 `spec.page_setup`）。
- `pgNumType start=1`、`cols space=720`、**`docGrid type=lines linePitch=312`**（=15.6磅行网格）。

### 3.3 样式（`styles.xml`）
- **docDefaults**：`Times New Roman`/`仿宋`/**五号(21)** + `<w:lang eastAsia=zh-CN>` + 空 `pPrDefault`。
- **Normal（名"Normal"/正文）＝五号**：**只驱动文档网格字体，别用于正文内容**。
- **FGW正文（正文内容，basedOn Normal）**：仿宋 **三号**、西文 Times、行距固定值28磅(line560 exact)、
  去段前段后、首行缩进2字符。**⚠️ 名字必须是"FGW正文"不能是"正文"**（否则抢网格字体链接）。
- **标题1-4**：黑体/楷体/仿宋/仿宋 三号、西文 Times、line560 exact、首行缩进2字符、outlineLvl 0-3。
- **封面题目**：方正小标宋_GBK 小一(40)、居中、西文 Times。
- **封面密级编号**：方正黑体_GBK 三号、居中、**西文 Times（数字走西文）**。
- **封面要素**：方正黑体_GBK 小三(30)、2倍行距(line480 auto)、两端对齐、首行缩进2字符、西文 Times。
- **图表标题**：仿宋 三号、line560 exact、居中、无缩进。
- **表格内容**：仿宋 四号(28)、line560 exact。
- **目录标题**（"目录"二字）：仿宋 三号、居中、无缩进（非标题、不进目录）。
- **TOC1/2/3**：仿宋 小三(30)、basedOn Normal、leftChars 0/200/400、**无首行缩进**；制表位写进样式——
  **字符单位随 Normal 字号**：Normal=五号→左 **840/1050/1260** + 右点线 **8665**（Normal=三号 则 1280/1600/1920+13203）。

### 3.4 编号（`numbering.xml`）
- **标题多级**：`一、` / **半角** `(一)` / `1.` / **半角** `(1)`，`suff=tab`，级别缩进中和（甲法目标形态）。
- **图/表标题**：`图%1` / `表%1` decimal，`suff=tab`（编号是整体、不可拆选）。

### 3.5 表格
- `tblInd=0`（表格左缩进）、`tblCellMar` 上0/左108/下0/右108（0.19cm）、单元格 `vAlign=center`。
  **`tblInd` 与 `tblCellMar` 是两回事，别混**。

---

## 4. 需要与用户敲定/实测的开放项

- **目录是"域"，用户 F9 更新域后 Word 是否用样式制表位重建正确**？参考件用"域+缓存条目"（条目继承样式
  制表位）验证过静态显示；F9 重建的真机行为仍需用户在真实样本上确认（沙箱验不了）。
- **封面西文推翻陷阱#7 的范围**：确认后要同步改 `checks.py` 的西文规则 + `spec` 描述（阶段2 一并做）。
- **严格-spec churn**：指派 pStyle + 清直接覆盖会 churn 一批"看着对但直接属性表达"的段落（用户已裁定
  按严格样式规范来）——落地后跑用户真实样本确认 churn 可接受。

---

## 5. 起步动作清单（新 session 照做）

1. `pip install lxml`；跑回归确认 79 passing 基线。
2. 读本文件 + `已知陷阱…md` #10/#11/#12 + `HANDOFF_方案C实施.md` §2 契约。
3. 通读 `scripts/make_canonical_reference.py`（全角色 canonical 样式的可执行规格）。
4. 把 §3 参数落进 `spec/format_spec.json`（新增网格/compat/单元格/目录制表位/封面西文字段）+
   `40_apply_fixes.py` 的全角色样式注入/指派/清覆盖；删非编号伴随值 hack。**每步跑回归 + 生成参考件比对**。
5. `45` 升级成全坍缩不变量（用 `cascade.py` provenance）。
6. 每落一步同步 `SKILL.md`/`references/format_spec.md`/spec 的 `source` + `已知陷阱…md`（CLAUDE.md 要求）。
7. 产出物请用户 Word 验收（沿用阶段0 的参考件验收闭环）。
