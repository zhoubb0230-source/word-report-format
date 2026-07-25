# Handoff：方案C（样式注入）实施交接

> **用途**：给**新 session 从零开始写方案C**用。冷启动只读本文件即可动手；需要背景时再按下方
> 「延伸阅读」查。
>
> **一句话**：把"规范格式值"从**段落直接属性**搬到**注入的 canonical 样式**里，给每段**指派**正确样式、
> **清掉冲突的直接覆盖**、**钳住编号层**，内容原地不动。这样一次绕过字符/目录/转换器差异，且在无 Word/
> 无语料的沙箱里也能确定性自证合规。

---

## 0. 现状（起点）

- 分支：`claude/scheme-c-style-injection-h3apcl`（已推送）。最新提交：
  - `03a68cb` 批次A（字体/字号/行距/间距类新规则）
  - `06e3b4a` #5 目录制表位（写进 `_patch_toc_styles`＝样式层，方案C第一块）
- **阶段1 已落地**（分支 `claude/scheme-c-codegen-wf5tw6`，零行为改变、全程回归）：
  - `docxcommon.StyleResolver.resolve_cascade()`：含**编号层**的统一 cascade
    （docDefaults→样式链→**编号层**→直接），正确优先级（编号层压过样式、直接压过编号层）。旧
    `resolve()` **保留不动**（阶段1 不改判定路径）；`20_extract_structure.py` 也**暂不**改用它。
  - `scripts/lib/cascade.py`（归一模块）：`effective_paragraph_props()` 便捷装配有效属性、
    `resolve_ppr_with_provenance()` / `contested_indent_report()` 产出**每个属性由哪层供给**的
    provenance——这是"量化还剩几层没钳干净"的量尺（apply 后重跑，凡有效值仍由 `numbering`/`direct`
    供给的 canonical 属性＝没钳干净、会在 Word 泄漏）。
  - `tests/test_cascade.py`：对抗性合成 fixtures（同一竞争值分塞样式/编号/直接层，断言坍缩到正确层，
    并对照旧 `resolve` 漏看编号层）。
  - **已知取舍**：编号层只折 pPr（缩进），**不折** rPr（与 `load_numbering_levels` 一致）；此为刻意
    限制，别当 bug"补全"——真要折 rPr 需同步扩 `load_numbering_levels` 的返回并验回归。
- **阶段2 首块已落地**（同分支，§6 决策已敲定，见下）：
  - **编号层甲法克隆钳**：`40_apply_fixes.py::_clamp_numbering_indent`——对"自动编号且拿到缩进修复"的段落
    **克隆 abstractNum**（新 numId、丢 nsid）、中和克隆级别缩进、把段落 numPr 改指克隆，**不原地改共享**
    （§2.5，根治 #17 与 #12/#14 的"编号层继承 hanging"这一半）。自动编号段的直接缩进改写为**纯字符单位**
    （`_set_first_line_and_clear_left(char_only=True)`，**不补**绝对伴随值——§2.2 严格-spec）。非编号的样式继承
    情形仍补伴随值（`char_only=False`），待全角色样式注入落地后再删（Word-gated）。见陷阱 #10/#11。
  - **目录只认 leftChars**：删掉 `_patch_toc_styles` 写死的制表位（`_set_toc_tabs`/`_toc_text_width_twips` 及
    spec 的 `tab_left_chars_by_level`/`tab_leader` 一并移除）；Word `updateFields` 会自建制表位。测试
    `TestTocTabStops`→`TestTocStyleLeftCharsOnly`。
- 测试：**75 passing**（`pip install lxml` 后 `python3 -m unittest discover -s tests -p "test_*.py"`）。
- 流水线：`05→10→20→(27)→30→40→45→50→59`，唯一依赖 lxml，判定阈值全部来自 `spec/format_spec.json`。
- 17 项实测问题的分诊/进度在 `HANDOFF_格式化架构讨论.md` §12。已落 #1-#4/#6/#8/#9/#10/#15/#5 +
  #12/#14 的"直接 hanging"部分，**+ 甲法落地后 #12/#14 的"编号层继承 hanging"这一半 + #17**（均待 Word 实测）。
  **搁置/留给方案C**：#11（内容层插 `\t`）、#7（提取层标 cell 行号）、#13/#16、#12/#14 的"**样式层**继承 hanging"
  这一半（非编号，需全角色样式注入）。
- **下一步（阶段2 剩余 + 阶段3，需 §5 阶段0 的 Word 人肉验收，沙箱验不了渲染，见 §1）**：把
  `resolve_cascade` 接进 20 的判定路径；把 `_patch_toc_styles` 推广成**全角色** canonical 样式注入 + 指派 +
  清直接覆盖，删掉**非编号情形**的绝对伴随值 hack（§2.2）；阶段3 在 `45` 加坍缩不变量断言（用 `cascade.py`
  的 provenance）。

---

## 1. 铁律 / 沙箱约束（先记牢，别踩）

- **红线**：只做格式、**不改原文**、判定值**只来自 spec**、**不手写 `fixes.json`**。样式注入/指派/清覆盖
  都受"不改原文"约束——动的是格式属性与样式指派，不动正文文字（renumber/删空格等既有内容编辑除外）。
- **本沙箱 soffice 无法加载任何文件**（环境限制），**Word 也不在流水线里** → 无法在此环境实测 LibreOffice
  转换或 Word 渲染。**真正判据是 Word 渲染**，只能靠**用户实测反馈**。所有数值（0.74/0.85cm、8674、
  840/1050/1260 twips 等）都来自用户 Word 观测或代码推导。
- **角色判定本就不确定**（陷阱#5：outline/style 可信，pattern 只批注不自动改）。样式指派**更依赖**它，
  判错=套错样式（但内容还在、批注仍标注，后果可控）→ **保留"存疑只批注不指派"安全阀**。

---

## 2. 已拍板的方案C设计契约（务必遵守，别重新发明）

来自 `HANDOFF_格式化架构讨论.md` §11（严格-spec 四原则）+ 后续讨论：

1. **合规 = 结构谓词，不是"位置差不多"**。首行缩进为例：合规 ⇔ `firstLineChars=spec值` **且无绝对伴随值**
   （`firstLine`/`left` 的 twips 缺席或0）**且整条 cascade 无残留 hanging**。`0.85cm`（带绝对 firstLine）即使
   位置对也判**不合规**（用户裁决"必须严格遵守 spec、字号不换算磅值"）。好处：判据纯从 XML 可判，
   `45` 步能不渲染就断言"产物100%合规"。
2. **绝对伴随值 hack 在严格-spec 下自身违规 → 必须删**。当前 `_set_first_line_and_clear_left`/
   `_patch_toc_styles` 里写的 `w:firstLine`/`w:left` 绝对伴随值，是"字符压不过继承 hanging"的临时补丁；
   完整C钳住样式层+编号层后就不需要它 → **删掉**，只留字符单位，Word 才显示"2字符"而非厘米。
3. **检测必须把编号层折进有效值**。现在 `docxcommon.StyleResolver.resolve()` 只合并 docDefaults→样式链→
   直接，**不含编号层**；`20_extract_structure.py` 只对 `auto_num` 段落把编号级 `ind` 部分兜进 `eff`。
   标题 hanging 主要来自编号层（§4.3），漏看=把"hanging 藏编号层"的标题误判成合规。**做一份含编号层的
   统一 cascade resolver**，同时服务：检测 / no-op 判定 / `45` 坍缩不变量。
4. **批注只描述"违规"不描述"手段"**。写"标题应为黑体三号（现宋体）"，不写"已套用标题2样式"。锚点按段落
   （套 run 文本区间）、文本来自 checks 违规串——二者都在C改动面外，换 apply 引擎不动审查界面。样式层
   全局注入无段落锚点：靠"幂等 + 只对可信段落指派"保证不产生"改了却没提示"的盲区。
5. **⚠️ 克隆而非改共享（C头号风险，对应 bug #17）**。`numbering.xml` 的 `abstractNum` 级别、以及样式，
   都可能被多段共享。**原地改共享对象会溢到非目标段落**（#17"改一个标题行距、同级全变"就是这类）。
   要钳编号层时**克隆该 abstractNum（新 numId）让目标段 numPr 指向克隆**，或改用"直接层显式 `hanging=0`"
   压制（直接逐属性压过编号层，但 `firstLine`+`hanging=0` 会互斥抵消，需验证 Word 行为）。**别原地
   mutate 共享 abstractNum/共享样式。**

**幂等 + 信号驱动**：只有当某段有效解析值 ≠ canonical 时才动它；已合规就 no-op。这样 Word 原生（本就
合规）输入零改动——保护 Word 路径，也守"合规不动"。

---

## 3. 什么会被C改、什么绝不动（避免误伤已落成果）

| 层 | 例子 | C 怎么处理 |
|---|---|---|
| **判定层** `checks.py` | `check_paragraph`、`_check_*`、continuity | **不动**（架构中性，是C复用资产） |
| **spec 值** | `format_spec.json` 全部 | **不动**（C 读同一份） |
| **内容编辑** | renumber、删空格、#11 的 `\t` | **不动**（跟样式无关） |
| **check 层测试** | `tests/test_checks.py` 判定断言 | **不动** |
| **apply 写入位置** | `_set_line_exact`/`_apply_run_props`/`_clear_space_before_after`/`_set_first_line_and_clear_left`/`_patch_toc_styles` | **重指向**：把值写进**注入样式**、清直接覆盖；**同一套原语**（`_set_fonts`/`_set_size`/`_set_line_exact`/写 `w:ind`/`w:tabs`），换挂载元素 |

**结论**：批次A/#5 的判定+spec+内容编辑全部是C的地基,不会白做;只有"往段落直接写"的薄薄一层被
重指向到样式。**别删 checks/spec/内容编辑;别把 `_patch_toc_styles` 推倒——它是C的正例种子,往上长。**

---

## 4. 种子：`_patch_toc_styles` 就是方案C的雏形

`scripts/40_apply_fixes.py::_patch_toc_styles` 现在做的正是C：**注入/覆盖 TOC1/2/3 样式的字体/字号/缩进/
制表位**，让 updateFields 刷新后仍合规。**推广它**＝对每个角色（body / H1..H4 / caption / cover-title /
cover-field / cover-classification / toc1..N）都：① 注入一份完全指定的 canonical 命名样式；② 给可信识别
的段落指派 `pStyle`；③ 清直接覆盖 + 钳编号层。取值全来自 spec（守红线）。

---

## 5. 建议实施阶段（低风险起步，逐步长出C）

- **阶段0（人肉一次性）**：构造一份 **canonical 参考 .docx**——每角色一段、套注入样式。用户在 **Word** 里
  打开、按清单确认"每角色渲染正确 + 首行缩进显示为'N字符'而非厘米 + 目录页码不换行"。这是整套方案
  唯一的人肉环节，替代收集大量语料（用户明确收不到足够 LibreOffice 语料）。
- **阶段1（纯代码、零行为改变、可全跑回归）**：
  ① 写**含编号层的统一 cascade resolver**（docDefaults→样式链→**编号层**→直接，逐属性合并）；
  ② 把散在 `checks/headings/apply` 的兼容逻辑收进一个 `compat`/`normalize` 模块（纯搬移）；
  ③ 建**对抗性合成 fixtures**：同一竞争值（hanging/左缩进/字号）分别塞进样式层/编号层/直接层，
  手搓最小 docx（参考 `tests/helpers.py::build_docx`），断言 apply 后**坍缩到 canonical**。
  做完能量化"还剩几层没钳干净"。
- **阶段2（把C做实）**：把 `_patch_toc_styles` 推广成"注入全部 canonical 样式 + 指派 + 清直接覆盖 +
  **钳编号层（克隆而非改共享，见 §2.5）**"，并**删掉绝对伴随值 hack**（§2.2）。此步溶解遗留 #1(0.85cm)、
  #2(页码换行)，并根除 #12/#14 里"继承 hanging"这一半、修掉 #17。
- **阶段3（无语料安全网）**：在 `45_validate_output.py` 加**坍缩不变量断言**——用统一 resolver 重解析每段
  owned 属性，断言 `==canonical` 或该段=存疑/只批注。泄漏当场硬失败（退2），不靠语料覆盖。

---

## 6. 开放决策（2026-07 已由用户敲定）

- **目录合规谓词** → **只认 `leftChars`(0/200/400)**。制表位不写进样式（Word `updateFields` 自建、写死会被盖掉）。
  已删 `_set_toc_tabs`/`_toc_text_width_twips` 及 spec 的 tab 字段。
- **"渲染对但用直接属性表达"算不算合规** → **不算，严格按样式规范**（canonical 值必须由注入的命名样式承载、
  直接覆盖清掉）。接受随之而来的 churn。
- **编号层钳法** → **克隆 abstractNum（甲法）**，不用直接层 `hanging=0` 压制（避免 `firstLine/hanging` 互斥雷、
  且天然隔离共享 #17）。已实现 `_clamp_numbering_indent`；**仍待用户 Word 验收**能否真显示"2字符"且压住继承。

---

## 7. 代码位置索引（动手前先定位）

- `scripts/40_apply_fixes.py`：`_patch_toc_styles`(种子)、`_set_first_line_and_clear_left`(绝对伴随值 hack 在此)、
  `_set_line_exact`/`_clear_space_before_after`/`_apply_run_props`(要重指向的直接写入)、`_toc_text_width_twips`/
  `_set_toc_tabs`(#5)、`_get_or_make`(CT_PPr 子元素排序：tabs 在 spacing/ind 之前)、`pending_comments`(每段一批注)。
- `scripts/lib/docxcommon.py`：`StyleResolver.resolve`(**缺编号层**，要扩)、`read_ppr`(已含 spacing before/after)、
  `load_numbering_levels`、`INDENT_KEYS`。
- `scripts/lib/checks.py`：`check_paragraph`(判定层，**不动**)、`_check_line_spacing`/`_check_space_before_after`
  (批次A新助手)、`_check_first_line`(#12/#14 直接 hanging 部分)。
- `scripts/20_extract_structure.py`：`eff` 组装（含 spacing before/after）、auto_num 段兜编号层 ind（部分）。
- `spec/format_spec.json`：全部规范值；`toc.tab_left_chars_by_level`/`tab_leader`(#5)。
- `tests/`：`test_checks.py::TestNewFormatRules2026`(批次A)、`::TestTocTabStops`(#5)、`test_e2e.py`(全流水线)。

---

## 8. 延伸阅读（按需）

- `HANDOFF_格式化架构讨论.md`：§4 技术根因（缩进表达/目录/标题的真实 XML）、§6 方案A/B/C 对比、
  §11 严格-spec 四原则、§12 17项分诊+进度。
- `CLAUDE.md`：开发环境、回归测试、流水线陷阱、代码约定、提交规范。
- `SKILL.md`：完整工作流与严格边界（权威）。
- `references/已知陷阱与设计决策.md`：只看顶部「索引」表定位，别整篇读；陷阱 #1/#5/#7/#10 与C最相关。
- `references/format_spec.md`：人类可读规范（权威仍是 JSON）。

---

## 9. 起步动作清单（新 session 照做）

1. `pip install lxml`；`python3 -m unittest discover -s tests -p "test_*.py"` 确认 65 passing 基线。
2. 读本文件 + `HANDOFF_格式化架构讨论.md` §4/§11。
3. 与用户确认 §6 三个开放决策，并请用户准备阶段0 的 Word 人肉验收。
4. 从**阶段1**动手（统一 cascade resolver + 对抗 fixtures + compat 模块），零行为改变、可全程回归。
5. 每落一步同步 `SKILL.md`/`references/format_spec.md`/spec 的 `source`，避免文档-代码漂移（CLAUDE.md 要求）。
