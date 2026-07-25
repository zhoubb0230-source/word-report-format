# Handoff：LibreOffice 兼容性 & 格式化架构讨论

> **用途**：给**新 session**接续讨论用。它总结了一段跨多轮的排查 + 架构讨论。
> 新 session 请先读本文件 + `CLAUDE.md` + `references/已知陷阱与设计决策.md` 的陷阱 #10。
>
> **当前状态**：代码已做**两轮修复**（已提交推送到分支 `claude/doc-indent-formatting-issue-qtwhur`），
> 但第二轮又暴露出更深的问题，且引出了"要不要重构成归一化/样式化架构"的战略讨论。
> **架构方向尚未拍板**，这正是下一步要继续聊的。

---

## 1. 项目是什么（一句话）

中文 Word 报告的**格式**审查 skill：在副本上自动修正格式 + 加批注。纯脚本流水线，判定阈值全部来自
`spec/format_spec.json`。**红线：只做格式、不改原文、判定值只来自 spec、不手写 fixes.json。**

关键脚本：`05_new_workdir → 10_prepare_input → 20_extract_structure → (27_apply_review) →
30_check_format → 40_apply_fixes → 45_validate_output → 50_finalize → 59_cleanup`。
唯一依赖 lxml。`.doc↔.docx` 转换后端：soffice(优先) / msword(回退)。

---

## 2. 业务问题的起点

用 **LibreOffice(soffice)** 把 `.doc` 转成 `.docx` 再跑 skill，产出在 **Word** 里目录/标题缩进异常；
但**先在 Windows 用 Word 手动把 .doc 转成 .docx** 再跑 skill，效果就**符合预期**。
→ 根因不在 skill 判定逻辑，而在**转换器保真度差异**：两个转换器对同一 .doc 产出**结构不同**的 OOXML。

---

## 3. 已做的代码改动（已推送）

分支 `claude/doc-indent-formatting-issue-qtwhur`，两个 commit，测试 **55 passing**
（`python3 -m unittest discover -s tests -p "test_*.py"`；先 `pip install lxml`）。

**commit 1 `e8e7c5f`**：绝对伴随值 + 认 Contents 样式
- `40_apply_fixes.py`：给字符单位缩进（`firstLineChars`/`leftChars`）**补绝对伴随值**（`firstLine`/`left`），
  以压过 LibreOffice 从样式/编号层继承来的**绝对 hanging**（否则首行缩进被渲染成 -0.74cm 悬挂缩进）。
- 认 LibreOffice 的目录样式名 **`Contents N`**（`headings.py` 的 `_TOC_LVL_RE`/`style_is_toc`、
  `40` 的 `RE_TOC_STYLE_*`/`_toc_style_level` 都加 `contents?` 支路）。

**commit 2 `0fddd86`**：修正 commit1 引入的新问题
- 绝对伴随值**改按【文档默认字号】换算**（新增 `_default_char_unit_hp` 读 `docDefaults` 的 `w:sz`，缺省21），
  而不是段落自身字号——否则 2字符按三号(32)算成 1.13cm、目录三级 2.26cm 并把页码挤换行。
- **目录条目不再挂批注**（`checks.py::_check_toc` 的 fix 置 `comment=False`）：`updateFields` 刷新目录会把
  批注区间孤儿化成空批注。
- 移除已不用的 `_para_eff_size_hp`。

---

## 4. 技术根因（**务必带入新 session**）

### 4.1 缩进表达
- **LibreOffice 只发绝对 twips**（`firstLine`/`hanging`/`left`），**从不发 `*Chars` 字符单位变体**；
  且缩进位置飘忽——有时段落直接属性、有时样式层、有时编号层（继承）。
- **`*Chars` 的字符宽度基准是"文档默认字号"（docDefaults 的 `w:sz`，常见五号10.5pt=21 或 小四12pt=24），
  不是段落自身字号。** 这是"1.13cm/2.26cm 过度缩进"的根因（早期按段落三号32算错）。
- **纯字符 vs 字符+绝对，是个两难**：
  - 只写 `firstLineChars` → 显示"2字符"、自适应 ✓，但**压不过继承的绝对 hanging** → -0.74cm 悬挂缩进 ✗。
  - 字符 + 绝对伴随值 → 压得过 hanging ✓，但**用户的 Word 一旦看到绝对 `firstLine` 就用绝对值、显示厘米**，
    盖掉"2字符" ✗（这就是**当前仍未解决**的"0.85cm 而非 2字符"）。
  - **唯一干净出路 = 从源头清掉继承的缩进（样式层 + 编号层），然后只写字符单位。** ← 归一化方向。

### 4.2 目录（来自用户提供的真实 LibreOffice 输出 XML）
- 目录是 `<w:sdt>`(内容控件) 包 TOC 字段，每条是 `Contents1/2/3` 样式 + `hyperlink` + `IndexLink` 字符样式。
- **每条自带制表位**（这是排版的关键）：
  - `right @8664 leader="dot"` = **页码列**（正文宽 = 页宽11906 − 左右页边距1616×2 = **8674**，@8664≈右边界；三级相同）。
  - `left @840/1050/1260` = **标题文字起点**（逐级右移，这才是目录"逐级缩进"的视觉来源）。
- **skill 破坏点**：目录本是**制表位驱动**的自洽结构；skill 硬套 `leftChars` 左缩进模型 + 绝对伴随值
  + 强制 `updateFields` 刷新 → 刷新后按被改过、且没带 `right @8664 dot` 的目录样式重建条目 → **页码换行**。
  用户手动把制表位调回（"二级6.1字符/三级7字符"，"字符"= Word 的字符宽度度量单位显示）就是补回丢失的制表位。
- **结论：LibreOffice 目录其实已经是对的，应"最小干预 + 保住制表位"，而不是套自己的模型。**

### 4.3 标题（真实 XML）
- 标题**自动编号**：各级 `numId` 不同（18/15/22），**编号不在正文文字里**（靠编号定义生成）。
- 直接缩进是绝对 twips：H2 `ind left=1050 hanging=420`、H3 `ind left=1050 firstLine=643`。
- **继承的 hanging 主要来自编号层（numId 那一层）**，不只是段落直接属性 → 只删段落直接 hanging 不够，
  要连编号层一起归一，字符缩进才能干净生效并显示"2字符"。

---

## 5. 当前仍未解决 / 用户最新反馈

1. **标题、正文**：首行缩进显示 **0.85cm** 而非"2字符"（=绝对伴随值副作用；0.85cm=2字符@小四12pt，说明该 doc docDefaults=24）。
2. **目录**：二级左缩进 0.85cm、三级 1.69cm，应为 **2字符/4字符**；且**页码换行**（制表位被破坏）。
3. **目录空白批注**：commit2 已置 `comment=False` 修复，**待用户验证**。

---

## 6. 架构讨论（**新 session 的核心议题**）

反复"修一个坏一个"已证明：**在 apply 层按转换器特征打补丁，是零和的**（修好 soffice 路径会碰坏 Word 路径）。
讨论过的候选方案：

- **方案 A — 归一化中间层**（推荐方向之一）：判定/应用前插一个 `15_normalize`，把任何转换器产物先转成
  **同一份标准内部形态**，后续只对标准形态校准/测试。**关键设计：幂等 + 信号驱动，对已是标准形态的
  Word 输入 no-op**（这样"针对 LibreOffice 的归一"不会误伤 Word 路径）。canonical form 草案：
  缩进只用字符单位直接属性、清空样式/编号层缩进；角色用多信号识别并盖规范标记；目录保留制表位与字段、
  只统一字体字号、不套 leftChars、不强制刷新；字体别名归一；清冗余 run。
- **方案 B — 固定用 Office/Word 转换**（性价比最高，若在 Windows）：把 `docconv.py` 后端偏好改成
  "有 Word 就优先 Word"。因为**渲染目标就是 Word，同引擎转换最忠实**，整类差异基本消失（用户已实测 Word 转就正常）。
  缺点：绑 Windows+Word，云端/Linux/CI 不可行、不可复现。→ 适合"主要在 Windows 跑"的场景；soffice 仅作无 Word 环境兜底。
- **方案 C — 样式化归一**（用户最新提议"先生成标准模板、再迁移内容"的**安全变体**）：
  - **不要**"把整篇内容搬进新模板"——内容保真风险（表格/图片/文本框/页眉页脚/分节/封面/域/脚注）与"不改原文"红线冲突；
    且角色判定本就不确定（陷阱 #5：pattern 标题只批注），压全篇输出太脆。
  - **正解 = 把"标准模板"理解成"标准 styles.xml"**：注入 canonical 样式 → 给每段**指派**正确样式 → **清冲突的直接覆盖**，
    **内容原地不动**。= `_patch_toc_styles`（现在就在注入目录样式）的**推广**。能一次绕过字符/目录/转换器差异，且守住红线。
    代价：更吃角色判定准确度（判错=套错样式，但内容还在、批注仍可标注，后果可控）；保留"存疑只批注不指派"的安全阀。

**支柱（任何方案都该配）**：① **golden-master 校准语料**（收集多份真实 .doc，soffice 转换后快照归一结果并断言）——
应对"别的 doc 会冒新问题"；② 固定 soffice 版本 + 字体，`commentwriter.py` 的**随机批注 ID 改确定性**（现用 `random._hex8()`），
以求可复现。

---

## 7. 关键约束 / 边界（别忘）

- **真正判据是 Word 的渲染，而它跑不进流水线** → 没有绝对"一劳永逸"；每条归一规则都是"对 Word 的最佳建模"。
  可达到的"彻底" = **单点收口 + 会长大的校准语料**，新差异 = 归一层加一条规则 + 一个样例。
- **本沙箱里 soffice 无法加载任何文件**（环境限制，非文件问题），**无法在此环境实测 LibreOffice 转换或 Word 渲染** →
  所有诊断都是"代码推导 + 用户实测反馈"。数值（0.74/0.85/1.13/1.69/2.26cm 等）都来自用户 Word 里的观测。
- **"不改原文"红线**：任何"重造/迁移"方案都受它约束。
- **角色判定本就不确定**（outline/style 可信，pattern 只批注不自动改）——样式指派/重造类方案都受制于它。

---

## 8. 下一步（供新 session 起步）

- **需要用户提供**：几份**有代表性的真实 .doc**（不同模板/来源，含目录、多级标题、图表题注、封面、表格），
  用于校准 canonical form + 建 golden-master 语料。（沙箱 soffice 跑不动转换，但可分析用户转好的 docx 的 XML。）
- **待决策**：走 B（固定 Word 转换）？还是 A/C（归一化/样式化）？还是两条腿（Windows 优先 Word + soffice 路径配归一化）？
- **低风险起步建议**（零行为改变、零回归）：阶段0 = 建 golden-master 骨架；阶段1 = 把散在
  `checks/headings/apply` 的兼容逻辑收拢成一个 `compat` 模块（纯搬移）。做完能清楚看到"还剩多少差异要归一"。

---

## 9. 代码位置索引

- `scripts/40_apply_fixes.py`：`_char_twips`、`_default_char_unit_hp`、`_set_first_line_and_clear_left`、
  `_patch_toc_styles`、`_set_update_fields`（updateFields=true 的取舍）、`RE_TOC_STYLE_*`/`_toc_style_level`、
  `CommentWriter` 用法、`pending_comments`(每段合并成一条批注)。
- `scripts/lib/checks.py`：`check_paragraph`(按 region+角色分派)、`_check_toc`、`_check_first_line`、
  `_check_no_indent`、`continuity`(编号连续性/自动编号交给 Word)、`_mk_format`。
- `scripts/lib/headings.py`：`_TOC_LVL_RE`、`style_is_toc`、`toc_level_from_style`、`infer_heading_level`、
  `CAPTION_STYLE_HINTS`、`ANY_LABEL_RE`。
- `scripts/lib/docxcommon.py`：`read_ppr`(读 ind 各属性)、`StyleResolver`(样式链解析有效属性)、
  `load_numbering_levels`、`INDENT_KEYS`。
- `scripts/lib/docconv.py` / `soffice.py` / `msword.py`：转换后端（soffice 优先、msword 回退）。
- `spec/format_spec.json`：全部规范值（headings/body/toc/caption/cover…；toc.indent_chars_by_level 0/200/400）。
- `references/已知陷阱与设计决策.md`：陷阱 #10（本次 LibreOffice 兼容）。
- `tests/`：`test_checks.py::TestLibreOfficeTocStyle`、`TestTocExcluded::test_toc_fix_carries_no_comment`；
  `test_e2e.py::TestApplyAbsoluteIndentCompanion`、`::TestLibreOfficeHeadingIndentPipeline`、`::TestDefaultCharUnit`(在 Companion 类内)。

---

## 10. Git 状态

- 分支：`claude/doc-indent-formatting-issue-qtwhur`
- 已推送 commit：`e8e7c5f`（绝对伴随值 + 认 Contents）、`0fddd86`（按默认字号换算 + 目录不挂批注）、
  加本 handoff 的 commit。
- 测试：55 passing。

---

## 11. 严格-spec 判定原则（已与用户拍板，尚未落代码）

方案 C（样式注入）方向下，用户明确了合规判据。这几条一旦落代码就 graduate 进
`已知陷阱与设计决策.md`；在此之前先记在这里，别让文档跑在代码前面（防漂移）。

1. **合规 = 结构谓词，不是"位置差不多"**。以首行缩进为例，合规 ⇔ `firstLineChars = spec 值`
   **且无绝对伴随值**（`firstLine`/`left` 的 twips 缺席或 0）**且整条 cascade 无残留 hanging**。
   `0.85cm`（带绝对 `firstLine`）即使位置对，也判**不合规**——用户裁决"必须严格遵守 spec"。
   好处：该判据纯从解析后 XML 可判，`45` 步能确定性自证"产物 100% 合规"，不必渲染 Word。
2. **绝对伴随值 hack 在严格-spec 下自身违规 → 必须删 → 完整方案 C 从"可选"变"必需"**。
   唯一合规产物 = "字符单位 + 无伴随值 + 无残留 hanging"，只有钳住样式层+编号层+直接层才产得出。
3. **检测必须把编号层折进有效值**。现在 `StyleResolver` 只覆盖 docDefaults+样式链+直接；
   标题 hanging 主要来自编号层（4.3），漏看编号层会把"hanging 藏在编号层"的标题**误判成合规**。
   这份"含编号层的统一 cascade resolver"一石三鸟：检测 / no-op 判定 / 45 步坍缩不变量。
4. **批注只描述"违规"、不描述"手段"**。写"标题应为黑体三号（现宋体）"，不写"已套用标题2样式"。
   锚点按段落（套 run 文本区间），文本来自 checks 的违规串——二者都在方案 C 改动面之外，
   换 apply 引擎不动审查界面。唯一标不了的是样式层全局注入（无段落锚点），靠"幂等+只对可信
   段落指派"保证不产生"改了却没提示"的盲区；存疑段落走"只批注不指派"安全阀（陷阱 #5）。

**无语料怎么办（用户收不到足够 LibreOffice 语料）**：语料原本只给"格式应用"兜底，而严格-spec 把
格式应用变成可构造证明的结构谓词 → 不需要语料。用"对抗性合成 fixture（同一竞争值分别塞进样式层/
编号层/直接层）+ cascade 坍缩不变量"替代 golden-master，沙箱可跑、按构造穷尽"层×属性"。真实 .doc
只需 1–3 份做角色判定冒烟 + 一次性人肉确认"字符-only 表达在 Word 里真显示成 N 字符"。

## 12. 17 项问题分诊（2026-07-25，用户提供的实测新问题）

**心智模型**：方案 C 只改"怎么把判定结果干净严格地落 XML"，**不长出新判定项**。新增"要检查 XX"一律
要动 `checks.py`+`spec`；C 只覆盖"判定项已存在、但因继承/共享应用不干净"的 bug。

**A 纯规则值变更（改 spec，代码已能消费）**
- #1 密级/文本编号 → 方正黑体_GBK（`cover_classification.east_asia`；字号待确认仍三号）
- #4 目录 → 小三（`toc.size_hp` 32→30，`_patch_toc_styles`/`_check_toc` 读 spec 自动跟随）
- #10 字号严格不换算磅值 → 代码已满足（`_check_font_size` 精确相等无容差），记为原则确认

**B 新增判定项（必动 checks/spec，C 不覆盖）**
- #2 题目数字用西文 + 题目字体→方正小标宋_GBK（title 分支现刻意不套西文=陷阱#7，需对 title 破例）
- #3 题目下元素→2倍行距（改 `cover_field` 行距 + 扩展 apply 支持 `lineRule=auto`；作废陷阱#1 的 29.4磅固定值）
- #5 目录制表位 一/二/三级 4/5/6字符（现 `_check_toc` 明确不碰 `w:tabs`；tab 可入 C 注入的目录样式）
- #6 一~四级标题行距固定28磅（heading 分支现无行距检查，复用 `line_spacing`=560）
- #7 表格第一行标题居中（需 `20_extract` 标记 cell 行号 + `_check_table_body` 加居中）
- #8 图/表标题 仿宋三号+28磅（`_check_caption_format` 现只居中+清缩进；spec caption 加字体字号行距）
- #9 表格内容 四号+28磅+无任何缩进（四号=28 已有；新增行距+清缩进，清缩进复用 C 钳位）
- #11 标题序号后加制表符、tab位固定4字符（renumber token 后插 `\t` + 写 tab@4字符）
- #15 正文去段前/段后（新增清 `spacing before/after`）

**C1 方案 C 能覆盖（继承没钳干净）**
- #12 正文首行缩进偶现变悬挂 / #14 四级标题左缩进+悬挂没清：同根因——继承 hanging/左缩进落在
  编号层/样式层、段落读不到，"检测到才清"就漏。完整 C（检测含编号层 + 一律钳 canonical + 显式
  `hanging=0`）系统性消除。C 落地前止血：`_set_first_line_and_clear_left` 无条件显式写 `hanging=0`。

**C2 方案 C 不覆盖，单独修（识别/编号/共享突变）**
- #13 标题、图表标题误识别 → 识别层，C 帮不上且更依赖它；提升 `headings.py` 启发式 + 更靠 27 模型复核
- #16 空内容图标题识别/编号失败 → 识别+renumber 逻辑 bug
- #17 改一个标题行距、同级全变 → **共享对象突变**：行距被写进共享样式/编号层。**⚠️ 直接命中方案 C
  头号风险，C 落地必须"克隆而非改共享"（独立样式 / clone abstractNum），否则把偶发放大成系统性。**

**待用户确认（阻塞落地）**：① #5 4/5/6字符=标题文字左制表位？是否同时保住页码 `right+点线号` 制表位（防
页码换行）；② #3 确认 `lineRule=auto/line=480` 且取代 29.4磅固定值；③ #11 四级都加？与"首行缩进2字符"
如何共存；④ #1 密级/文本编号字号是否仍三号。
