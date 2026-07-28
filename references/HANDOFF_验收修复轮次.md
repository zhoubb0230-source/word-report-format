# Handoff：验收—修复循环的接力棒（新 session 冷启动读这一份）

> **用途**：方案C 阶段2/3 已全部落地，项目进入**「用户在 Word 里验收 → 报问题 → 修 → 再验收」**
> 的循环。这份文件给接手的新 session：现在做到哪、**还没确认的第一优先事项是什么**、怎么复现、
> 以及哪些坑绝对不能踩回去。
>
> **一句话**：功能都在了，卡在**「用户打开成品，Word 提示『发现无法读取的内容』」**这一个问题上——
> 已修掉四类成因，但**尚未拿到用户确认**；下一轮第一件事就是让用户跑 `diagnose_docx.py` 把真凶定位。

分支：`claude/doc-grid-plan-c-fom917`　　回归：**116 passing**
（`pip install lxml` 后 `python3 -m unittest discover -s tests -p "test_*.py"`）

---

## 0. 三分钟上手

```bash
pip install lxml
python3 -m unittest discover -s tests -p "test_*.py"     # 应 116 passing
python3 scripts/make_canonical_reference.py /tmp/ref.docx # 人肉验收参考件（可选）
```

**必读顺序**：本文件 → `references/已知陷阱与设计决策.md` 的**索引表**（只看表，按需展开条目）
→ `CLAUDE.md`（开发约定）。`SKILL.md` 是运行工作流的权威。**别整篇读陷阱文件**，18 条里单次任务
通常只需 1–2 条。

**架构一句话**：判定层（`scripts/lib/checks.py`，纯函数出 `fixes.json`）与应用层
（`scripts/40_apply_fixes.py`，改 XML）分离；规范值全部来自 `spec/format_spec.json`；
canonical 样式的**唯一定义**在 `scripts/lib/canonstyles.py`（参考件生成器与流水线共用同一份字符串）。

---

## 1. ⚠️ 第一优先：Word「发现无法读取的内容」尚未确认修好

用户连续两轮反馈打开成品时 Word 报此错。已定位并修复**四类**成因，但**沙箱里没有 Word**，
最终判据只能是用户实测（这是本项目的长期约束，见 §4）。

已修的四类（全部属于「**良构 ≠ 合法**」——XML 解析没问题、zip 也完好，但 Word 拒绝打开）：

| # | 成因 | 状态 |
|---|---|---|
| 1 | `<w:suff>` 写在 `<w:numFmt>` 之前（违反 CT_Lvl 的 sequence） | 已修，第二轮引入、第三轮修 |
| 2 | `<w:jc>` 在 `<w:ind>` 前、`<w:outlineLvl>` 在 `<w:spacing>` 前、往只有 `<w:sz>` 的 Normal 里 append `<w:rFonts>` | 已修 |
| 3 | `docGrid` append 到 `printerSettings` 之后、`updateFields` 被塞到 settings 最前、`w:num` 落到 `numIdMacAtCleanup` 之后 | 已修 |
| 4 | **样式重名**（注入的"标题 1"/"图标题"与模板已有样式撞车；Word 要求样式名唯一） | 已修：样式名加 `FGW` 前缀 + 注入时兜底改名 |

**下一轮开场就该做的事**（不要先猜，先取证）：

```bash
python3 scripts/diagnose_docx.py <用户手上那个报错的成品.docx>
python3 scripts/diagnose_docx.py <对应的原始文档.docx>     # 对照：原件本来干不干净
```

`diagnose_docx.py` 是这轮新加的**只读**排查工具，专查 XML 良构检查发现不了的那类损坏：
元素顺序违反 sequence、样式名/styleId 重复、`pStyle`/`numId`/关系 `r:id` 悬空、缺 Content-Type。
输出 JSON、退出码 2 表示有问题。**两个文件都要查**——如果原件本来就带问题，那就不是我们引入的。

若 `diagnose_docx.py` 报 **clean 但 Word 仍报错**，说明成因不在已覆盖的检查范围内，
下一步建议：把成品在 Word 里"另存为"后用 `diagnose_docx.py` 对比两份的差异（Word 修复后会
重写它认为非法的部分），差异处即真凶。

---

## 2. 已实现的功能全景（都已跑通并有测试锁定）

**方案C 阶段2/3（样式注入）**——规范值不再写成段落直接属性，而是由**注入的 canonical 命名样式**
承载，段落只指派 `pStyle` 并清掉样式已承载的直接属性。45 步有**全坍缩不变量**兜底（属性仍由
`direct`/`numbering` 供给且值不合规即硬失败退 2）。契约细节见陷阱 #12。

三轮验收累计落地：

- **文档网格**每页 41 行 / 行距 15.6 磅（`docGrid` + settings 的 compat 块 + `Normal`=五号）；
- **目录**制表位按字符数写进 TOC 样式（点线、页码不换行）——已通过验收；
- **首行缩进**只写字符单位，Word 显示"2 字符"——已通过验收；
- **一~四级标题加粗**，加粗由样式承载（连 run 上的直接加粗、显式取消加粗、抢戏的字符样式一起清）；
- **标题/图表标题删首尾空格**；
- **图/表标题改用 Word 自动编号**（`图%1`/`表%1`、`suff=tab`），编号**挂在样式上**（套上样式即生成
  编号）——图表编号显示已通过验收；
- **封面外空行**套正文样式；**图片/文本框段落**不套（固定行距会把图裁成一行）；
- **表格**默认值（`tblInd`/`tblCellMar`/`vAlign`）。

---

## 3. 本轮修掉但**尚未经用户确认**的三件事

下一轮验收请重点看这三项：

1. **标题编号逐级归零**：二级在其一级下从头数、三级在二级下、四级在三级下。
   根因是模板编号级别带 `<w:lvlRestart w:val="0"/>`（永不重新计数），克隆编号定义时已去掉它。
   *（上一轮的"编号错位"则是克隆钳按 `numId` 分组把一条多级列表劈成几条，已改为按 `abstractNum`
   归组——两个问题不同根因，别混。）*
2. **图片不再被裁**：图片/文本框段落不指派样式。
3. **表标题编号与文字粗细一致**：自动编号的渲染取自段落标记的 `rPr`，指派时清掉它的加粗。

---

## 4. 铁律与长期约束（别踩）

- **红线**：只做格式、**不改原文**、判定值**只来自 spec**、**不手写 `fixes.json`**。
- **沙箱没有 Word / soffice**：渲染类结论**只能靠用户实测**。凡是"Word 里看起来如何"的判断，
  写代码时就要想好怎么让用户一眼验证，别自己拍脑袋下结论。
- **良构 ≠ 合法**：`45` 的 XML 良构检查发现不了元素顺序、样式重名这类问题。往 XML 里插元素
  一律走 `_get_or_make`（自动查 `docxcommon.ELEMENT_ORDER` 按序插），**别 append 了事**。
- **安全阀**（陷阱#5）：仅凭形状认出、没有样式/大纲级别撑腰的标题与图表标题（`pattern`），
  **绝不自动改文字**，也不指派 canonical 样式（否则下一轮它会凭样式"转正"、绕过安全阀）。
- **共享对象不可原地改**：`abstractNum`、字符样式都可能被多段共享，要改就**克隆**或**摘引用**，
  原地 mutate 会溢到无关段落（老 bug #17）。
- **同一规则只写一份**：canonical 样式、区域划分、字符样式判断都已收敛到共享模块。
  再出现"两处各写一遍"必然漂移——这是本项目栽过最多次的地方。

---

## 5. 代码地图（动手前先定位）

| 关注点 | 位置 |
|---|---|
| canonical 样式的唯一定义（角色→样式、`governs` 表） | `scripts/lib/canonstyles.py` |
| 判定层（出 fixes，**架构中性、尽量别动**） | `scripts/lib/checks.py`（`paragraph_role` 是判定与指派的共用分派） |
| 应用层（注入/指派/清覆盖/编号/网格/表格） | `scripts/40_apply_fixes.py` |
| 元素顺序表、字符样式判断、cascade | `scripts/lib/docxcommon.py` |
| 产物自检（含元素顺序、全坍缩不变量） | `scripts/45_validate_output.py` |
| 任意 docx 的损坏排查（新） | `scripts/diagnose_docx.py` |
| 人肉验收参考件生成器 | `scripts/make_canonical_reference.py` |
| 规范值（唯一权威） | `spec/format_spec.json`；人类可读镜像 `references/format_spec.md` |

---

## 6. 维护提醒

- `references/已知陷阱与设计决策.md` 现有 **18 条**，超出它自己定的"约 10 条以内"。下次改动时
  顺手审一遍哪些可以退役（判据是：决策被推翻、或对应代码已删除）。**别为了压条数删掉仍然生效的
  条目**——那些是防回退的唯一记录。
- 每落一个改动，**同步** `SKILL.md` / `references/format_spec.md` / spec 的 `source` 描述，
  并给"别回退"级别的决策补一条陷阱记录 + 一个锁住它的测试（CLAUDE.md 的硬要求）。
- 历史 handoff：`HANDOFF_方案C实施.md`（方案C 背景与契约）、`HANDOFF_阶段2_样式注入.md`
  （阶段2 实施记录）、`HANDOFF_格式化架构讨论.md`（早期技术根因与 17 项分诊）。
  都是**背景资料**，现状以本文件为准。
