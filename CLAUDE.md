# CLAUDE.md — 开发/维护指引

> 本文件面向**改这份代码**的 agent。**运行 skill 的完整工作流、规范边界、每步做什么**见
> `SKILL.md`（权威来源），此处不重复，只补开发时需要、且不在 SKILL.md 里的东西。
>
> **动手改判定/应用逻辑前先查、改完这类逻辑后回写** `references/已知陷阱与设计决策.md`——记录
> "看起来像 bug、其实是刻意为之、别照直觉改回去"的回退陷阱（封面空格处理、updateFields 弹窗取舍、
> 目录段落排除、pattern 标题不自动改等）。用法：
> - **读**：只看它顶部的「索引」表定位,再局部读相关条目(offset/limit 或 grep 区域关键词),别整篇读;
>   该文件按需读，不进本文件以免每次会话吃 token。
> - **写**：当你做出/推翻一条"回退陷阱"级别的决策时,**必须**按该文件顶部的三层结构归档——追加条目到
>   「生效中」区、同步在「索引」表加一行、被推翻的旧条目压缩一行挪到「已退役」区、「生效中」维持约 10 条内。
>   (只归档"别回退"级别的决策;普通改动交给 git log,不必写这里。)

## 这是什么

一个 Claude skill：对中文 Word 报告做**格式**审查，并在副本上自动修正 + 加批注。
纯脚本流水线（`scripts/`），判定阈值全部来自 `spec/format_spec.json`。核心红线（改代码也要守）：
**只做格式、不改原文、判定值只来自 spec、不手写 `fixes.json`**。细节见 `SKILL.md` 的「严格边界」。

## 开发环境

- 唯一依赖是 **`lxml`**，且**默认没装**——先 `pip install lxml`（`scripts/lib/` 是自包含的，
  不用 python-docx / defusedxml）。
- `python3 scripts/00_check_env.py` 探测 `python/lxml/soffice`；`soffice` 只在 `.doc`↔`.docx`
  转换时需要（`.docx` 全程用不到）。
- 无第三方测试框架，无 `requirements.txt`。测试用 stdlib `unittest`（只依赖 lxml）。

## 回归测试（先跑这个，再手动冒烟）

`tests/` 有一套零依赖回归测试，**改判定/应用逻辑后必跑**：

```
python3 -m unittest discover -s tests -p "test_*.py"
```

- `tests/test_checks.py`：判定层单测，把《已知陷阱与设计决策》里每条"别回退"锁成断言
  （封面题目/字段、目录排除、pattern 只批注、层级判定、图表分组编号、西文字体、页码/目录深度提示…）。
  改了 `checks.py`/`headings.py` 的行为却没同步改这里 = 有测试会红，正是防回退的意义。
- `tests/test_e2e.py`：手搓最小 docx 跑 `05→10→20→30→40→45`，断言输出自检 `ok` 且段落数守恒。
- `tests/helpers.py`：`build_docx()` 造最小 docx、`load_script()` 以模块方式载入带数字前缀的脚本。

新增一条陷阱级决策时，**同时**在 `tests/` 补一个锁它的用例，别只写进散文记录。

## 冒烟测试（手动跑一遍最直观）

最省事的两种方式：

1. **纯逻辑**：直接 import `scripts/lib/checks.py`，构造 `rec` 字典喂给 `check_paragraph(rec, spec)`
   看返回的 fix。`rec` 至少要有 `i / region / eff{east_asia,size_hp,jc,...} / text`，封面还看
   `is_title / cover_role`。不需要真实 docx。
2. **端到端**：手搓一个最小 docx（`[Content_Types].xml`、`_rels/.rels`、`word/document.xml`，
   **别忘了 `word/_rels/document.xml.rels`——`40` 的 CommentWriter 会读它，缺了会报错**），
   然后按顺序（`<workdir>` 用 `05_new_workdir.py` 建的唯一目录，别手写路径）：
   ```
   python3 scripts/05_new_workdir.py                 # 打印 {workdir}；下面的 <workdir> 全用它
   python3 scripts/10_prepare_input.py <src.docx> <workdir>
   python3 scripts/20_extract_structure.py <workdir>
   python3 scripts/30_check_format.py <workdir>     # 全量模式，直接写 workdir/fixes.json
   python3 scripts/40_apply_fixes.py <workdir>       # 产出 workdir/formatted.docx + out_pkg/
   python3 scripts/45_validate_output.py <workdir>   # 自检：ok=true 才算过（坏了退 2）
   python3 scripts/50_finalize.py <workdir> <out_dir> # <out_dir> 须在 workdir 之外
   python3 scripts/59_cleanup.py <workdir>            # 交付后删掉工作目录（过程件全清）
   ```

## 流水线陷阱（真踩过）

- **工作目录唯一 + 交付后清理**：所有过程件都落在 `05_new_workdir.py` 建的**唯一** `<workdir>`
  （`<cwd>/.word_report_work/run_<时间戳>_<随机>/`，`tempfile.mkdtemp` 原子创建，多 session 并发不冲突），
  交付后由 `59_cleanup.py` 整体删除。两条铁律：① `50_finalize.py <workdir> <out_dir>` 的 `<out_dir>`
  **必须在 workdir 之外**——成品若留在 workdir 里会被 `59` 一并删掉；② `59_cleanup.py` 有安全护栏，
  只删名字含 `.word_report_work` 的目录、且不删基目录本身，别绕过它去 `rmtree` 别的路径。基目录/护栏
  逻辑在 `scripts/lib/workdir.py`（`base_dir()` / `is_inside_base()`），改路径约定要同步改这里。
- **区域划分**：`cover` = 第一个 TOC / 第一个标题**之前**的所有段落。测试 docx 若既无目录又无标题，
  封面段落会全部落到 `body`——要放一个带 `outlineLvl` 的标题段来界定封面结束。
- **区域划分 / 分片写盘是共享逻辑**：`tag_regions()` 与 `write_shards()` 都在 `scripts/lib/structure.py`，
  由 `20_extract_structure.py`（抽取时）和 `27_apply_review.py`（复核改层级后重算边界）**共用同一份**。
  27 复核可能提升/取消最早一条标题、从而移动封面/正文边界，必须重跑 `tag_regions`。**别把这两段再各自
  内联回两个脚本**（历史上就是各写一遍、易漂移）。`shard_size` 已持久化进 `structure.json`，27 直接读、
  不再靠"读第一个分片记录数"反推。
- **`30` 有两套模式**：全量模式**直接写 `fixes.json`**；`--shard` 模式写 `fixes_parts/`，之后才用
  `35_merge_fixes.py` 合并。**全量模式下不要跑 `35`**——它从 `fixes_parts/` 读，会把 `fixes.json`
  覆盖掉。分片路径见 SKILL.md 的 Path B。
- `cover_role` 的标题启发式阈值是**字号 ≥ 36 半点（小一）且居中**，或等于 title 的 `size_hp`。
  正常 15 磅（=30 半点）字段行低于此值，不会被误判成标题。
- 输出命名由 `50_finalize.py` 负责（`原名_格式化版本_时间戳.原扩展名`），别在别处改名。

## 代码约定

### `scripts/lib/checks.py`（判定层，产出 fix，不碰 XML）
- `check_paragraph(rec, spec)` 按 `region`（cover/toc/body）+ 角色分派，返回一个 `format` fix 或 `None`。
- 内部统一用一个 **`sets` 字典**累积要设置的属性，`violations` 累积中文说明，最后 `_mk_format(i, sets, violations)`
  组装成 fix；**`violations` 为空则返回 `None`**（＝合规就不动）。
- **新增一种可修项 = 给 `sets` 加一个键**（如 `strip_text`）：① 在 `sets` 初始化处加默认值，
  ② 在对应角色分支里按需赋值 + append 一条 violation，③ 到 `40_apply_fixes.py` 里实现该键的 XML 应用。
- 封面角色：`title` / `classification`(密级·编号) / `field`(项目名称等) / `other`。字体字号取值在
  `spec` 的 `title` / `cover_classification` / `cover_field`。

### `scripts/40_apply_fixes.py`（应用层，直接改 `document.xml`）
- 所有改动写成**段落/run 的直接属性（override）**，覆盖样式继承值；清缩进用「显式置 0」而非删属性
  （否则会露出样式里的缩进）。
- `format` op 的每个 `set_*` / `clear_*` / `strip_text` 键，在 `main()` 的 `op == "format"` 分支里
  各有一段应用逻辑；改 `w:t` 文本时只动内容 run（`_iter_runs` 已跳过文本框 `w:txbxContent`），
  保留行内空格与 `xml:space`。

## 提交

开发分支见任务要求；提交信息用中文、说清「问题 → 方案 → 涉及文件」。改了判定/应用行为时，
**同步更新** `SKILL.md`、`references/format_spec.md`、`spec/format_spec.json` 里对应的 `source` 描述，
避免文档与代码漂移。
