# -*- coding: utf-8 -*-
"""
checks.py — the deterministic rule engine. Pure functions; no I/O, no model.

Every function returns fix dicts with a FIXED schema (generated here, never by a
model) so that field names are always correct and JSON is always valid.

Fix object shapes
-----------------
format fix (auto-fixable, one per paragraph, combines all its violations):
  {"para_index": i, "op": "format",
   "set_east_asia": str|None, "set_ascii": str|None, "set_size_hp": int|None,
   "set_line_exact": int|None, "set_first_line_chars": int|None,
   "clear_left_indent": bool, "clear_right_indent": bool, "set_jc": str|None,
   "rule_id": "combined", "rule_text": "<multi-line 违反规范>", "comment": true}

renumber fix (auto-fixable text change on the leading ordinal):
  {"para_index": i, "op": "renumber_caption", "kind": "figure"|"table",
   "new_num": "2", "rule_id":..., "rule_text":..., "comment": true}
  {"para_index": i, "op": "renumber_heading", "level": 1,
   "new_token": "二、", "rule_id":..., "rule_text":..., "comment": true}

hint fix (NOT auto-fixable, comment only):
  {"para_index": i, "op": "hint", "rule_id":..., "rule_text":..., "comment": true}

section fix (page setup, document level):
  {"op": "section", "set_pgmar": {...}, "rule_id":..., "rule_text":..., "comment": false}
"""
import json
import os
import re

CN_DIGITS = "零一二三四五六七八九"

# Canonical location of the authoritative spec, resolved once relative to this
# file (scripts/lib/checks.py -> ../../spec/format_spec.json). Every pipeline
# stage imports this instead of recomputing the path, so there is one source of
# truth for where the spec lives.
SPEC_PATH = os.path.normpath(
    os.path.join(os.path.dirname(os.path.abspath(__file__)),
                 "..", "..", "spec", "format_spec.json"))


def load_spec(spec_path):
    with open(spec_path, encoding="utf-8") as f:
        return json.load(f)


def load_default_spec():
    """Load the authoritative spec at SPEC_PATH."""
    return load_spec(SPEC_PATH)


def int2cn(n):
    """1..99 -> Chinese numeral (一, 十, 十一, 二十, 二十三 ...).

    Outside 1..99 it falls back to the arabic form: a heading level never
    realistically reaches 100 same-level items, and returning str(n) is far
    better than the IndexError the tens/ones split would otherwise raise at 100+.
    """
    if n <= 0 or n >= 100:
        return str(n)
    if n < 10:
        return CN_DIGITS[n]
    if n == 10:
        return "十"
    if n < 20:
        return "十" + CN_DIGITS[n - 10]
    tens, ones = divmod(n, 10)
    s = CN_DIGITS[tens] + "十"
    if ones:
        s += CN_DIGITS[ones]
    return s


# Section titles that are conventionally UNNUMBERED even when they carry a real
# heading/outline style (摘要/前言/结论/目录/参考文献/…). A confirmed heading
# whose title is one of these is neither auto-numbered nor counted, so it does
# NOT shift the ordinals of the numbered headings around it. Matched against the
# heading text with any leading label and trailing colon/space stripped.
UNNUMBERED_HEADING_EXACT = {
    "摘要", "中文摘要", "英文摘要", "内容摘要",
    "前言", "引言", "序言", "绪论", "绪言",
    "结论", "结语", "结束语",
    "目录", "索引",
    "致谢", "鸣谢", "后记",
    "声明", "独创性声明", "原创性声明", "版权声明",
    "abstract",
}
# Prefixes that start an unnumbered section even with a trailing token
# ("附录A"/"附录一"/"参考文献 [续]").
UNNUMBERED_HEADING_PREFIX = ("附录", "附件", "参考文献", "参考资料")


def _norm_title(text, num_raw):
    """Heading text minus any leading numbering label and trailing colon/space,
    lower-cased, for denylist comparison."""
    t = text or ""
    if num_raw and t.startswith(num_raw):
        t = t[len(num_raw):]
    return t.strip().strip("：: 　").lower()


def is_unnumbered_section(text, num_raw):
    """True if this heading is a conventionally-unnumbered section
    (摘要/前言/结论/目录/参考文献/附录/致谢/...)."""
    t = _norm_title(text, num_raw)
    if not t:
        return False
    if t in UNNUMBERED_HEADING_EXACT:
        return True
    return t.startswith(UNNUMBERED_HEADING_PREFIX)


def cover_role(rec, spec):
    """Formatting role of a cover paragraph: 'title' | 'classification' |
    'field' | 'other'.

    An explicit rec['cover_role'] (set by the step-2.5 AI review via
    27_apply_review.py) always wins -- that is how an arbitrary cover layout, or
    a 密级/文本编号 line with NO key (just "机密  202501323023"), is classified
    when the heuristics below cannot. Otherwise:
      * title          -- extraction's is_title block, or a per-paragraph
                          fallback (centered & large, or exactly the title size);
      * classification -- text contains a 密级/文本编号 keyword (spec-configured,
                          incl. common 密级 values so a key-less line still hits);
      * field          -- every other non-blank cover line (项目名称/承担单位/…).
    Shared by check_paragraph() and 26_export_review.py so the guess shown to the
    model and the one used to format never drift apart."""
    forced = rec.get("cover_role")
    if forced:
        return forced
    eff = rec.get("eff", {})
    if rec.get("is_title") \
            or (eff.get("jc") == "center" and (eff.get("size_hp") or 0) >= 36) \
            or (eff.get("size_hp") == spec.get("title", {}).get("size_hp")):
        return "title"
    text = rec.get("text") or ""
    kws = (spec.get("cover") or {}).get("classification_keywords") or []
    # classification if it carries a 密级/文本编号 keyword OR it sits above the
    # title (the template's fixed position for that line — catches the key-less
    # "机密  202501323023" case).
    if rec.get("above_title") or any(k in text for k in kws):
        return "classification"
    return "field"


def _size_name(hp):
    """Half-point value -> a human-readable Chinese size label for comments."""
    names = {40: "20磅", 36: "小一", 32: "三号",
             28: "14磅", 24: "小四", 21: "五号"}
    if hp in names:
        return names[hp]
    return "%g磅" % (hp / 2.0)


def _font_ok(actual, spec_entry):
    """True if actual font matches any accepted spelling."""
    if actual is None:
        return False
    for m in spec_entry.get("east_asia_match", [spec_entry["east_asia"]]):
        if m and m in actual:
            return True
    return actual == spec_entry["east_asia"]


# ---------------------------------------------------------------------------
# Per-paragraph format check (region-aware)
# ---------------------------------------------------------------------------
def _new_sets():
    """Fresh accumulator of every settable property a format fix may carry, all
    defaulted to 'no change'. One factory keeps every fix shape uniform, so the
    apply stage and the summary never trip over a key that only some check
    branches happened to include."""
    return {"set_east_asia": None, "set_ascii": None, "set_size_hp": None,
            "set_line_exact": None, "set_first_line_chars": None,
            "set_left_chars": None,
            "clear_left_indent": False, "clear_right_indent": False,
            "set_jc": None, "strip_text": None}


def check_paragraph(rec, spec):
    """Return a combined 'format' fix for this paragraph, or None if compliant/skip."""
    if rec.get("is_blank"):
        return None  # an empty line has no format to judge

    region = rec.get("region", "body")
    if region == "toc":
        # 目录条目只做字体/字号校验（仿宋 三号）；不动缩进/行距，避免破坏 TOC
        # 域代码自身的制表位/悬挂缩进结构。
        return _check_toc(rec, spec)

    # Caption paragraphs (图.../表...): center them and remove all indent
    # (spec.caption_format). Font/size are left alone (the spec states no
    # caption font rule). Numbering is handled separately by continuity().
    if rec.get("caption"):
        return _check_caption_format(rec, spec)

    # Table-cell content has its OWN font/size rule (仿宋 14磅), distinct from
    # ordinary body text (仿宋 三号). Without this a cell paragraph would be
    # judged against the body rule and, worse, have a 2-char first-line indent
    # forced onto it — neither of which is right for tabular content. Only the
    # font/size are enforced here (the spec states nothing about a cell's indent
    # or line spacing, so those are left untouched).
    if rec.get("in_table"):
        return _check_table_body(rec, spec)

    eff = rec["eff"]
    sets = _new_sets()
    violations = []

    # -- Cover region: title / classification (密级·文本编号) / other fields --
    if region == "cover":
        role = cover_role(rec, spec)
        if role == "other":
            return None  # AI explicitly exempted this line; leave it untouched
        if role == "title":
            t = spec["title"]
            if not _font_ok(eff.get("east_asia"), t):
                sets["set_east_asia"] = t["east_asia"]
                violations.append("题目字体应为%s（实际：%s）"
                                  % (t["east_asia"], eff.get("east_asia")))
            if eff.get("size_hp") != t["size_hp"]:
                sets["set_size_hp"] = t["size_hp"]
                violations.append("题目字号应为20磅")
            if eff.get("jc") != "center":
                sets["set_jc"] = "center"
                violations.append("题目应居中")
            # 题目应无缩进；残留的首行/左/右缩进都会让 jc=center 的居中基准偏移
            # （center 是相对左右缩进之后的可用宽度居中，不是相对整个页面宽度），
            # 只清左缩进、留着右缩进一样会导致视觉上不是真正居中。
            _check_no_indent(eff, sets, violations, "题目")
            # 题目首尾的多余空格会参与居中计算，导致视觉上偏移（不是真正居中）。
            # 删除首尾空白（保留标题内部空格），才能真正按正文内容居中。
            text = rec.get("text") or ""
            if text != text.strip():
                sets["strip_text"] = "both"
                violations.append("题目首尾多余空格应删除（否则居中会偏移）")
            return _mk_format(rec["i"], sets, violations)
        # classification (密级/文本编号): font/size only, per spec.
        # Alignment/indent are NOT enforced -- the spec states nothing about them
        # and the 密级/编号 line's position is template layout.
        if role == "classification":
            entry = spec.get("cover_classification")
            label = "封面密级/文本编号"
            if not entry or not entry.get("east_asia") or not entry.get("size_hp"):
                return None  # spec not configured; nothing to check
            _check_font_size(eff, entry, sets, violations, label=label)
            return _mk_format(rec["i"], sets, violations)
        # field role (项目名称/承担单位/项目负责人/起止时间/编制时间 等):
        # font/size per spec, PLUS these lines must be left-aligned. When they
        # carry leading spaces or a left/first-line indent (template layout),
        # that pushes the text right so it no longer sits flush at the margin --
        # strip the leading whitespace, clear the indent, and force left
        # alignment. Internal spaces (fill-in blanks like "20   年   月") are
        # preserved -- only the leading padding is removed.
        entry = spec.get("cover_field")
        label = "封面要素"
        if entry and entry.get("east_asia") and entry.get("size_hp"):
            _check_font_size(eff, entry, sets, violations, label=label)
        _check_cover_field_left(rec, eff, sets, violations)
        return _mk_format(rec["i"], sets, violations)

    # -- Body region: headings / captions / normal body --
    # An UNCONFIRMED ("pattern") heading is treated as body text here, not
    # as a heading: the script isn't confident it's really a heading at all
    # (see continuity()'s docstring), so forcing heading-only fonts
    # (黑体/楷体) onto what might just be ordinary body text would be its own
    # false-positive risk. It still gets the (safe either way) body checks,
    # and continuity() separately raises a heading.unconfirmed review hint.
    # Once the model confirms it (level_source becomes "model_confirmed"),
    # it gets real heading-font treatment on the next run.
    western = spec.get("western_font")
    has_western = rec.get("has_western", False)
    if rec.get("is_heading") and rec.get("level_source") != "pattern":
        lvl = rec.get("level") or 1
        lvl = min(max(lvl, 1), 4)
        h = spec["headings"][str(lvl)]
        _check_font_size(eff, h, sets, violations,
                         label="%d级标题" % lvl, western=western, has_western=has_western)
        _check_first_line(eff, h, sets, violations, auto_num=rec.get("auto_num"))
        # headings are not forced to a specific line rule by the spec table
    else:
        b = spec["body"]
        _check_font_size(eff, b, sets, violations, label="正文",
                         western=western, has_western=has_western)
        # line spacing: fixed value 28pt (exact 560)
        ls = spec["line_spacing"]
        if not (eff.get("line") == ls["line_twips"] and eff.get("line_rule") == ls["line_rule"]):
            sets["set_line_exact"] = ls["line_twips"]
            violations.append("行距应为固定值28磅")
        _check_first_line(eff, b, sets, violations, auto_num=rec.get("auto_num"))

    return _mk_format(rec["i"], sets, violations)


def _check_table_body(rec, spec):
    """表格内容：仿宋 14磅。Only font + size are checked (the spec says nothing
    about a cell's indent/line spacing, so we do not invent those)."""
    tb = spec.get("table_body")
    if not tb or not tb.get("east_asia") or not tb.get("size_hp"):
        return None  # spec not configured for table content; nothing to check
    eff = rec["eff"]
    sets = _new_sets()
    violations = []
    _check_font_size(eff, tb, sets, violations, label="表格内容",
                     western=spec.get("western_font"),
                     has_western=rec.get("has_western", False))
    return _mk_format(rec["i"], sets, violations)


def _check_caption_format(rec, spec):
    """图表标题：居中、无任何缩进（首行/左/右全部清零）。字体字号规范未定义，不动。"""
    cf = spec.get("caption_format")
    if not cf:
        return None
    eff = rec["eff"]
    sets = _new_sets()
    violations = []
    want_jc = cf.get("jc")
    if want_jc and eff.get("jc") != want_jc:
        sets["set_jc"] = want_jc
        violations.append("图表标题应居中")
    if cf.get("no_indent"):
        _check_no_indent(eff, sets, violations, "图表标题")
    return _mk_format(rec["i"], sets, violations)


def _check_toc(rec, spec):
    t = spec.get("toc") or {}
    if not t.get("east_asia") or not t.get("size_hp"):
        return None  # spec not configured for toc fonts; nothing to check
    eff = rec["eff"]
    sets = _new_sets()
    violations = []
    _check_font_size(eff, t, sets, violations, label="目录")
    # 目录按级缩进：一级 0 / 二级 2 字符 / 三级 4 字符（leftChars 0/200/400）。
    # 只动 w:ind（首行/左/右），不碰 w:tabs（TOC 域自身用来对齐点线号+页码的
    # 制表位，与 w:ind 是两套独立机制，清 w:ind 不会破坏它）。
    by_level = t.get("indent_chars_by_level") or {}
    lvl = rec.get("toc_level")
    want_left = by_level.get(str(lvl)) if (lvl is not None) else None
    if want_left is None:
        want_left = 0  # 级别未知时按不缩进的安全默认
    # 首行/悬挂缩进一律清零
    if any(eff.get(k) for k in ("first_line_chars", "first_line", "hanging_chars", "hanging")):
        sets["set_first_line_chars"] = 0
        violations.append("目录不应有首行缩进")
    # 左缩进设为该级目标值（字符数）
    cur_left_chars = eff.get("left_chars") or 0
    has_abs_left = bool(eff.get("left") or eff.get("start") or eff.get("start_chars"))
    if cur_left_chars != want_left or has_abs_left:
        sets["set_left_chars"] = want_left
        if want_left:
            violations.append("目录%s级应缩进%d字符" % (lvl, want_left // 100))
        else:
            violations.append("目录一级不应缩进")
    # 右缩进清零
    if any(eff.get(k) for k in ("right_chars", "right", "end_chars", "end")):
        sets["clear_right_indent"] = True
        violations.append("目录不应有右缩进")
    return _mk_format(rec["i"], sets, violations)


def _check_cover_field_left(rec, eff, sets, violations):
    """封面要素行（项目名称/承担单位/项目负责人/起止时间/编制时间 等）版式修复。

    规范：两端对齐、首行缩进2字符；并保留原有的"清除左侧空格 + 清除左侧字符
    缩进"逻辑——行首空白与左缩进会把文本顶离左边距，需清掉。行内填空空格
    （如"20   年   月"）保留，只清行首。"""
    text = rec.get("text") or ""
    # 1) 对齐方式：改为两端对齐
    if eff.get("jc") != "both":
        sets["set_jc"] = "both"
        violations.append("封面要素应两端对齐")
    # 2) 首行缩进2字符
    if eff.get("first_line_chars") != 200:
        sets["set_first_line_chars"] = 200
        violations.append("封面要素首行缩进应为2字符")
    # 3) 左侧字符缩进：清零（保留首行缩进，不与其叠加）
    if any(eff.get(k) for k in ("left_chars", "left", "start_chars", "start")):
        sets["clear_left_indent"] = True
        violations.append("封面要素应清除左侧缩进")
    # 4) 行首空格：删除（否则文本仍被空格顶开）
    if text != text.lstrip():
        # 既有清首尾需求时用 both，否则仅清行首
        sets["strip_text"] = "both" if sets.get("strip_text") == "both" else "leading"
        violations.append("封面要素行首空格应删除")


def _check_no_indent(eff, sets, violations, label):
    """Flag+clear ANY indentation (first-line/left/right, in both the
    twips and *Chars forms) on a paragraph that per spec should have none
    at all (title, TOC entries)."""
    indent_keys = ("first_line_chars", "first_line", "left_chars", "left",
                   "start_chars", "start", "right_chars", "right",
                   "end_chars", "end", "hanging_chars", "hanging")
    if any(eff.get(k) for k in indent_keys):
        sets["set_first_line_chars"] = 0
        sets["clear_left_indent"] = True
        sets["clear_right_indent"] = True
        violations.append("%s不应有缩进" % label)


def _check_font_size(eff, spec_entry, sets, violations, label,
                     western=None, has_western=False):
    """Check east-asian font + size against spec_entry. When `western` is given
    (the spec's western_font, for running text — 正文/标题/表格), the paragraph's
    EFFECTIVE Latin/digit font must equal it too, so 西文 renders as the required
    Times New Roman — but ONLY when the paragraph actually CONTAINS Western text
    (has_western). A pure-Chinese line has no Western characters to normalize, so
    it is never flagged for its Latin font (which is often set to the same 仿宋 as
    the Chinese font and would otherwise raise a confusing '西文应为 TNR' note on
    every Chinese paragraph). Only enforced where a western font was passed —
    cover/目录 layout lines deliberately opt out."""
    if not _font_ok(eff.get("east_asia"), spec_entry):
        sets["set_east_asia"] = spec_entry["east_asia"]
        violations.append("%s中文字体应为%s（实际：%s）"
                          % (label, spec_entry["east_asia"], eff.get("east_asia")))
    if western and has_western and eff.get("ascii") != western:
        sets["set_ascii"] = western
        violations.append("%s西文字体应为%s（实际：%s）"
                          % (label, western, eff.get("ascii")))
    if eff.get("size_hp") != spec_entry["size_hp"]:
        sets["set_size_hp"] = spec_entry["size_hp"]
        violations.append("%s字号应为%s" % (label, _size_name(spec_entry["size_hp"])))


def _check_first_line(eff, spec_entry, sets, violations, auto_num=False):
    want = spec_entry.get("first_line_chars")
    if want is None:
        return
    if auto_num:
        # 自动编号段落：编号会带来左/悬挂缩进，而且该级缩进不一定完整写在
        # XML 里（尤其二级及以下常落在 Word 内置列表缩进上，读不到、也就检测
        # 不到），所以不依赖"检测到才清"——一律强制写成"首行缩进2字符、左右0"，
        # 直接缩进会覆盖编号带来的缩进。这样各级标题清缩进才一致，不会只有一级
        # 生效。
        sets["set_first_line_chars"] = want
        sets["clear_left_indent"] = True
        sets["clear_right_indent"] = True
        violations.append("首行缩进应为2字符，并清除左右缩进（含自动编号带来的缩进）")
        return
    if eff.get("first_line_chars") != want:
        sets["set_first_line_chars"] = want
        violations.append("首行缩进应为2字符")
    # left/right indent must NOT be stacked on top of the first-line indent:
    # an existing left indent makes the first line indent by (left + 2 chars),
    # i.e. more than the required 2 characters, so any left/right indent is
    # zeroed and only the 2-char first-line indent is kept.
    if any(eff.get(k) for k in ("left_chars", "left", "start_chars", "start")):
        sets["clear_left_indent"] = True
        if "首行缩进应为2字符" not in violations:
            violations.append("应仅保留首行缩进，需清除左缩进")
        else:
            violations.append("同时需清除左缩进（不应与首行缩进叠加）")
    if any(eff.get(k) for k in ("right_chars", "right", "end_chars", "end")):
        sets["clear_right_indent"] = True
        violations.append("需清除右缩进")


def _mk_format(i, sets, violations):
    if not violations:
        return None
    fix = {"para_index": i, "op": "format",
           "rule_id": "combined",
           "rule_text": "；".join(violations),
           "comment": True}
    fix.update(sets)
    return fix


# ---------------------------------------------------------------------------
# Page setup
# ---------------------------------------------------------------------------
def check_page_setup(page_setup, spec):
    if not page_setup:
        return None
    ps = spec["page_setup"]
    tol = ps["tolerance_twips"]
    want = {"top": ps["margin_top_twips"], "bottom": ps["margin_bottom_twips"],
            "left": ps["margin_left_twips"], "right": ps["margin_right_twips"],
            "header": ps["header_twips"], "footer": ps["footer_twips"]}
    setmar, viol = {}, []
    names = {"top": "上", "bottom": "下", "left": "左", "right": "右",
             "header": "页眉", "footer": "页脚"}
    for k, w in want.items():
        cur = page_setup.get(k)
        if cur is None or abs(cur - w) > tol:
            setmar[k] = w
            viol.append("%s边距" % names[k])
    if not setmar:
        return None
    return {"op": "section", "set_pgmar": setmar,
            "rule_id": "page.margins",
            "rule_text": "页面边距不符合规范（应：上3.5/下3.2/左右2.85cm，页眉1.5/页脚1.75cm）：" + "、".join(viol),
            "comment": False}


# ---------------------------------------------------------------------------
# Continuity (headings + captions) -> renumber fixes
#
# Both loops below split records into two trust tiers before doing anything:
#   - CONFIRMED (level_source/caption_source is "outline"/"style", or
#     "model_confirmed" after the model explicitly reviewed and approved it
#     via 27_apply_review.py): the script is confident this really is a
#     heading/caption, so it is folded into the numbering sequence and
#     auto-renumbered on mismatch, no further confirmation needed.
#   - UNCONFIRMED ("pattern": detected purely because the text happens to
#     match a numbering shape, with no style/outline backing it): the script
#     is NOT confident this is really a heading/caption at all -- plenty of
#     ordinary text can accidentally start with "3." or "图1" -- so it is
#     never auto-edited and never folded into the counting sequence (doing
#     either could corrupt the numbering of paragraphs that ARE genuinely
#     confirmed). It only ever gets a review hint. Once a human-equivalent
#     review (the model, via overrides.json) confirms it, it graduates to
#     "model_confirmed" and is treated as confirmed on the next run.
# ---------------------------------------------------------------------------
def continuity(records, spec):
    fixes = []
    # -- headings --
    counters = {1: 0, 2: 0, 3: 0, 4: 0}
    for r in records:
        if r.get("region") != "body" or not r.get("is_heading"):
            continue
        lvl = r.get("level")
        if not lvl or lvl < 1 or lvl > 4:
            continue
        if r.get("level_source") == "pattern":
            fixes.append({
                "para_index": r["i"], "op": "hint",
                "rule_id": "heading.unconfirmed",
                "rule_text": "疑似%d级标题（仅按序号格式识别，未见标题样式/大纲级别），请人工确认是否为标题及层级" % lvl,
                "comment": True,
            })
            continue
        if is_unnumbered_section(r.get("text"), r.get("num_raw")):
            # 惯例不编号的章节（摘要/前言/结论/目录/参考文献/附录/致谢…）：即便挂了
            # 标题/大纲样式也不自动补号，且不占用同级序号计数——否则它后面真正要
            # 编号的同级标题会被多算一位。
            continue
        counters[lvl] += 1
        for d in range(lvl + 1, 5):
            counters[d] = 0
        expected = counters[lvl]
        expected_token = _heading_token(lvl, expected)
        raw = r.get("num_raw")
        if raw is None:
            # No ordinal in the TEXT.
            if r.get("auto_num"):
                # Pure automatic numbering (Word 多级列表): the number is
                # generated by Word, not in the text, and stays continuous on
                # its own -- leave it to Word (方案1). The counter was already
                # advanced above so any manually-numbered siblings still line up.
                continue
            # 已确认样式/大纲级别、且不在"惯例不编号"名单里的标题——漏了序号就按
            # 位置自动补（与图表标题补号一致），不再只提示让用户确认。惯例不编号的
            # 章节已在上面提前跳过，不会走到这里。
            fixes.append({
                "para_index": r["i"], "op": "renumber_heading", "level": lvl,
                "new_token": expected_token, "insert": True,
                "rule_id": "heading.missing_number",
                "rule_text": "%d级标题缺少序号，已按位置自动补为“%s”" % (lvl, expected_token),
                "comment": True,
            })
            continue
        # raw is not None: there IS a selectable ordinal typed in the text. It is
        # authoritative and gets normalized to the position-correct token EVEN IF
        # the style also carries an automatic-numbering definition -- a typed,
        # selectable number is "manual" per the user's rule ("自动编号交给 Word，
        # 其余的要识别并修改"), so a level-1 heading typed "1    项目" becomes
        # "一、项目".
        if raw != expected_token:
            # Covers BOTH kinds of violation with one comparison: sequence gaps
            # (e.g. "一、" skipping to "三、") AND format errors (e.g. arabic
            # "3、" used where the level-1 spec requires the Chinese numeral
            # "三、") -- either way the rendered label differs from the
            # position-derived canonical token, so it gets corrected.
            fixes.append({
                "para_index": r["i"], "op": "renumber_heading", "level": lvl,
                "new_token": expected_token,
                "rule_id": "heading.continuity",
                "rule_text": "%d级标题序号格式或连续性有误，应为“%s”（原为“%s”）" % (lvl, expected_token, raw.rstrip()),
                "comment": True,
            })
            # keep counters at expected (we renumbered to expected)
    # -- captions --
    # decide flat vs chapter-based per kind from the raw numbers present
    for kind in ("figure", "table"):
        kname = "图" if kind == "figure" else "表"
        all_of_kind = [r for r in records if r.get("region") == "body" and r.get("caption")
                       and r["caption"]["kind"] == kind]
        if not all_of_kind:
            continue
        confirmed = []
        for r in all_of_kind:
            cap = r["caption"]
            if cap.get("source") == "pattern":
                fixes.append({
                    "para_index": r["i"], "op": "hint",
                    "rule_id": "caption.%s.unconfirmed" % kind,
                    "rule_text": "疑似%s标题（仅按“%s+数字”格式识别，未见题注样式），请人工确认是否为标题" % (kname, kname),
                    "comment": True,
                })
                continue
            # Confirmed caption (题注/图标题/表标题/... style, or model-confirmed).
            # ALL confirmed captions join the numbering sequence and get a
            # deterministic static 图N/表N -- whether the original number was
            # typed, missing (num_raw is None -> AUTO-INSERTED below), or produced
            # by Word automatic numbering (auto_num). For auto_num captions the
            # apply stage additionally cancels the paragraph's w:numPr (numId=0)
            # so Word's own number does not double up on top of the static one;
            # this is what lets an ABNORMAL/broken auto-number be corrected to a
            # proper 图N/表N instead of merely un-hidden. A caption style is a
            # reliable enough signal to do all this without further confirmation.
            confirmed.append(r)
        if not confirmed:
            continue
        # chapter-based scheme is inferred only from captions that actually
        # carry a separator-bearing number ("图1-1"); an un-numbered confirmed
        # caption (num_raw is None) does not vote here.
        chapter_based = any(r["caption"]["num_raw"] and
                            any(s in r["caption"]["num_raw"] for s in ("-", ".", "–"))
                            for r in confirmed)
        if chapter_based:
            group_counter = {}
            for r in confirmed:
                raw = r["caption"]["num_raw"]
                if raw is None:
                    # Cannot infer which chapter an un-numbered caption belongs
                    # to (its "图1-?" prefix is unknown), so hint rather than
                    # fabricate a chapter prefix.
                    fixes.append({
                        "para_index": r["i"], "op": "hint",
                        "rule_id": "caption.%s.missing_number" % kind,
                        "rule_text": "%s标题样式已确认，但未识别到编号，请确认是否漏编号" % kname,
                        "comment": True,
                    })
                    continue
                sep = next((s for s in ("-", ".", "–") if s in raw), "-")
                prefix = raw.rsplit(sep, 1)[0]
                group_counter[prefix] = group_counter.get(prefix, 0) + 1
                expected = "%s%s%d" % (prefix, sep, group_counter[prefix])
                if raw != expected:
                    fixes.append(_caption_fix(r, kind, expected, raw))
        else:
            n = 0
            for r in confirmed:
                n += 1
                expected = str(n)
                raw = r["caption"]["num_raw"]
                if raw is None:
                    # Confirmed caption with no number at all -> auto-insert
                    # "图N"/"表N" per the sequence position.
                    fixes.append(_caption_fix(r, kind, expected, None, insert=True))
                elif raw != expected:
                    fixes.append(_caption_fix(r, kind, expected, raw))
    return fixes


def _caption_fix(r, kind, new_num, old_num, insert=False):
    kname = "图" if kind == "figure" else "表"
    # auto_num captions carry Word automatic numbering (w:numPr) that the apply
    # stage must cancel before writing the static number, or the two would stack.
    strip_auto = bool(r.get("auto_num"))
    if strip_auto:
        rule_text = "%s标题原为自动编号（可能异常），已改为按顺序静态编号“%s%s”" % (kname, kname, new_num)
    elif insert:
        rule_text = "%s标题缺少编号，已按顺序自动添加为“%s%s”" % (kname, kname, new_num)
    else:
        rule_text = "%s编号不连续，应为“%s%s”（原为“%s%s”）" % (kname, kname, new_num, kname, old_num)
    return {
        "para_index": r["i"], "op": "renumber_caption", "kind": kind,
        "new_num": new_num, "insert": insert, "strip_auto": strip_auto,
        "rule_id": "caption.%s.continuity" % kind,
        "rule_text": rule_text,
        "comment": True,
    }


def _heading_token(level, n):
    if level == 1:
        return int2cn(n) + "、"
    if level == 2:
        return "（" + int2cn(n) + "）"
    if level == 3:
        return "%d." % n
    if level == 4:
        return "（%d）" % n
    return str(n)


# ---------------------------------------------------------------------------
# Document-level hints (comment only): cover content, green bg, caption content,
# TOC depth, page-number font (best-effort).
# ---------------------------------------------------------------------------
def _cover_missing_fields(structure, spec):
    """Required cover items that appear to be MISSING.

    If the step-2.5 AI review supplied structure['cover_present'] (the items it
    semantically confirmed are actually filled in — the reliable source for
    title-derived items like 项目名称/考核年份, or a still-placeholder title),
    that is authoritative. Otherwise fall back to a best-effort deterministic
    baseline: title-derived items come from the detected title block / a real
    year, the rest from a label substring match on the cover text."""
    records = structure["records"]
    cov = spec.get("cover", {})
    required = cov.get("required_fields", [])
    ai_present = structure.get("cover_present")
    if isinstance(ai_present, list):
        present = set(ai_present)
        return [f for f in required if f not in present]
    cover_recs = [r for r in records if r.get("region") == "cover" and not r.get("is_blank")]
    cover_text = "".join(r.get("text") or "" for r in cover_recs)
    title_text = "".join(r.get("text") or "" for r in cover_recs if r.get("is_title"))
    title_derived = set(cov.get("title_derived_fields", []))
    has_title = bool(title_text.strip())
    has_year = bool(re.search(r"(?:19|20)\d{2}", cover_text)) or ("年度" in title_text)
    secrecy_values = cov.get("secrecy_values", [])
    present = set()
    for f in required:
        if f == "项目名称" and f in title_derived:
            if has_title:
                present.add(f)
        elif f == "考核年份" and f in title_derived:
            if has_year:
                present.add(f)
        elif f == "密级":
            # 密级 is often written as its VALUE (机密/秘密/…) with no "密级"
            # label, so accept either.
            if "密级" in cover_text or any(v in cover_text for v in secrecy_values):
                present.add(f)
        elif f in cover_text:
            present.add(f)
    return [f for f in required if f not in present]


def _cover_placeholder(structure, spec):
    """True if the title still contains an un-replaced template placeholder
    (XXXX / ×××× / 模板)."""
    markers = (spec.get("cover") or {}).get("placeholder_markers") or []
    title_text = "".join(r.get("text") or "" for r in structure["records"]
                         if r.get("region") == "cover" and r.get("is_title"))
    return any(m and m in title_text for m in markers)


def doc_hints(structure, spec):
    records = structure["records"]
    hints = []
    # cover content presence (AI-verified when available, else best-effort)
    missing = _cover_missing_fields(structure, spec)
    # attach cover hints to the first cover paragraph
    first_cover = next((r["i"] for r in records if r.get("region") == "cover"), None)
    if first_cover is not None:
        if missing:
            hints.append({"para_index": first_cover, "op": "hint",
                          "rule_id": "cover.fields",
                          "rule_text": "封皮缺少必要信息：" + "、".join(missing),
                          "comment": True})
        if _cover_placeholder(structure, spec):
            hints.append({"para_index": first_cover, "op": "hint",
                          "rule_id": "cover.placeholder",
                          "rule_text": "封面标题疑似仍为模板占位符（如 XXXX/模板），请替换为实际项目名称与考核年份",
                          "comment": True})
        if spec["cover"].get("background_green") and not structure.get("has_background"):
            hints.append({"para_index": first_cover, "op": "hint",
                          "rule_id": "cover.green",
                          "rule_text": "未检测到封皮背景颜色设置，请确认封皮为绿色",
                          "comment": True})
    # caption content presence
    for r in records:
        cap = r.get("caption")
        if cap and not cap.get("has_content"):
            kname = "图" if cap["kind"] == "figure" else "表"
            hints.append({"para_index": r["i"], "op": "hint",
                          "rule_id": "caption.content",
                          "rule_text": "%s编号后缺少内容说明" % kname,
                          "comment": True})

    # TOC depth: the spec caps the table of contents at max_level (三级). If any
    # TOC entry style is deeper, hint (removing entries is a content edit, not a
    # format one — so hint-only, never auto-changed).
    toc = spec.get("toc") or {}
    max_level = toc.get("max_level")
    if max_level:
        first_toc = next((r["i"] for r in records if r.get("region") == "toc"), None)
        deepest = max((r.get("toc_level") or 0)
                      for r in records if r.get("region") == "toc") if any(
                          r.get("region") == "toc" for r in records) else 0
        if first_toc is not None and deepest > max_level:
            hints.append({"para_index": first_toc, "op": "hint",
                          "rule_id": "toc.depth",
                          "rule_text": "目录层级超过%d级（检测到%d级），规范要求目录不超过%d级标题，请人工精简"
                                       % (max_level, deepest, max_level),
                          "comment": True})

    # Page-number font (best-effort, hint only): the page number lives in a
    # header/footer PAGE field, not the body, so 20_extract_structure.py records
    # its effective western font/size into structure['page_number'] when found.
    # Compare against spec; only hint on a concrete mismatch (never on an
    # unread value, to avoid false positives).
    pn_spec = spec.get("page_number") or {}
    pn = structure.get("page_number")
    anchor = first_cover if first_cover is not None else (
        records[0]["i"] if records else None)
    if pn_spec and pn and anchor is not None:
        bad = []
        want_ascii = pn_spec.get("ascii")
        want_sz = pn_spec.get("size_hp")
        if want_ascii and pn.get("ascii") and pn["ascii"] != want_ascii:
            bad.append("字体应为%s（实际：%s）" % (want_ascii, pn["ascii"]))
        if want_sz and pn.get("size_hp") and pn["size_hp"] != want_sz:
            bad.append("字号应为%s（实际：%g磅）" % (_size_name(want_sz), pn["size_hp"] / 2.0))
        if bad:
            hints.append({"para_index": anchor, "op": "hint",
                          "rule_id": "page_number.font",
                          "rule_text": "页码" + "、".join(bad) + "，请人工核实页眉/页脚页码格式",
                          "comment": True})
    return hints
