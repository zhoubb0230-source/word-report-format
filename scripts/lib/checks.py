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

CN_DIGITS = "\u96f6\u4e00\u4e8c\u4e09\u56db\u4e94\u516d\u4e03\u516b\u4e5d"


def load_spec(spec_path):
    with open(spec_path, encoding="utf-8") as f:
        return json.load(f)


def int2cn(n):
    """1..99 -> Chinese numeral (一, 十, 十一, 二十, 二十三 ...)."""
    if n <= 0:
        return str(n)
    if n < 10:
        return CN_DIGITS[n]
    if n == 10:
        return "\u5341"
    if n < 20:
        return "\u5341" + CN_DIGITS[n - 10]
    tens, ones = divmod(n, 10)
    s = CN_DIGITS[tens] + "\u5341"
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
    sets = {"set_east_asia": None, "set_ascii": None, "set_size_hp": None,
            "set_line_exact": None, "set_first_line_chars": None,
            "clear_left_indent": False, "clear_right_indent": False, "set_jc": None}
    violations = []

    # -- Cover region: title / classification (\u5bc6\u7ea7\u00b7\u6587\u672c\u7f16\u53f7) / other fields --
    if region == "cover":
        role = cover_role(rec, spec)
        if role == "other":
            return None  # AI explicitly exempted this line; leave it untouched
        if role == "title":
            t = spec["title"]
            if not _font_ok(eff.get("east_asia"), t):
                sets["set_east_asia"] = t["east_asia"]
                violations.append("\u9898\u76ee\u5b57\u4f53\u5e94\u4e3a%s\uff08\u5b9e\u9645\uff1a%s\uff09"
                                  % (t["east_asia"], eff.get("east_asia")))
            if eff.get("size_hp") != t["size_hp"]:
                sets["set_size_hp"] = t["size_hp"]
                violations.append("\u9898\u76ee\u5b57\u53f7\u5e94\u4e3a20\u78c5")
            if eff.get("jc") != "center":
                sets["set_jc"] = "center"
                violations.append("\u9898\u76ee\u5e94\u5c45\u4e2d")
            # \u9898\u76ee\u5e94\u65e0\u7f29\u8fdb\uff1b\u6b8b\u7559\u7684\u9996\u884c/\u5de6/\u53f3\u7f29\u8fdb\u90fd\u4f1a\u8ba9 jc=center \u7684\u5c45\u4e2d\u57fa\u51c6\u504f\u79fb
            # \uff08center \u662f\u76f8\u5bf9\u5de6\u53f3\u7f29\u8fdb\u4e4b\u540e\u7684\u53ef\u7528\u5bbd\u5ea6\u5c45\u4e2d\uff0c\u4e0d\u662f\u76f8\u5bf9\u6574\u4e2a\u9875\u9762\u5bbd\u5ea6\uff09\uff0c
            # \u53ea\u6e05\u5de6\u7f29\u8fdb\u3001\u7559\u7740\u53f3\u7f29\u8fdb\u4e00\u6837\u4f1a\u5bfc\u81f4\u89c6\u89c9\u4e0a\u4e0d\u662f\u771f\u6b63\u5c45\u4e2d\u3002
            _check_no_indent(eff, sets, violations, "\u9898\u76ee")
            return _mk_format(rec["i"], sets, violations)
        # classification (\u5bc6\u7ea7/\u6587\u672c\u7f16\u53f7) and field roles: font/size only, per spec.
        # Alignment/indent are NOT enforced -- the spec states nothing about them
        # for these cover lines, and their leading spaces are template layout.
        if role == "classification":
            entry = spec.get("cover_classification")
            label = "\u5c01\u9762\u5bc6\u7ea7/\u6587\u672c\u7f16\u53f7"
        else:  # "field"
            entry = spec.get("cover_field")
            label = "\u5c01\u9762\u8981\u7d20"
        if not entry or not entry.get("east_asia") or not entry.get("size_hp"):
            return None  # spec not configured for this cover role; nothing to check
        _check_font_size(eff, entry, sets, violations, label=label)
        return _mk_format(rec["i"], sets, violations)

    # -- Body region: headings / captions / normal body --
    # An UNCONFIRMED ("pattern") heading is treated as body text here, not
    # as a heading: the script isn't confident it's really a heading at all
    # (see continuity()'s docstring), so forcing heading-only fonts
    # (\u9ed1\u4f53/\u6977\u4f53) onto what might just be ordinary body text would be its own
    # false-positive risk. It still gets the (safe either way) body checks,
    # and continuity() separately raises a heading.unconfirmed review hint.
    # Once the model confirms it (level_source becomes "model_confirmed"),
    # it gets real heading-font treatment on the next run.
    if rec.get("is_heading") and rec.get("level_source") != "pattern":
        lvl = rec.get("level") or 1
        lvl = min(max(lvl, 1), 4)
        h = spec["headings"][str(lvl)]
        _check_font_size(eff, h, sets, violations,
                         label="%d\u7ea7\u6807\u9898" % lvl)
        _check_first_line(eff, h, sets, violations, auto_num=rec.get("auto_num"))
        # headings are not forced to a specific line rule by the spec table
    else:
        b = spec["body"]
        _check_font_size(eff, b, sets, violations, label="\u6b63\u6587")
        # line spacing: fixed value 28pt (exact 560)
        ls = spec["line_spacing"]
        if not (eff.get("line") == ls["line_twips"] and eff.get("line_rule") == ls["line_rule"]):
            sets["set_line_exact"] = ls["line_twips"]
            violations.append("\u884c\u8ddd\u5e94\u4e3a\u56fa\u5b9a\u503c28\u78c5")
        _check_first_line(eff, b, sets, violations, auto_num=rec.get("auto_num"))

    return _mk_format(rec["i"], sets, violations)


def _check_table_body(rec, spec):
    """表格内容：仿宋 14磅。Only font + size are checked (the spec says nothing
    about a cell's indent/line spacing, so we do not invent those)."""
    tb = spec.get("table_body")
    if not tb or not tb.get("east_asia") or not tb.get("size_hp"):
        return None  # spec not configured for table content; nothing to check
    eff = rec["eff"]
    sets = {"set_east_asia": None, "set_ascii": None, "set_size_hp": None,
            "set_line_exact": None, "set_first_line_chars": None,
            "clear_left_indent": False, "clear_right_indent": False, "set_jc": None}
    violations = []
    _check_font_size(eff, tb, sets, violations, label="表格内容")
    return _mk_format(rec["i"], sets, violations)


def _check_caption_format(rec, spec):
    """图表标题：居中、无任何缩进（首行/左/右全部清零）。字体字号规范未定义，不动。"""
    cf = spec.get("caption_format")
    if not cf:
        return None
    eff = rec["eff"]
    sets = {"set_east_asia": None, "set_ascii": None, "set_size_hp": None,
            "set_line_exact": None, "set_first_line_chars": None,
            "clear_left_indent": False, "clear_right_indent": False, "set_jc": None}
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
    sets = {"set_east_asia": None, "set_ascii": None, "set_size_hp": None,
            "set_line_exact": None, "set_first_line_chars": None, "set_left_chars": None,
            "clear_left_indent": False, "clear_right_indent": False, "set_jc": None}
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


def _check_font_size(eff, spec_entry, sets, violations, label):
    if not _font_ok(eff.get("east_asia"), spec_entry):
        sets["set_east_asia"] = spec_entry["east_asia"]
        violations.append("%s\u4e2d\u6587\u5b57\u4f53\u5e94\u4e3a%s\uff08\u5b9e\u9645\uff1a%s\uff09"
                          % (label, spec_entry["east_asia"], eff.get("east_asia")))
    if eff.get("size_hp") != spec_entry["size_hp"]:
        sets["set_size_hp"] = spec_entry["size_hp"]
        violations.append("%s\u5b57\u53f7\u5e94\u4e3a%s" % (label, _size_name(spec_entry["size_hp"])))


def _check_first_line(eff, spec_entry, sets, violations, auto_num=False):
    want = spec_entry.get("first_line_chars")
    if want is None:
        return
    if auto_num:
        # \u81ea\u52a8\u7f16\u53f7\u6bb5\u843d\uff1a\u7f16\u53f7\u4f1a\u5e26\u6765\u5de6/\u60ac\u6302\u7f29\u8fdb\uff0c\u800c\u4e14\u8be5\u7ea7\u7f29\u8fdb\u4e0d\u4e00\u5b9a\u5b8c\u6574\u5199\u5728
        # XML \u91cc\uff08\u5c24\u5176\u4e8c\u7ea7\u53ca\u4ee5\u4e0b\u5e38\u843d\u5728 Word \u5185\u7f6e\u5217\u8868\u7f29\u8fdb\u4e0a\uff0c\u8bfb\u4e0d\u5230\u3001\u4e5f\u5c31\u68c0\u6d4b
        # \u4e0d\u5230\uff09\uff0c\u6240\u4ee5\u4e0d\u4f9d\u8d56"\u68c0\u6d4b\u5230\u624d\u6e05"\u2014\u2014\u4e00\u5f8b\u5f3a\u5236\u5199\u6210"\u9996\u884c\u7f29\u8fdb2\u5b57\u7b26\u3001\u5de6\u53f30"\uff0c
        # \u76f4\u63a5\u7f29\u8fdb\u4f1a\u8986\u76d6\u7f16\u53f7\u5e26\u6765\u7684\u7f29\u8fdb\u3002\u8fd9\u6837\u5404\u7ea7\u6807\u9898\u6e05\u7f29\u8fdb\u624d\u4e00\u81f4\uff0c\u4e0d\u4f1a\u53ea\u6709\u4e00\u7ea7
        # \u751f\u6548\u3002
        sets["set_first_line_chars"] = want
        sets["clear_left_indent"] = True
        sets["clear_right_indent"] = True
        violations.append("\u9996\u884c\u7f29\u8fdb\u5e94\u4e3a2\u5b57\u7b26\uff0c\u5e76\u6e05\u9664\u5de6\u53f3\u7f29\u8fdb\uff08\u542b\u81ea\u52a8\u7f16\u53f7\u5e26\u6765\u7684\u7f29\u8fdb\uff09")
        return
    if eff.get("first_line_chars") != want:
        sets["set_first_line_chars"] = want
        violations.append("\u9996\u884c\u7f29\u8fdb\u5e94\u4e3a2\u5b57\u7b26")
    # left/right indent must NOT be stacked on top of the first-line indent:
    # an existing left indent makes the first line indent by (left + 2 chars),
    # i.e. more than the required 2 characters, so any left/right indent is
    # zeroed and only the 2-char first-line indent is kept.
    if any(eff.get(k) for k in ("left_chars", "left", "start_chars", "start")):
        sets["clear_left_indent"] = True
        if "\u9996\u884c\u7f29\u8fdb\u5e94\u4e3a2\u5b57\u7b26" not in violations:
            violations.append("\u5e94\u4ec5\u4fdd\u7559\u9996\u884c\u7f29\u8fdb\uff0c\u9700\u6e05\u9664\u5de6\u7f29\u8fdb")
        else:
            violations.append("\u540c\u65f6\u9700\u6e05\u9664\u5de6\u7f29\u8fdb\uff08\u4e0d\u5e94\u4e0e\u9996\u884c\u7f29\u8fdb\u53e0\u52a0\uff09")
    if any(eff.get(k) for k in ("right_chars", "right", "end_chars", "end")):
        sets["clear_right_indent"] = True
        violations.append("\u9700\u6e05\u9664\u53f3\u7f29\u8fdb")


def _mk_format(i, sets, violations):
    if not violations:
        return None
    fix = {"para_index": i, "op": "format",
           "rule_id": "combined",
           "rule_text": "\uff1b".join(violations),
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
    names = {"top": "\u4e0a", "bottom": "\u4e0b", "left": "\u5de6", "right": "\u53f3",
             "header": "页眉", "footer": "页脚"}
    for k, w in want.items():
        cur = page_setup.get(k)
        if cur is None or abs(cur - w) > tol:
            setmar[k] = w
            viol.append("%s\u8fb9\u8ddd" % names[k])
    if not setmar:
        return None
    return {"op": "section", "set_pgmar": setmar,
            "rule_id": "page.margins",
            "rule_text": "页面边距不符合规范（应：上3.5/下3.2/左右2.85cm，页眉1.5/页脚1.75cm）：" + "\u3001".join(viol),
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
        return int2cn(n) + "\u3001"
    if level == 2:
        return "\uff08" + int2cn(n) + "\uff09"
    if level == 3:
        return "%d." % n
    if level == 4:
        return "\uff08%d\uff09" % n
    return str(n)


# ---------------------------------------------------------------------------
# Document-level hints (comment only): cover content, green bg, caption content,
# TOC depth, page-number font (best-effort).
# ---------------------------------------------------------------------------
def _cover_missing_fields(structure, spec):
    """Required cover items that appear to be MISSING.

    If the step-2.5 AI review supplied structure['cover_present'] (the items it
    semantically confirmed are actually filled in \u2014 the reliable source for
    title-derived items like \u9879\u76ee\u540d\u79f0/\u8003\u6838\u5e74\u4efd, or a still-placeholder title),
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
    has_year = bool(re.search(r"(?:19|20)\d{2}", cover_text)) or ("\u5e74\u5ea6" in title_text)
    secrecy_values = cov.get("secrecy_values", [])
    present = set()
    for f in required:
        if f == "\u9879\u76ee\u540d\u79f0" and f in title_derived:
            if has_title:
                present.add(f)
        elif f == "\u8003\u6838\u5e74\u4efd" and f in title_derived:
            if has_year:
                present.add(f)
        elif f == "\u5bc6\u7ea7":
            # \u5bc6\u7ea7 is often written as its VALUE (\u673a\u5bc6/\u79d8\u5bc6/\u2026) with no "\u5bc6\u7ea7"
            # label, so accept either.
            if "\u5bc6\u7ea7" in cover_text or any(v in cover_text for v in secrecy_values):
                present.add(f)
        elif f in cover_text:
            present.add(f)
    return [f for f in required if f not in present]


def _cover_placeholder(structure, spec):
    """True if the title still contains an un-replaced template placeholder
    (XXXX / \u00d7\u00d7\u00d7\u00d7 / \u6a21\u677f)."""
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
                          "rule_text": "\u5c01\u76ae\u7f3a\u5c11\u5fc5\u8981\u4fe1\u606f\uff1a" + "\u3001".join(missing),
                          "comment": True})
        if _cover_placeholder(structure, spec):
            hints.append({"para_index": first_cover, "op": "hint",
                          "rule_id": "cover.placeholder",
                          "rule_text": "\u5c01\u9762\u6807\u9898\u7591\u4f3c\u4ecd\u4e3a\u6a21\u677f\u5360\u4f4d\u7b26\uff08\u5982 XXXX/\u6a21\u677f\uff09\uff0c\u8bf7\u66ff\u6362\u4e3a\u5b9e\u9645\u9879\u76ee\u540d\u79f0\u4e0e\u8003\u6838\u5e74\u4efd",
                          "comment": True})
        if spec["cover"].get("background_green") and not structure.get("has_background"):
            hints.append({"para_index": first_cover, "op": "hint",
                          "rule_id": "cover.green",
                          "rule_text": "\u672a\u68c0\u6d4b\u5230\u5c01\u76ae\u80cc\u666f\u989c\u8272\u8bbe\u7f6e\uff0c\u8bf7\u786e\u8ba4\u5c01\u76ae\u4e3a\u7eff\u8272",
                          "comment": True})
    # caption content presence
    for r in records:
        cap = r.get("caption")
        if cap and not cap.get("has_content"):
            kname = "\u56fe" if cap["kind"] == "figure" else "\u8868"
            hints.append({"para_index": r["i"], "op": "hint",
                          "rule_id": "caption.content",
                          "rule_text": "%s\u7f16\u53f7\u540e\u7f3a\u5c11\u5185\u5bb9\u8bf4\u660e" % kname,
                          "comment": True})
    return hints
