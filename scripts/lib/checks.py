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
   "clear_left_indent": bool, "set_jc": str|None,
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
    region = rec.get("region", "body")
    if region == "toc":
        # 目录条目只做字体/字号校验（仿宋 三号）；不动缩进/行距，避免破坏 TOC
        # 域代码自身的制表位/悬挂缩进结构。
        return _check_toc(rec, spec)

    # Caption paragraphs (图.../表...) have NO formatting rule in the spec, so we
    # do not touch their font/indent/line spacing here. They are handled only by
    # continuity (renumber) and the content-presence hint — this avoids inventing
    # rules the spec does not state.
    if rec.get("caption"):
        return None

    eff = rec["eff"]
    sets = {"set_east_asia": None, "set_ascii": None, "set_size_hp": None,
            "set_line_exact": None, "set_first_line_chars": None,
            "clear_left_indent": False, "set_jc": None}
    violations = []

    # -- Title on the cover --
    if region == "cover":
        looks_title = (eff.get("jc") == "center" and (eff.get("size_hp") or 0) >= 36) \
            or (eff.get("size_hp") == spec["title"]["size_hp"])
        if not looks_title:
            return None  # other cover lines: content handled by doc-level hints
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
        # \u9898\u76ee\u5e94\u65e0\u7f29\u8fdb\uff1b\u6b8b\u7559\u7684\u9996\u884c/\u5de6\u7f29\u8fdb\u4f1a\u8ba9 jc=center \u7684\u5c45\u4e2d\u57fa\u51c6\u504f\u79fb\uff0c
        # \u89c6\u89c9\u4e0a"\u5c45\u4e2d"\u5176\u5b9e\u662f\u76f8\u5bf9\u7f29\u8fdb\u540e\u7684\u53ef\u7528\u5bbd\u5ea6\u5c45\u4e2d\uff0c\u4e0d\u662f\u771f\u6b63\u7684\u9875\u9762\u5c45\u4e2d\u3002
        indent_keys = ("first_line_chars", "first_line", "left_chars", "left",
                       "start_chars", "start", "hanging_chars", "hanging")
        if any(eff.get(k) for k in indent_keys):
            sets["set_first_line_chars"] = 0
            sets["clear_left_indent"] = True
            violations.append(
                "\u9898\u76ee\u4e0d\u5e94\u6709\u7f29\u8fdb\uff08\u9700\u6e05\u9664\u9996\u884c/\u5de6\u7f29\u8fdb\uff0c\u5426\u5219\u5c45\u4e2d\u4e0d\u51c6\u786e\uff09")
        return _mk_format(rec["i"], sets, violations)

    # -- Body region: headings / captions / normal body --
    if rec.get("is_heading"):
        lvl = rec.get("level") or 1
        lvl = min(max(lvl, 1), 4)
        h = spec["headings"][str(lvl)]
        _check_font_size(eff, h, sets, violations,
                         label="%d\u7ea7\u6807\u9898" % lvl)
        _check_first_line(eff, h, sets, violations)
        # headings are not forced to a specific line rule by the spec table
    else:
        b = spec["body"]
        _check_font_size(eff, b, sets, violations, label="\u6b63\u6587")
        # line spacing: fixed value 28pt (exact 560)
        ls = spec["line_spacing"]
        if not (eff.get("line") == ls["line_twips"] and eff.get("line_rule") == ls["line_rule"]):
            sets["set_line_exact"] = ls["line_twips"]
            violations.append("\u884c\u8ddd\u5e94\u4e3a\u56fa\u5b9a\u503c28\u78c5")
        _check_first_line(eff, b, sets, violations)

    return _mk_format(rec["i"], sets, violations)


def _check_toc(rec, spec):
    t = spec.get("toc") or {}
    if not t.get("east_asia") or not t.get("size_hp"):
        return None  # spec not configured for toc fonts; nothing to check
    eff = rec["eff"]
    sets = {"set_east_asia": None, "set_ascii": None, "set_size_hp": None,
            "set_line_exact": None, "set_first_line_chars": None,
            "clear_left_indent": False, "set_jc": None}
    violations = []
    _check_font_size(eff, t, sets, violations, label="目录")
    return _mk_format(rec["i"], sets, violations)


def _check_font_size(eff, spec_entry, sets, violations, label):
    if not _font_ok(eff.get("east_asia"), spec_entry):
        sets["set_east_asia"] = spec_entry["east_asia"]
        violations.append("%s\u4e2d\u6587\u5b57\u4f53\u5e94\u4e3a%s\uff08\u5b9e\u9645\uff1a%s\uff09"
                          % (label, spec_entry["east_asia"], eff.get("east_asia")))
    if eff.get("size_hp") != spec_entry["size_hp"]:
        sets["set_size_hp"] = spec_entry["size_hp"]
        violations.append("%s\u5b57\u53f7\u5e94\u4e3a\u4e09\u53f7" % label)


def _check_first_line(eff, spec_entry, sets, violations):
    want = spec_entry.get("first_line_chars")
    if want is None:
        return
    if eff.get("first_line_chars") != want:
        sets["set_first_line_chars"] = want
        violations.append("\u9996\u884c\u7f29\u8fdb\u5e94\u4e3a2\u5b57\u7b26")
    # left indent must NOT be stacked on top of first-line indent
    if any(eff.get(k) for k in ("left_chars", "left", "start_chars", "start")):
        sets["clear_left_indent"] = True
        if "\u9996\u884c\u7f29\u8fdb\u5e94\u4e3a2\u5b57\u7b26" not in violations:
            violations.append("\u5e94\u4ec5\u4fdd\u7559\u9996\u884c\u7f29\u8fdb\uff0c\u9700\u6e05\u9664\u5de6\u7f29\u8fdb")
        else:
            violations.append("\u540c\u65f6\u9700\u6e05\u9664\u5de6\u7f29\u8fdb\uff08\u4e0d\u5e94\u4e0e\u9996\u884c\u7f29\u8fdb\u53e0\u52a0\uff09")


def _mk_format(i, sets, violations):
    if not violations:
        return None
    fix = {"para_index": i, "op": "format",
           "rule_id": "combined",
           "rule_text": "\u3010XAgent\u683c\u5f0f\u3011" + "\uff1b".join(violations),
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
            "rule_text": "【XAgent格式】页面边距不符合规范（应：上3.5/下3.2/左右2.85cm，页眉1.5/页脚1.75cm）：" + "\u3001".join(viol),
            "comment": False}


# ---------------------------------------------------------------------------
# Continuity (headings + captions) -> renumber fixes
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
        counters[lvl] += 1
        for d in range(lvl + 1, 5):
            counters[d] = 0
        expected = counters[lvl]
        expected_token = _heading_token(lvl, expected)
        raw = r.get("num_raw")
        if raw is None:
            continue  # no recognizable leading number at all; nothing safe to replace
        if raw != expected_token:
            # Covers BOTH kinds of violation with one comparison: sequence gaps
            # (e.g. "\u4e00\u3001" skipping to "\u4e09\u3001") AND format errors (e.g. arabic
            # "3\u3001" used where the level-1 spec requires the Chinese numeral
            # "\u4e09\u3001") \u2014 either way the rendered label differs from the
            # position-derived canonical token, so it gets corrected.
            fixes.append({
                "para_index": r["i"], "op": "renumber_heading", "level": lvl,
                "new_token": expected_token,
                "rule_id": "heading.continuity",
                "rule_text": "\u3010XAgent\u683c\u5f0f\u3011%d\u7ea7\u6807\u9898\u5e8f\u53f7\u683c\u5f0f\u6216\u8fde\u7eed\u6027\u6709\u8bef\uff0c\u5e94\u4e3a\u201c%s\u201d\uff08\u539f\u4e3a\u201c%s\u201d\uff09" % (lvl, expected_token, raw),
                "comment": True,
            })
            # keep counters at expected (we renumbered to expected)
    # -- captions --
    # decide flat vs chapter-based per kind from the raw numbers present
    for kind in ("figure", "table"):
        seq = [r for r in records if r.get("region") == "body" and r.get("caption")
               and r["caption"]["kind"] == kind]
        if not seq:
            continue
        chapter_based = any(any(s in r["caption"]["num_raw"] for s in ("-", ".", "\u2013"))
                            for r in seq)
        if chapter_based:
            group_counter = {}
            for r in seq:
                raw = r["caption"]["num_raw"]
                sep = next((s for s in ("-", ".", "\u2013") if s in raw), "-")
                prefix = raw.rsplit(sep, 1)[0]
                group_counter[prefix] = group_counter.get(prefix, 0) + 1
                expected = "%s%s%d" % (prefix, sep, group_counter[prefix])
                if raw != expected:
                    fixes.append(_caption_fix(r, kind, expected, raw))
        else:
            n = 0
            for r in seq:
                n += 1
                expected = str(n)
                if r["caption"]["num_raw"] != expected:
                    fixes.append(_caption_fix(r, kind, expected, r["caption"]["num_raw"]))
    return fixes


def _caption_fix(r, kind, new_num, old_num):
    kname = "\u56fe" if kind == "figure" else "\u8868"
    return {
        "para_index": r["i"], "op": "renumber_caption", "kind": kind,
        "new_num": new_num,
        "rule_id": "caption.%s.continuity" % kind,
        "rule_text": "\u3010XAgent\u683c\u5f0f\u3011%s\u7f16\u53f7\u4e0d\u8fde\u7eed\uff0c\u5e94\u4e3a\u201c%s%s\u201d\uff08\u539f\u4e3a\u201c%s%s\u201d\uff09" % (kname, kname, new_num, kname, old_num),
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
def doc_hints(structure, spec):
    records = structure["records"]
    hints = []
    # cover content presence
    cover_text = "".join(r["text"] for r in records if r.get("region") == "cover")
    missing = [f for f in spec["cover"]["required_fields"] if f not in cover_text]
    # attach cover hints to the first cover paragraph
    first_cover = next((r["i"] for r in records if r.get("region") == "cover"), None)
    if first_cover is not None:
        if missing:
            hints.append({"para_index": first_cover, "op": "hint",
                          "rule_id": "cover.fields",
                          "rule_text": "\u3010XAgent\u683c\u5f0f\u3011\u5c01\u76ae\u7f3a\u5c11\u5fc5\u8981\u4fe1\u606f\uff1a" + "\u3001".join(missing),
                          "comment": True})
        if spec["cover"].get("background_green") and not structure.get("has_background"):
            hints.append({"para_index": first_cover, "op": "hint",
                          "rule_id": "cover.green",
                          "rule_text": "\u3010XAgent\u683c\u5f0f\u3011\u672a\u68c0\u6d4b\u5230\u5c01\u76ae\u80cc\u666f\u989c\u8272\u8bbe\u7f6e\uff0c\u8bf7\u786e\u8ba4\u5c01\u76ae\u4e3a\u7eff\u8272",
                          "comment": True})
    # caption content presence
    for r in records:
        cap = r.get("caption")
        if cap and not cap.get("has_content"):
            kname = "\u56fe" if cap["kind"] == "figure" else "\u8868"
            hints.append({"para_index": r["i"], "op": "hint",
                          "rule_id": "caption.content",
                          "rule_text": "\u3010XAgent\u683c\u5f0f\u3011%s\u7f16\u53f7\u540e\u7f3a\u5c11\u5185\u5bb9\u8bf4\u660e" % kname,
                          "comment": True})
    return hints
