# -*- coding: utf-8 -*-
"""Apply fixes.json to a NEW document (never touches the original).

Usage:
    python 40_apply_fixes.py <workdir>

Reads:
    <workdir>/meta.json      (working_docx, original_ext/stem)
    <workdir>/fixes.json     (list of fix objects from checks)
Writes:
    <workdir>/out_pkg/       (unpacked, edited package)
    <workdir>/formatted.docx (rezipped result)
    <workdir>/apply_report.json

All edits are applied as DIRECT paragraph/run properties (overrides) so the
EFFECTIVE format matches the spec regardless of styles.xml. Paragraph indexing
uses the SAME iter_body_paragraphs() used by extraction, so para_index aligns.

Every fix with "comment": true gets an XAgent comment whose text is the
violated rule (rule_text).
"""
import copy
import json
import os
import re
import sys

sys.path.insert(0, os.path.join(os.path.dirname(__file__), "lib"))

from lxml import etree
from docxcommon import (qn, parse_xml, unzip_docx, rezip_docx,
                        iter_body_paragraphs, in_textbox,
                        StyleResolver, load_numbering_levels,
                        get_style_id, get_pPr, get_mark_rpr,
                        char_styles_overriding, ELEMENT_ORDER, ordered_insert,
                        local_name, order_for, para_inline_items)
from commentwriter import CommentWriter
from headings import ANY_LABEL_RE, CAPTION_STYLE_HINTS
from checks import load_default_spec, paragraph_role, is_unnumbered_section
import canonstyles
from canonstyles import STYLE_ID_BY_ROLE

# Paragraph styles that format the AUTO-GENERATED table-of-contents entries.
# Word regenerates these paragraphs from their style (not from direct
# formatting) whenever the TOC field refreshes — and we set updateFields=true
# so it refreshes on open — so the only way to make the TOC font/size/indent
# actually stick is to patch the styles themselves, not just the current
# entry runs.
# "Contents N" is LibreOffice's name for TOC entry styles (Word uses "TOC N" /
# "目录 N"); match it too so a LibreOffice-converted document's TOC styles are
# patched — otherwise a refreshed TOC rebuilds from an un-patched style and
# loses the per-level indent.
RE_TOC_STYLE_ID = re.compile(r"^(?:toc|contents?)\s*\d+$", re.IGNORECASE)
RE_TOC_STYLE_NAME = re.compile(r"^(?:toc|目录|contents?)\s*\d+$", re.IGNORECASE)

# Level-independent: whatever leading enumeration label a heading currently
# has (regardless of numeral system / punctuation, and regardless of whether
# it matches the shape its assigned level "should" use) gets stripped and
# replaced with the canonical token computed by checks.py. Shared with
# 20_extract_structure.py / 27_apply_review.py via lib/headings.py so all
# three never drift out of sync.
STRIP_HEADING = ANY_LABEL_RE
STRIP_CAPTION = re.compile(r"^\s*(?:图|表)\s*[0-9]+(?:[-\.–][0-9]+)?")
# A bare 图/表 prefix with NO number. Used to strip the residual prefix left
# behind when a caption's number lived in a Word field that gets deleted (the
# static "图/表" survives the field removal), or a 表标题-styled line that
# carries the 图/表 word but no digit -- so writing the new 图N/表N never
# doubles the prefix ("图1图 说明").
STRIP_CAPTION_RESIDUE = re.compile(r"^\s*(?:图|表)")


# ---------------------------------------------------------------------------
# pPr / rPr helpers (create-or-get, keep OOXML child order roughly valid)
# ---------------------------------------------------------------------------
def _get_or_make(parent, tag, before_tags=()):
    """取子元素，没有就**按 schema 顺序**建一个。

    WordprocessingML 的复杂类型几乎都是 `xsd:sequence`——子元素顺序错了，文件仍是
    良构 XML、zip 也完好，但**真实 Word 会拒绝打开**（"发现无法读取的内容，是否恢复
    此文档的内容"）。所以这里优先查 `docxcommon.ELEMENT_ORDER` 决定插入位置，而不是
    一律 append；只有该容器没登记顺序时才退回旧的 before_tags/append 行为。
    调用方因此不必再自己记 pPr/rPr 的子元素顺序。"""
    el = parent.find(qn(tag))
    if el is not None:
        return el
    el = etree.Element(qn(tag))
    order = order_for(parent)
    if order:
        return ordered_insert(parent, el, order)
    # insert before the first of before_tags that exists, else append
    anchor = None
    for bt in before_tags:
        cand = parent.find(qn(bt))
        if cand is not None:
            anchor = cand
            break
    if anchor is not None:
        anchor.addprevious(el)
    else:
        parent.append(el)
    return el


def get_pPr(p):
    pPr = p.find(qn("w:pPr"))
    if pPr is None:
        pPr = etree.Element(qn("w:pPr"))
        p.insert(0, pPr)
    return pPr


def _iter_runs(p):
    # Recurse (w:hyperlink-wrapped runs — e.g. every TOC entry — must be
    # reached), but skip runs that live inside a nested textbox
    # (w:txbxContent): those belong to a different logical paragraph and
    # must not be restyled as a side effect of editing the host paragraph.
    for r in p.iter(qn("w:r")):
        if in_textbox(r, stop_at=p):
            continue
        yield r


def _run_rpr(r):
    rpr = r.find(qn("w:rPr"))
    if rpr is None:
        rpr = etree.Element(qn("w:rPr"))
        r.insert(0, rpr)
    return rpr


def _set_fonts(rpr, east_asia=None, ascii_=None):
    rf = _get_or_make(rpr, "w:rFonts")
    if east_asia is not None:
        rf.set(qn("w:eastAsia"), east_asia)
    if ascii_ is not None:
        rf.set(qn("w:ascii"), ascii_)
        rf.set(qn("w:hAnsi"), ascii_)


def _set_size(rpr, half_pt):
    sz = _get_or_make(rpr, "w:sz")
    sz.set(qn("w:val"), str(half_pt))
    szcs = _get_or_make(rpr, "w:szCs")
    szcs.set(qn("w:val"), str(half_pt))


def _set_bold(rpr, on=True):
    """加粗开关。`w:b` 管中日韩与拉丁正文、`w:bCs` 管复杂文种——两个都写，否则同一段里
    的西文/数字可能不跟着加粗。"""
    for tag in ("w:b", "w:bCs"):
        el = _get_or_make(rpr, tag)
        el.set(qn("w:val"), "1" if on else "0")


def _apply_run_props(p, east_asia, ascii_, size_hp, bold=None):
    """Apply font/size/bold to every content run + the paragraph-mark rPr."""
    if east_asia is None and ascii_ is None and size_hp is None and bold is None:
        return   # 无事可做：别顺手给每个 run 建一个空 <w:rPr/> 壳
    for r in _iter_runs(p):
        rpr = _run_rpr(r)
        if east_asia is not None or ascii_ is not None:
            _set_fonts(rpr, east_asia, ascii_)
        if size_hp is not None:
            _set_size(rpr, size_hp)
        if bold is not None:
            _set_bold(rpr, bold)
    # paragraph mark run properties (pPr/rPr)
    pPr = get_pPr(p)
    mark = _get_or_make(pPr, "w:rPr",
                        before_tags=("w:sectPr",))
    if east_asia is not None or ascii_ is not None:
        _set_fonts(mark, east_asia, ascii_)
    if size_hp is not None:
        _set_size(mark, size_hp)
    if bold is not None:
        _set_bold(mark, bold)


def _set_line_exact(pPr, line_twips, line_rule="exact"):
    """Set line spacing. line_rule="exact" (固定值, twips) or "auto" (倍数，
    line=480 => 2倍). The caller passes the spec's own rule so 封面要素的 2倍行距
    (auto) and正文/标题的固定28磅(exact) share one apply path."""
    sp = _get_or_make(pPr, "w:spacing")
    sp.set(qn("w:line"), str(line_twips))
    sp.set(qn("w:lineRule"), line_rule)
    # remove auto-spacing that would override the fixed value
    for a in ("w:beforeAutospacing", "w:afterAutospacing"):
        if sp.get(qn(a)) is not None:
            del sp.attrib[qn(a)]


def _clear_space_before_after(pPr):
    """Zero the paragraph's 段前/段后 spacing as a DIRECT override (正文规范：
    去除段前段后). Explicit 0 (not attribute deletion) so an inherited非零
    段前/段后 from the style is overridden; the *Lines forms are dropped so they
    can't re-supply spacing alongside our 0."""
    sp = _get_or_make(pPr, "w:spacing")
    sp.set(qn("w:before"), "0")
    sp.set(qn("w:after"), "0")
    for a in ("w:beforeLines", "w:afterLines",
              "w:beforeAutospacing", "w:afterAutospacing"):
        if sp.get(qn(a)) is not None:
            del sp.attrib[qn(a)]


def _char_twips(chars, size_hp):
    """Character-unit indent (hundredths of a char, e.g. 200 = 2 chars) -> its
    absolute width in twips at a given font size (half-points).

    One CJK character is one em wide, so its width in twips is size_hp * 10
    (size_hp/2 pt * 20 twips/pt). Used to emit an ABSOLUTE w:firstLine/w:left
    alongside the char-based value so the direct override also wins over an
    absolute indent inherited from a style / numbering level (see
    _set_first_line_and_clear_left)."""
    return int(round(chars / 100.0 * size_hp * 10))


def _set_first_line_and_clear_left(pPr, first_line_chars, clear_left, clear_right=False,
                                   set_left_chars=None):
    """Write the paragraph's indent as a DIRECT override, **纯字符单位**。

    严格-spec §2.2：实际缩进量只写 `w:firstLineChars`/`w:leftChars`，**绝不补非零的
    绝对伴随值**（`w:firstLine`/`w:left` 的 twips）——补了 Word 就把缩进显示成厘米而
    不是"N 字符"，那本身即不合规。要清掉的方向仍显式写 0（零值没有单位歧义），这样
    才能压过继承来的绝对缩进。

    **历史（别照直觉加回来）**：这里曾按陷阱 #10 给每个字符单位值配一个绝对伴随值，
    用来压过 LibreOffice 转换出的、继承自样式/编号层的绝对 hanging。方案C 阶段2 之后
    两条继承路径都被根治——编号层由甲法克隆钳中和（`_clamp_numbering_indent`），样式层
    由 canonical 命名样式注入+指派接管（`_inject_canonical_styles`/`_assign_canonical_style`，
    canonical 样式自身把左右缩进显式归零、不带 hanging）——伴随值失去存在理由，删掉。
    本函数如今只服务**没有 canonical 样式承载缩进**的角色（目录条目：它的缩进由文档
    自己的 TOC 样式承载，见 `_patch_toc_styles`）。"""
    ind = _get_or_make(pPr, "w:ind")
    if first_line_chars is not None:
        ind.set(qn("w:firstLineChars"), str(first_line_chars))
        if ind.get(qn("w:firstLine")) is not None:
            del ind.attrib[qn("w:firstLine")]
        # remove hanging: firstLine and hanging are mutually exclusive, and any
        # direct hanging would otherwise take precedence over our first-line.
        for a in ("w:hanging", "w:hangingChars"):
            if ind.get(qn(a)) is not None:
                del ind.attrib[qn(a)]
    if set_left_chars is not None:
        # Set the left indent to a SPECIFIC character count (TOC per-level
        # indent: 0/200/400). leftChars governs (char-based, East-Asian aware);
        # the absolute w:left form and the w:start synonyms are dropped.
        ind.set(qn("w:leftChars"), str(set_left_chars))
        for a in ("w:left", "w:startChars", "w:start"):
            if ind.get(qn(a)) is not None:
                del ind.attrib[qn(a)]
    elif clear_left:
        # Force the left indent to 0 rather than merely deleting the direct
        # attribute: the indent we need to override is frequently INHERITED
        # from the paragraph style, and deleting a direct attribute that isn't
        # there leaves the style's indent in effect. An explicit direct 0 wins
        # over the inherited value. The w:start/w:startChars synonyms are
        # removed so they can't re-supply a non-zero left indent alongside our 0.
        ind.set(qn("w:leftChars"), "0")
        ind.set(qn("w:left"), "0")
        for a in ("w:startChars", "w:start"):
            if ind.get(qn(a)) is not None:
                del ind.attrib[qn(a)]
    if clear_right:
        # Same reasoning for the right indent. Zeroing this is what makes
        # jc=center actually center on the full page width instead of the width
        # left after an un-cleared (often style-inherited) right indent.
        ind.set(qn("w:rightChars"), "0")
        ind.set(qn("w:right"), "0")
        for a in ("w:endChars", "w:end"):
            if ind.get(qn(a)) is not None:
                del ind.attrib[qn(a)]


def _set_jc(pPr, val):
    jc = _get_or_make(pPr, "w:jc")
    jc.set(qn("w:val"), val)


def _default_char_unit_hp(pkg_dir, default=21):
    """The font size (half-points) Word uses for the CHARACTER UNIT of the
    *Chars indents (firstLineChars / leftChars) — i.e. how wide "1 字符" is.

    This is the DOCUMENT DEFAULT font size (styles.xml docDefaults → rPrDefault
    → rPr → sz), NOT the paragraph's own font. That distinction is the whole bug
    behind "标题首行缩进 1.13cm / 目录二级 1.13cm、三级 2.26cm": Word measures a
    2-char indent against the default 五号(10.5pt=21 半点) → 0.74cm, but computing
    it against a heading/目录's 三号(16pt=32) gives 1.13cm — ~1.52× too wide, and
    for the 三号 目录 wide enough to wrap the page number. Emitting the absolute
    companion at THIS size makes it match Word's native *Chars rendering (so the
    two forms never disagree) while still overriding an inherited absolute
    hanging. Falls back to 21 (五号) — Word's usual default — when docDefaults
    carries no explicit size."""
    path = os.path.join(pkg_dir, "word", "styles.xml")
    if not os.path.exists(path):
        return default
    try:
        root = parse_xml(path).getroot()
    except (OSError, ValueError):
        return default
    sz = root.find(qn("w:docDefaults") + "/" + qn("w:rPrDefault") + "/"
                   + qn("w:rPr") + "/" + qn("w:sz"))
    if sz is not None and sz.get(qn("w:val")):
        try:
            return int(sz.get(qn("w:val")))
        except ValueError:
            pass
    return default


# whitespace stripped from a paragraph's ends (space / tab / full-width /
# no-break space). Internal runs of these are left alone -- only the leading
# and/or trailing padding is removed.
_STRIP_WS = " \t　   "


def _strip_para_ws(p, mode):
    """Strip leading and/or trailing whitespace from a paragraph's visible text.

    mode: 'leading' | 'trailing' | 'both'. Operates on the w:t elements of the
    paragraph's content runs in document order, so multi-run text is handled
    correctly. Internal whitespace (e.g. the fill-in blanks in
    "20   年   月至20   年   月") is preserved -- only the padding at the very
    start and/or very end of the paragraph is removed. Leaves each surviving
    w:t's xml:space attribute untouched."""
    ts = [t for r in _iter_runs(p) for t in r.findall(qn("w:t"))]
    if not ts:
        return
    if mode in ("leading", "both"):
        for t in ts:
            s = t.text or ""
            stripped = s.lstrip(_STRIP_WS)
            t.text = stripped
            if stripped:
                break  # first run with real content reached; stop
    if mode in ("trailing", "both"):
        for t in reversed(ts):
            s = t.text or ""
            stripped = s.rstrip(_STRIP_WS)
            t.text = stripped
            if stripped:
                break


# ---------------------------------------------------------------------------
# text renumbering (collapse to first w:t; labels are single-style lines)
# ---------------------------------------------------------------------------
XML_SPACE = "{http://www.w3.org/XML/1998/namespace}space"


def _field_run_spans(p):
    """[(begin_idx, end_idx)] into list(p) for each top-level field
    (w:fldChar begin..matching end). Handles nesting via a depth counter."""
    kids = list(p)
    spans, depth, start = [], 0, None
    for idx, ch in enumerate(kids):
        if ch.tag != qn("w:r"):
            continue
        fc = ch.find(qn("w:fldChar"))
        if fc is None:
            continue
        typ = fc.get(qn("w:fldCharType"))
        if typ == "begin":
            if depth == 0:
                start = idx
            depth += 1
        elif typ == "end" and depth > 0:
            depth -= 1
            if depth == 0 and start is not None:
                spans.append((start, idx))
                start = None
    return spans, kids


def _replace_leading(p, strip_re, new_prefix, residue_re=None):
    """Replace a paragraph's leading numbering label with new_prefix.

    Field-aware: when the leading number is produced by a Word FIELD (e.g.
    a heading numbered by "{ = 1 \\* Arabic }" — the digit lives in the field
    RESULT, not as typed text), the whole field is deleted. Otherwise Word
    would regenerate the number from the field code on the next field refresh
    (we set updateFields=true), wiping our replacement AND, because the old
    code collapsed all text into that field-result run, the real title text
    with it — leaving only the recomputed "1". After removing any label field,
    the remaining leading separator whitespace is trimmed and the new token is
    prepended to the actual content run."""
    all_t = [t for t in p.iter(qn("w:t"))]
    if not all_t:
        return False
    full = "".join(t.text or "" for t in all_t)
    m = strip_re.match(full)
    if not m:
        return False
    end = m.end()

    # char range of each w:t within the concatenated text
    off, pos = {}, 0
    for t in all_t:
        L = len(t.text or "")
        off[id(t)] = pos
        pos += L

    # delete every field whose visible result falls inside the leading label
    spans, kids = _field_run_spans(p)
    for s, e in spans:
        field_runs = kids[s:e + 1]
        field_ts = [t for r in field_runs for t in r.iter(qn("w:t"))]
        starts = [off[id(t)] for t in field_ts if id(t) in off]
        if starts and min(starts) < end:
            # Remove only the field's RUNS (begin/instrText/separate/result/end);
            # keep any bookmarkStart/End that sit inside the field span — those
            # are the heading's TOC anchor and must survive, or the TOC entry
            # breaks.
            for r in field_runs:
                if r.tag != qn("w:r"):
                    continue
                parent = r.getparent()
                if parent is not None:
                    parent.remove(r)

    rem_t = [t for t in p.iter(qn("w:t"))]
    if not rem_t:
        r = etree.SubElement(p, qn("w:r"))
        t = etree.SubElement(r, qn("w:t"))
        t.text = new_prefix
        t.set(XML_SPACE, "preserve")
        return True
    full2 = "".join(t.text or "" for t in rem_t)
    m2 = strip_re.match(full2)
    # if the label survived as text, strip it; if only a bare prefix survived
    # (the number was inside a now-deleted field), strip that residual prefix;
    # if it was entirely inside the deleted field, only its trailing separator
    # whitespace remains -> trim it
    if m2:
        rest = full2[m2.end():]
    elif residue_re is not None and residue_re.match(full2):
        rest = full2[residue_re.match(full2).end():]
    else:
        rest = full2.lstrip(" \t　")
    rem_t[0].text = new_prefix + rest
    rem_t[0].set(XML_SPACE, "preserve")
    for t in rem_t[1:]:
        t.text = ""
    return True


def _prepend_text(p, prefix):
    """Insert `prefix` at the very start of a paragraph's text without
    disturbing existing runs. Used to ADD a caption number (图N/表N) to a
    caption-styled paragraph that has no leading 图/表 label to replace."""
    ts = [t for t in p.iter(qn("w:t"))]
    if ts:
        ts[0].text = prefix + (ts[0].text or "")
        ts[0].set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
        return True
    # No text runs at all: create a minimal run carrying the number.
    r = etree.SubElement(p, qn("w:r"))
    t = etree.SubElement(r, qn("w:t"))
    t.text = prefix
    t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    return True


def _suppress_auto_numbering(p):
    """Cancel any inherited (style-level) or direct Word automatic list
    numbering on this paragraph by writing a direct numId=0 override (the OOXML
    "no numbering" value), so a caption we renumber to static 图N/表N text is not
    ALSO auto-numbered by Word (which would stack two numbers)."""
    pPr = get_pPr(p)
    for npr in pPr.findall(qn("w:numPr")):
        pPr.remove(npr)
    numPr = _get_or_make(pPr, "w:numPr",
                         before_tags=("w:spacing", "w:ind", "w:jc", "w:rPr", "w:sectPr"))
    numId = _get_or_make(numPr, "w:numId")
    numId.set(qn("w:val"), "0")


def _apply_renumber_caption(p, fix):
    if fix.get("strip_auto"):
        # Original number came from Word automatic numbering (w:numPr); cancel it
        # so the static 图N/表N we write below is the only number that renders.
        _suppress_auto_numbering(p)
    kname = "图" if fix.get("kind") == "figure" else "表"
    new_prefix = kname + str(fix["new_num"])
    # 1) a proper 图/表 + number label (typed, or a field whose result is in
    #    the text) -- replace it; residue_re cleans a bare 图/表 left if the
    #    number was a field that got deleted, so we never double the prefix.
    if _replace_leading(p, STRIP_CAPTION, new_prefix, residue_re=STRIP_CAPTION_RESIDUE):
        return True
    # 2) a bare 图/表 prefix with no number (style-detected caption whose number
    #    was missing) -- replace just that prefix.
    if _replace_leading(p, STRIP_CAPTION_RESIDUE, new_prefix):
        return True
    # 3) no 图/表 label in the text at all (auto-numbered caption whose number
    #    lived only in numPr, or a 表标题 line reading just "设备清单") -- add one.
    return _prepend_text(p, new_prefix + " ")


def _heading_insert_prefix(token):
    """Prefix used when INSERTING a missing heading number. Arabic dotted tokens
    ("1.") read better with a trailing space before the title; full-width
    punctuation tokens ("一、"/"（一）"/"（1）") need none."""
    return token + " " if token.endswith(".") else token


def _apply_renumber_heading(p, fix):
    if fix.get("insert"):
        # Confirmed heading that lost its number entirely -- prepend the
        # position token (there is no existing label to replace).
        return _prepend_text(p, _heading_insert_prefix(fix["new_token"]))
    return _replace_leading(p, STRIP_HEADING, fix["new_token"])


def _heading_number_tab(p):
    """把**手写在文字里**的标题序号与标题文字之间的分隔改成一个真正的制表符。

    为什么需要：自动编号那条路径的制表符来自编号定义的 `suff=tab`；手写序号是纯文本，
    `renumber_heading` 只把序号归位、分隔符原样保留，于是"同一份规范、两类文档长得不
    一样"（用户实测："原文档是手写序号的，生成的标题编号后没有制表符"）。落点与自动
    编号一致，都由 `settings` 的 `defaultTabStop`（2 字符）决定。

    制表符**必须是 `w:tab` 元素**——`w:t` 里塞一个 `\\t` 字符 Word 会当成普通空白。
    做法：在序号结束处把那个 `w:t` 一切为二，中间插 `w:tab`；序号与标题文字之间原有的
    空白与制表符一并清掉（只清到第一个非空白字符为止，标题内部的制表符不动）。
    重复运行是安全的：拆完再拆结果一样（判定层还会先按 `num_tab` 判断已经有制表符了）。

    返回是否改动了段落。"""
    items = para_inline_items(p)
    ts = [el for el in items if el.tag == qn("w:t")]
    if not ts:
        return False
    label = ANY_LABEL_RE.match("".join(t.text or "" for t in ts))
    if not label:
        return False
    n = label.end()
    if not n:
        return False

    # 序号结束的位置落在哪个 w:t 的第几个字符
    pos, split_el, split_k = 0, None, 0
    for t in ts:
        L = len(t.text or "")
        if split_el is None and pos + L >= n:
            split_el, split_k = t, n - pos
        pos += L
    tail = (split_el.text or "")[split_k:]

    # 序号之后、第一个非空白字符之前的东西：空白要吃掉、制表符要删掉（下面重新插一个）
    stripped = tail.lstrip(" \t　")
    rest = items[items.index(split_el) + 1:]
    # `_replace_leading`（归位序号）会把整段文字并进第一个 w:t、把后面的清空，原有的
    # 制表符就被甩到文字后面成了尾巴。识别这个形态（后面的 w:t 全空）后把残留的制表符
    # 全部清掉；否则只清"序号与标题文字之间"那一段，标题内部的制表符不动。
    collapsed = all(not (el.text or "") for el in rest if el.tag == qn("w:t"))
    drop, seen_text = [], bool(stripped)
    if collapsed:
        drop = [el for el in rest if el.tag == qn("w:tab")]
    else:
        for el in rest:
            if el.tag == qn("w:tab"):
                if not seen_text:
                    drop.append(el)
                continue
            if seen_text:
                break
            txt = el.text or ""
            if txt.strip():
                el.text = txt.lstrip(" \t　")
                seen_text = True
            else:
                el.text = ""
    if not (stripped or seen_text):
        return False          # 只有序号、没有标题文字：不加制表符
    for el in drop:
        parent = el.getparent()
        if parent is not None:
            parent.remove(el)

    split_el.text = (split_el.text or "")[:split_k]
    split_el.set(XML_SPACE, "preserve")
    run = split_el.getparent()
    at = list(run).index(split_el) + 1
    tab = etree.Element(qn("w:tab"))
    run.insert(at, tab)
    if stripped:
        t2 = etree.Element(qn("w:t"))
        t2.text = stripped
        t2.set(XML_SPACE, "preserve")
        run.insert(at + 1, t2)
    return True


# ---------------------------------------------------------------------------
# section (page margins) — apply to EVERY sectPr for whole-doc consistency
# ---------------------------------------------------------------------------
def _apply_section(doc_root, setmar):
    n = 0
    for sect in doc_root.iter(qn("w:sectPr")):
        pgmar = _get_or_make(sect, "w:pgMar")
        for k, v in setmar.items():
            pgmar.set(qn("w:" + k), str(v))
        n += 1
    return n


# ---------------------------------------------------------------------------
# settings.xml : force TOC field refresh on open
# ---------------------------------------------------------------------------
def _toc_style_level(sid, name):
    """Trailing digit of a TOC style id/name ('TOC1'/'toc 2'/'Contents 3')
    -> its level."""
    m = re.search(r"(?:toc|目录|contents?)\s*([1-9])", sid + " " + name, re.IGNORECASE)
    return int(m.group(1)) if m else None


def _patch_toc_styles(pkg_dir, toc_spec, char_unit_hp=21):
    """Force the TOC entry styles (toc 1..N) to the spec's font/size and the
    per-level indent (一级0/二级2字符/三级4字符), so a REFRESHED TOC renders
    per spec. Returns the number of styles patched. No-op when styles.xml or
    the toc spec is missing. ``char_unit_hp`` is the document character-unit
    size used to compute the absolute left companion — the default font size,
    NOT the TOC's own 三号 (see _default_char_unit_hp), so 二级/三级 come out at
    0.74cm/1.49cm rather than the too-wide 1.13cm/2.26cm that wraps page
    numbers.

    目录合规【只认 leftChars】——判定层只按 leftChars 判合规（`_check_toc`）。但样式
    里仍要写**制表位**：阶段0 的 canonical 参考件经用户 Word 逐项验收确认，目录条目的
    "编号→标题→点线→页码"排布靠 TOC 样式自带的左制表位 + 右点线制表位撑起来，缺了
    点线会跑到编号与标题之间、页码换行（陷阱 #12）。制表位没有字符单位形式，只能写
    绝对 twips，按 spec 的字符数 × 文档字符单位字号换算（Normal=五号 → 左 840/1050/1260
    ＋右点线 8665）——`char_unit_hp` 就是这把尺子，取 docDefaults 的字号而**不是**目录
    自己的小三，否则整体偏宽、三级页码换行。

    历史（别照直觉再删一次）：2026-07 曾以"Word updateFields 会自建制表位、写死会被
    盖掉"为由删掉这段；阶段0 Word 实测**推翻**了它——样式制表位确实生效。"""
    if not toc_spec or not toc_spec.get("east_asia") or not toc_spec.get("size_hp"):
        return 0
    path = os.path.join(pkg_dir, "word", "styles.xml")
    if not os.path.exists(path):
        return 0
    tree = parse_xml(path)
    root = tree.getroot()
    ea = toc_spec["east_asia"]
    sz = str(toc_spec["size_hp"])
    by_level = toc_spec.get("indent_chars_by_level") or {}
    patched = 0
    for st in root.findall(qn("w:style")):
        if st.get(qn("w:type")) != "paragraph":
            continue
        sid = st.get(qn("w:styleId")) or ""
        nm_el = st.find(qn("w:name"))
        name = (nm_el.get(qn("w:val")) if nm_el is not None else "") or ""
        if not (RE_TOC_STYLE_ID.match(sid) or RE_TOC_STYLE_NAME.match(name)):
            continue
        # rPr: east-asian font + size
        rpr = _get_or_make(st, "w:rPr", before_tags=())
        _set_fonts(rpr, east_asia=ea)
        _set_size(rpr, int(sz))
        # pPr: left indent per this style's level; no first-line/hanging indent
        lvl = _toc_style_level(sid, name)
        want_left = by_level.get(str(lvl), 0) if (lvl is not None) else 0
        ppr = _get_or_make(st, "w:pPr", before_tags=("w:rPr",))
        # 制表位：左（编号→标题）+ 右点线（→页码）。w:tabs 在 CT_PPr 里排在
        # spacing/ind 之前，_get_or_make 的 before_tags 保证插对位置。
        tabs_spec = canonstyles.toc_level_props(toc_spec, lvl, char_unit_hp)["tabs"]
        if tabs_spec:
            tabs = _get_or_make(ppr, "w:tabs",
                                before_tags=("w:spacing", "w:ind", "w:jc", "w:outlineLvl",
                                             "w:rPr"))
            for old in tabs.findall(qn("w:tab")):
                tabs.remove(old)
            for val, pos, leader in tabs_spec:
                tab = etree.SubElement(tabs, qn("w:tab"))
                tab.set(qn("w:val"), val)
                if leader:
                    tab.set(qn("w:leader"), leader)
                tab.set(qn("w:pos"), str(pos))
        ind = _get_or_make(ppr, "w:ind")
        ind.set(qn("w:leftChars"), str(want_left))
        # absolute companion (0 for 一级) so the per-level indent overrides any
        # absolute left/hanging inherited from a basedOn parent style. Computed
        # at the document char-unit size (NOT the TOC's own 三号) so it matches
        # Word's native leftChars width and doesn't over-indent / wrap.
        ind.set(qn("w:left"), str(_char_twips(want_left, char_unit_hp)))
        for a in ("w:firstLineChars", "w:firstLine"):
            ind.set(qn(a), "0")
        for a in ("w:startChars", "w:start", "w:hanging", "w:hangingChars"):
            if ind.get(qn(a)) is not None:
                del ind.attrib[qn(a)]
        patched += 1
    if patched:
        tree.write(path, xml_declaration=True, encoding="UTF-8", standalone=True)
    return patched


def _style_name_is_caption(name_or_id):
    s = (name_or_id or "").lower()
    return any(h.lower() in s for h in CAPTION_STYLE_HINTS)


def _unhide_number_rpr(rpr):
    """Strip run properties that make an auto-generated number invisible:
    a solid non-white shading (a coloured block over it), a zero font size,
    hidden-text flags, and a white/auto text colour. Returns True if changed.
    Only ever REMOVES hiding — the number then inherits normal formatting."""
    if rpr is None:
        return False
    changed = False
    shd = rpr.find(qn("w:shd"))
    if shd is not None:
        fill = (shd.get(qn("w:fill")) or "auto").lower()
        if fill not in ("auto", "ffffff"):
            rpr.remove(shd)
            changed = True
    for tag in ("w:sz", "w:szCs"):
        el = rpr.find(qn(tag))
        if el is not None and (el.get(qn("w:val")) or "0") == "0":
            rpr.remove(el)
            changed = True
    for tag in ("w:vanish", "w:specVanish", "w:webHidden"):
        el = rpr.find(qn(tag))
        if el is not None and (el.get(qn("w:val")) or "true").lower() not in ("0", "false"):
            rpr.remove(el)
            changed = True
    clr = rpr.find(qn("w:color"))
    if clr is not None and (clr.get(qn("w:val")) or "").lower() in ("ffffff", "auto"):
        rpr.remove(clr)
        changed = True
    return changed


def _caption_abstract_num_ids(num_root, styles_root):
    """abstractNumIds that drive figure/table CAPTION numbering, found three
    ways so template variants are all covered:
      1. a caption paragraph style (name/id like 题注/图表标题/表标题/…) whose
         numPr → numId → abstractNum;
      2. a numbering level whose w:pStyle points at a caption style;
      3. a numbering level whose lvlText literally contains 图/表.
    """
    # styleId -> (name, numId used by that style)
    style_name, style_numid = {}, {}
    if styles_root is not None:
        for st in styles_root.findall(qn("w:style")):
            if st.get(qn("w:type")) != "paragraph":
                continue
            sid = st.get(qn("w:styleId"))
            nm = st.find(qn("w:name"))
            style_name[sid] = (nm.get(qn("w:val")) if nm is not None else "") or ""
            npr = st.find(qn("w:pPr") + "/" + qn("w:numPr") + "/" + qn("w:numId"))
            if npr is not None:
                style_numid[sid] = npr.get(qn("w:val"))
    caption_style_ids = {sid for sid, nm in style_name.items()
                         if _style_name_is_caption((sid or "") + " " + nm)}

    num2abs = {}
    for num in num_root.findall(qn("w:num")):
        a = num.find(qn("w:abstractNumId"))
        if a is not None:
            num2abs[num.get(qn("w:numId"))] = a.get(qn("w:val"))

    abstracts = set()
    for sid in caption_style_ids:                     # way 1
        nid = style_numid.get(sid)
        if nid and nid in num2abs:
            abstracts.add(num2abs[nid])
    for anum in num_root.findall(qn("w:abstractNum")):
        aid = anum.get(qn("w:abstractNumId"))
        for lvl in anum.findall(qn("w:lvl")):
            ps = lvl.find(qn("w:pStyle"))
            psid = ps.get(qn("w:val")) if ps is not None else None
            lt = lvl.find(qn("w:lvlText"))
            txt = (lt.get(qn("w:val")) if lt is not None else "") or ""
            if (psid in caption_style_ids                       # way 2
                    or _style_name_is_caption((psid or "") + " " + style_name.get(psid, ""))
                    or "图" in txt or "表" in txt):               # way 3
                abstracts.add(aid)
                break
    return abstracts


def _clean_caption_numbering(pkg_dir):
    """Un-hide auto-generated caption numbers (图N/表N).

    Some templates give the caption-numbering LEVEL run properties that hide
    the number — a solid dark shading (w:shd fill=000000, a black block over
    it), a zero font size, hidden-text flags, or a white colour. The number is
    really there and stays continuous (auto-numbered, 方案一); it is merely
    invisible. Every numbering level belonging to a caption abstractNum (found
    via caption styles / level pStyle / lvlText — see _caption_abstract_num_ids)
    has its hiding run properties stripped so 表1/图1 renders normally. The
    numId, num→abstractNum mapping and the styles' numPr are left untouched, so
    Word keeps auto-numbering. Returns the number of levels cleaned."""
    path = os.path.join(pkg_dir, "word", "numbering.xml")
    if not os.path.exists(path):
        return 0
    tree = parse_xml(path)
    root = tree.getroot()
    styles_path = os.path.join(pkg_dir, "word", "styles.xml")
    styles_root = parse_xml(styles_path).getroot() if os.path.exists(styles_path) else None

    caption_abstracts = _caption_abstract_num_ids(root, styles_root)
    cleaned = 0
    for anum in root.findall(qn("w:abstractNum")):
        if anum.get(qn("w:abstractNumId")) not in caption_abstracts:
            continue
        for lvl in anum.findall(qn("w:lvl")):
            if _unhide_number_rpr(lvl.find(qn("w:rPr"))):
                cleaned += 1
    if cleaned:
        tree.write(path, xml_declaration=True, encoding="UTF-8", standalone=True)
    return cleaned


# ---------------------------------------------------------------------------
# 方案C 甲法：钳住编号层（克隆而非改共享）
# ---------------------------------------------------------------------------
def _next_int_id(root, tag, attr, taken=()):
    """max existing int <tag attr=...> id + 1 (>= 1). ``taken`` excludes ids we
    already minted this run but haven't written to the tree yet."""
    ids = set(int(v) for v in taken)
    for el in root.findall(qn(tag)):
        v = el.get(qn(attr))
        try:
            ids.add(int(v))
        except (TypeError, ValueError):
            pass
    return (max(ids) + 1) if ids else 1


def _para_num_ref(p, resolver, numbering_levels):
    """Effective (numId, ilvl) for a paragraph, or (None, None) when it is not
    auto-numbered — numId absent, or the explicit '0' no-numbering override.
    Resolves through the style chain so a numPr inherited from the paragraph's
    style (not written directly on the paragraph) is still seen."""
    ppr, _ = resolver.resolve_cascade(get_style_id(p), get_pPr(p),
                                      get_mark_rpr(p), numbering_levels)
    nid = ppr.get("num_id")
    if not nid or nid == "0":
        return None, None
    try:
        il = int(ppr.get("ilvl") or 0)
    except (ValueError, TypeError):
        il = 0
    return nid, il


def _neutralize_level_indent(lvl):
    """In a CLONED numbering level, strip the indent that leaks into the
    paragraph: drop hanging/hangingChars and every first-line/left form, then
    pin left to 0. The paragraph carries its own char-unit first-line indent
    (firstLineChars) directly, so once the level supplies no competing indent
    that char value wins — no absolute companion needed (§2.2)."""
    ppr = _get_or_make(lvl, "w:pPr")
    ind = _get_or_make(ppr, "w:ind")
    for a in ("w:hanging", "w:hangingChars", "w:firstLine", "w:firstLineChars",
              "w:left", "w:leftChars", "w:start", "w:startChars"):
        if ind.get(qn(a)) is not None:
            del ind.attrib[qn(a)]
    ind.set(qn("w:left"), "0")
    ind.set(qn("w:leftChars"), "0")


# 克隆钳生成的 abstractNum 用 w:name 打这个前缀 + 源 abstractNumId 认领，重跑时复用
# 同一条而不是再克隆一份（否则每跑一次 numbering.xml 就多一条克隆）。
CLAMP_CLONE_MARKER = "FGWClampClone"


def _detach_clone_identity(clone):
    """把克隆的 abstractNum 与**源列表/源样式**的身份链接全部摘掉。

    克隆是一条**新的、独立的**列表，不能继续冒充源列表：

      * `w:styleLink` / `w:numStyleLink` —— 指向"编号样式"。两条 abstractNum 声明自己
        是同一个编号样式的定义是自相矛盾的，Word 打开时会判定文档需要修复
        （"发现无法读取的内容"）。
      * `w:tmpl` / `w:nsid` —— 列表标识。留着会让 Word 认为两条是同一个列表。
      * 每个级别上的 `w:pStyle` —— 它把"该级别"与"某个段落样式"绑定。克隆一份等于给
        同一个标题样式挂了两条编号定义，Word 究竟按哪条编号无从预期；而我们的段落已
        改指 canonical 样式 + 直接 numPr，本来就不需要这层样式绑定。"""
    for tag in ("w:styleLink", "w:numStyleLink", "w:tmpl", "w:nsid"):
        for el in clone.findall(qn(tag)):
            clone.remove(el)
    for lvl in clone.findall(qn("w:lvl")):
        for ps in lvl.findall(qn("w:pStyle")):
            lvl.remove(ps)


def _clone_marker(anum):
    """该 abstractNum 的 `w:name` 若是克隆钳打的标记，返回它，否则 None。"""
    nm = anum.find(qn("w:name"))
    val = nm.get(qn("w:val")) if nm is not None else None
    return val if (val and val.startswith(CLAMP_CLONE_MARKER)) else None


def _set_abstract_name(anum, value):
    """给 abstractNum 写 `w:name`（CT_AbstractNum 的 sequence：nsid → multiLevelType →
    tmpl → name → styleLink → numStyleLink → lvl，故按序插）。"""
    nm = anum.find(qn("w:name"))
    if nm is None:
        nm = ordered_insert(anum, etree.Element(qn("w:name")), order_for(anum))
    nm.set(qn("w:val"), value)


def _set_para_numid(p, new_num_id, ilvl):
    """Point a paragraph's numPr at ``new_num_id`` as a DIRECT override (works
    whether the original numPr was direct or inherited from the style chain).
    Preserves the paragraph's ilvl."""
    pPr = _get_or_make(p, "w:pPr")
    for npr in pPr.findall(qn("w:numPr")):
        pPr.remove(npr)
    numPr = _get_or_make(pPr, "w:numPr",
                         before_tags=("w:spacing", "w:ind", "w:jc", "w:rPr", "w:sectPr"))
    il = _get_or_make(numPr, "w:ilvl")
    il.set(qn("w:val"), str(ilvl))
    ni = _get_or_make(numPr, "w:numId")
    ni.set(qn("w:val"), str(new_num_id))


def _clamp_numbering_indent(pkg_dir, targets, resolver, numbering_levels):
    """方案C 甲法：钳住编号层。

    ``targets`` 是"自动编号且缩进违规"的段落元素列表。对每个段落解析其有效
    numId/ilvl；把落在**同一 abstractNum** 上的目标段落归为一组，**克隆**该 abstractNum
    （生成新 abstractNumId + 一个新 numId），在克隆里把这些段落用到的级别缩进中和
    （去 hanging、left 归 0），再把该组每个目标段落的 numPr 改指克隆的 numId。

    **按 abstractNum 分组、不是按 numId**（2026-07 修）：Word 文档里多个 `w:num` 指向
    同一条 `abstractNum` 极其常见（重复套用列表格式就会生成），而**它们共享同一个计数
    器**。早先按 numId 分组会为每个 numId 各克隆一份 abstractNum，等于把原本连号的一个
    多级列表**劈成几条互不相干的列表**——一级标题还对，二/三/四级从中间重新计数，用户
    实测到的"二级/三级/四级标题编号顺序出错"就是这么来的。按 abstractNum 归组、整组只
    克隆一次、共用一个新 numId，计数器才保持共享。

    为什么克隆而不原地改：一条 abstractNum 常被同级多段共享，原地改会溢到非目标
    段（bug #17"改一个标题行距、同级全变"）。只有拿到过缩进修复的段落才被改指
    克隆；未被修复的同 numId 段落仍指向原 abstractNum，格式不受影响。

    克隆必须**摘掉源身份**（`_detach_clone_identity`：styleLink/numStyleLink/tmpl/nsid
    与级别上的 pStyle），否则文档里会出现两条自称同一个编号样式、绑定同一批段落样式的
    列表——Word 打开时报"发现无法读取的内容"。重跑时按 `w:name` 上的标记**认领已有的
    克隆**，不再层层克隆（否则每跑一次 numbering.xml 就长一条）。

    **标题不走这条路**（2026-07 改）：一~四级标题的自动编号统一改指注入的 canonical
    四级列表（`canonstyles.heading_numbering_def`），逐级归零由"同一条多级列表"保证；
    这里只服务正文里的普通自动编号列表，它们的 `lvlRestart` 是作者的选择，不动。

    钳住编号层后，段落自身的字符单位首行缩进（firstLineChars，无绝对伴随值）即可
    生效——这是 #12/#14"继承 hanging"那一半的根治，替代绝对伴随值 hack（§2.2）。
    返回被改指的段落数。**Word 渲染需实测**（沙箱验不了，见 handoff §1）。"""
    path = os.path.join(pkg_dir, "word", "numbering.xml")
    if not os.path.exists(path) or not targets:
        return 0
    tree = parse_xml(path)
    root = tree.getroot()

    num2abs, abs_by_id = {}, {}
    for num in root.findall(qn("w:num")):
        a = num.find(qn("w:abstractNumId"))
        if a is not None:
            num2abs[num.get(qn("w:numId"))] = a.get(qn("w:val"))
    for anum in root.findall(qn("w:abstractNum")):
        abs_by_id[anum.get(qn("w:abstractNumId"))] = anum
    # 上一轮跑出来的克隆：{源 abstractNumId: 克隆的 abstractNumId}
    clone_of = {}
    for anum in root.findall(qn("w:abstractNum")):
        nm = anum.find(qn("w:name"))
        val = nm.get(qn("w:val")) if nm is not None else None
        if val and val.startswith(CLAMP_CLONE_MARKER):
            clone_of[val[len(CLAMP_CLONE_MARKER):]] = anum.get(qn("w:abstractNumId"))

    # group targets by source **abstractNum**（共享计数器的单位），keep each ilvl
    groups = {}   # src_abstractNumId -> {"ilvls": set, "paras": [(p, ilvl)]}
    for p in targets:
        nid, il = _para_num_ref(p, resolver, numbering_levels)
        if nid is None or nid not in num2abs:
            continue
        g = groups.setdefault(num2abs[nid], {"ilvls": set(), "paras": []})
        g["ilvls"].add(il)
        g["paras"].append((p, il))

    if not groups:
        return 0

    # where new abstractNum elements must go: after the last existing one
    # (schema requires all abstractNum before all num)
    last_abstract = None
    for anum in root.findall(qn("w:abstractNum")):
        last_abstract = anum

    new_abs_ids, new_num_ids = [], []
    changed = 0
    for src_aid, g in groups.items():
        src_anum = abs_by_id.get(src_aid)
        if src_anum is None:
            continue
        # 重跑：段落已经指向上一轮的克隆（或本轮源就是那条克隆）——就地把用到的级别
        # 再中和一遍即可，不再克隆克隆。
        reuse_aid = clone_of.get(src_aid)
        if reuse_aid is None and _clone_marker(src_anum):
            reuse_aid = src_aid
        if reuse_aid is not None and reuse_aid in abs_by_id:
            clone = abs_by_id[reuse_aid]
            for lvl in clone.findall(qn("w:lvl")):
                try:
                    if int(lvl.get(qn("w:ilvl"))) in g["ilvls"]:
                        _neutralize_level_indent(lvl)
                except (TypeError, ValueError):
                    continue
            reuse_num = next((n.get(qn("w:numId")) for n in root.findall(qn("w:num"))
                              if (n.find(qn("w:abstractNumId")) is not None
                                  and n.find(qn("w:abstractNumId")).get(qn("w:val"))
                                  == reuse_aid)), None)
            if reuse_num is not None:
                for p, il in g["paras"]:
                    _set_para_numid(p, reuse_num, il)
                    changed += 1
                continue
        # clone the abstractNum with a fresh id, detached from the source list's
        # identity (see _detach_clone_identity) so Word treats it as its own list
        new_aid = str(_next_int_id(root, "w:abstractNum", "w:abstractNumId",
                                   taken=new_abs_ids))
        new_abs_ids.append(new_aid)
        clone = copy.deepcopy(src_anum)
        clone.set(qn("w:abstractNumId"), new_aid)
        _detach_clone_identity(clone)
        _set_abstract_name(clone, CLAMP_CLONE_MARKER + src_aid)
        for lvl in clone.findall(qn("w:lvl")):
            lv = lvl.get(qn("w:ilvl"))
            try:
                if int(lv) in g["ilvls"]:
                    _neutralize_level_indent(lvl)
            except (TypeError, ValueError):
                continue
        if last_abstract is not None:
            last_abstract.addnext(clone)
        else:
            root.insert(0, clone)
        last_abstract = clone

        # new <w:num> -> cloned abstractNum
        new_numId = str(_next_int_id(root, "w:num", "w:numId", taken=new_num_ids))
        new_num_ids.append(new_numId)
        num_el = ordered_insert(root, etree.Element(qn("w:num")), order_for(root))
        num_el.set(qn("w:numId"), new_numId)
        a_el = etree.SubElement(num_el, qn("w:abstractNumId"))
        a_el.set(qn("w:val"), new_aid)

        for p, il in g["paras"]:
            _set_para_numid(p, new_numId, il)
            changed += 1

    if changed:
        tree.write(path, xml_declaration=True, encoding="UTF-8", standalone=True)
    return changed


# ---------------------------------------------------------------------------
# 方案C 阶段2：全角色 canonical 样式注入 + 指派 + 清直接覆盖
# ---------------------------------------------------------------------------
W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
CT_NS = "http://schemas.openxmlformats.org/package/2006/content-types"
REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
_OFFICE_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/"
_WML_CT = "application/vnd.openxmlformats-officedocument.wordprocessingml."



def _parse_fragment(xml):
    """把一个带 w: 前缀的 XML 片段解析成 lxml 元素（canonstyles 出的字符串）。"""
    wrapped = '<w:wrap xmlns:w="%s">%s</w:wrap>' % (W_NS, xml)
    return etree.fromstring(wrapped.encode("utf-8"))[0]


def _ensure_part(pkg_dir, fname, kind, empty_root):
    """确保 word/<fname> 存在，并已注册进 [Content_Types].xml 与 document.xml.rels。

    最小 docx（测试件、某些转换产物）可能根本没有 styles.xml / settings.xml，注入
    前得先把这个部件建出来并挂上关系，否则 Word 看不到它。"""
    path = os.path.join(pkg_dir, "word", fname)
    if os.path.exists(path):
        return path
    os.makedirs(os.path.dirname(path), exist_ok=True)
    with open(path, "w", encoding="utf-8") as f:
        f.write('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>%s' % empty_root)

    ct_path = os.path.join(pkg_dir, "[Content_Types].xml")
    if os.path.exists(ct_path):
        tree = parse_xml(ct_path)
        root = tree.getroot()
        part = "/word/%s" % fname
        if not any(ov.get("PartName") == part
                   for ov in root.findall("{%s}Override" % CT_NS)):
            ov = etree.SubElement(root, "{%s}Override" % CT_NS)
            ov.set("PartName", part)
            ov.set("ContentType", _WML_CT + kind + "+xml")
        tree.write(ct_path, xml_declaration=True, encoding="UTF-8", standalone=True)

    rels_path = os.path.join(pkg_dir, "word", "_rels", "document.xml.rels")
    if os.path.exists(rels_path):
        tree = parse_xml(rels_path)
        root = tree.getroot()
        rels = root.findall("{%s}Relationship" % REL_NS)
        if not any(r.get("Target") == fname for r in rels):
            maxid = 0
            for r in rels:
                rid = r.get("Id", "")
                if rid.startswith("rId") and rid[3:].isdigit():
                    maxid = max(maxid, int(rid[3:]))
            rel = etree.SubElement(root, "{%s}Relationship" % REL_NS)
            rel.set("Id", "rId%d" % (maxid + 1))
            rel.set("Type", _OFFICE_REL + kind)
            rel.set("Target", fname)
        tree.write(rels_path, xml_declaration=True, encoding="UTF-8", standalone=True)
    return path


def _merge_children(dst, src):
    """把 src 的子元素并进 dst：同名子元素**整体替换**，没有的**按 schema 顺序**插入。

    新子元素不能简单 append：文档自己的 `Normal` 可能只写了 `<w:sz>`，把 `<w:rFonts>`
    追加到它后面就违反 CT_RPr 的 sequence，Word 会拒绝打开（rFonts 必须排在 sz 前）。"""
    order = order_for(dst)
    for child in list(src):
        old = dst.find(child.tag)
        if old is not None:
            old.addprevious(child)
            dst.remove(old)
        elif order:
            ordered_insert(dst, child, order)
        else:
            dst.append(child)


def _patch_normal_and_defaults(styles_root, spec):
    """把 docDefaults 与 Normal(正文) 样式的字体/字号钉成 canonical 的**五号**。

    这是【文档网格】能不能是 15.6磅/41行 的前提（陷阱 #12）：Word 的"文档网格字体"
    就是 Normal 样式的字号，Normal=三号 时行高 21.75磅 顶破 15.6磅 的网格。正文内容
    不受影响——它由独立的 FGW正文(三号) 样式承载，且**每个有角色的段落都会被指派**
    canonical 样式，不再依赖 Normal 供给字号（这正是"指派要覆盖全部有角色段落、而不
    只是违规段落"的原因：否则改 Normal 会把原本合规、靠 Normal 拿到三号的段落悄悄
    变成五号，成为"改了却没提示"的盲区）。

    用**合并**而非整体替换：文档自己的 Normal/docDefaults 上的其它属性（语言、段落
    默认值等）保留不动，只钉字体字号。"""
    canon_def = _parse_fragment(canonstyles.doc_defaults_xml(spec))
    dd = styles_root.find(qn("w:docDefaults"))
    if dd is None:
        styles_root.insert(0, canon_def)
    else:
        src_rpr = canon_def.find(qn("w:rPrDefault") + "/" + qn("w:rPr"))
        rpr_default = _get_or_make(dd, "w:rPrDefault", before_tags=("w:pPrDefault",))
        dst_rpr = _get_or_make(rpr_default, "w:rPr")
        _merge_children(dst_rpr, src_rpr)

    canon_normal = _parse_fragment(canonstyles.normal_style_xml(spec))
    normal = None
    # 按 默认样式 → styleId=Normal → **名字**叫 Normal 三级找。名字这一级不能省：
    # 注入的 Normal 样式名就叫 "Normal"，文档里若已有一个名为 Normal、id 却不是
    # Normal 的样式（转换产物常见），漏掉它就会新增一个**重名**样式——Word 要求样式名
    # 唯一，重名即报"发现无法读取的内容"（陷阱 #18）。
    def _pick(pred):
        for st in styles_root.findall(qn("w:style")):
            if st.get(qn("w:type")) == "paragraph" and pred(st):
                return st
        return None

    def _name_of(st):
        nm = st.find(qn("w:name"))
        return nm.get(qn("w:val")) if nm is not None else None

    for pred in (lambda st: st.get(qn("w:default")) == "1",
                 lambda st: st.get(qn("w:styleId")) == canonstyles.NORMAL_STYLE_ID,
                 lambda st: _name_of(st) == _name_of(canon_normal)):
        normal = _pick(pred)
        if normal is not None:
            break
    if normal is None:
        styles_root.append(canon_normal)
    else:
        _merge_children(_get_or_make(normal, "w:rPr", before_tags=()),
                        canon_normal.find(qn("w:rPr")))


def _inject_canonical_styles(pkg_dir, spec, caption_num_ids=None):
    """注入全角色 canonical 命名样式（方案C §4：把 `_patch_toc_styles` 推广到每个角色）。

    同 styleId 的旧样式**整体替换**，因而幂等——对已处理过的文档重跑收敛到同一结果。
    文档原有的其它样式一概不动（段落靠 pStyle 改指 canonical 样式，而不是就地篡改
    共享样式：原地改会溢到未被指派的段落，就是 bug #17 那类事故）。

    样式定义本身来自 `scripts/lib/canonstyles.py`，与阶段0 经用户 Word 验收的参考件
    是**同一份字符串**。``caption_num_ids`` 把图/表标题的自动编号写进样式本身（套上
    样式即生成编号），故必须先注入编号定义、再注入样式。返回注入的样式数。"""
    path = _ensure_part(pkg_dir, "styles.xml", "styles",
                        '<w:styles xmlns:w="%s"/>' % W_NS)
    tree = parse_xml(path)
    root = tree.getroot()
    _patch_normal_and_defaults(root, spec)

    by_id = {st.get(qn("w:styleId")): st for st in root.findall(qn("w:style"))}
    # 文档里已被**别的 styleId** 占用的样式名。Word 要求样式名唯一，重名会让它在打开时
    # 报"发现无法读取的内容"。名字本身已带 FGW 前缀基本不会撞，这里再兜一道：真撞上就
    # 给我们的名字加后缀，绝不去改文档自己的样式。
    taken_names = {}
    for st in root.findall(qn("w:style")):
        nm = st.find(qn("w:name"))
        if nm is not None and nm.get(qn("w:val")):
            taken_names.setdefault(nm.get(qn("w:val")), set()).add(st.get(qn("w:styleId")))

    n = 0
    for d in canonstyles.canonical_style_defs(spec, caption_num_ids):
        el = _parse_fragment(d["xml"])
        nm = el.find(qn("w:name"))
        owners = taken_names.get(d["name"], set()) - {d["id"]}
        if owners and nm is not None:
            suffix = 2
            while ("%s%d" % (d["name"], suffix)) in taken_names:
                suffix += 1
            nm.set(qn("w:val"), "%s%d" % (d["name"], suffix))
        old = by_id.get(d["id"])
        if old is not None:
            old.addprevious(el)
            root.remove(old)
        else:
            root.append(el)
        n += 1
    tree.write(path, xml_declaration=True, encoding="UTF-8", standalone=True)
    return n


# 段落标记 rPr 上的**装饰性**属性：斜体、颜色、下划线、删除线、突出显示/底纹、各种
# 立体效果、隐藏、大写化、字符间距/缩放/升降。
#
# 为什么只对段落标记清、且清得比 run 狠：**自动编号的渲染完全取自段落标记的 rPr**。
# 标记上残留一个 `<w:i/>` 或 `<w:color w:val="FF0000"/>`，就会出现"序号是红色斜体、
# 标题文字却是黑色正体"（用户实测）。段落标记不是作者内容（它只是那个 ¶ 和自动编号的
# 载体），这些装饰留着有害无益；内容 run 上的同名属性属于作者的行内强调，**照旧保留**。
_MARK_DECORATION_TAGS = (
    "w:i", "w:iCs", "w:color", "w:u", "w:strike", "w:dstrike", "w:highlight",
    "w:shd", "w:bdr", "w:effect", "w:emboss", "w:imprint", "w:outline",
    "w:shadow", "w:vanish", "w:webHidden", "w:specVanish", "w:caps",
    "w:smallCaps", "w:position", "w:spacing", "w:w", "w:kern",
)


def _clear_run_props(p, clear_ea, clear_latin, clear_size, clear_bold=False,
                     overriding_char_styles=frozenset()):
    """清掉段落各 run 与段落标记 rPr 上、被 canonical 样式承载的字体/字号/加粗直接覆盖。

    严格-spec（§6 用户裁决）：canonical 值必须由注入的命名样式承载，"渲染对但用直接
    属性表达"不算合规。只清样式确实承载的键——颜色/下划线/上标之类样式没管的直接属性
    一律保留，指派样式不该顺手抹掉作者的行内强调。

    同时摘掉**会抢戏的字符样式引用**（`w:rStyle`）：内容 run 只摘那些确实设了字体/字号/
    加粗的字符样式，其余（超链接色、批注引用等）原样留着。字符样式是共享对象，绝不能就地
    改它的定义——那会溢到引用它的所有其它段落（#17 那类事故），所以摘引用而不是改样式。

    **段落标记（`pPr/rPr`）是例外，清得更狠**：它决定自动编号怎么渲染，所以标记上的加粗、
    全部装饰性属性（`_MARK_DECORATION_TAGS`：斜体/颜色/下划线/底纹…）以及**任何**字符样式
    引用一律清掉，保证序号与标题文字长得一样。"""
    # 只处理**已存在**的 rPr（别用 _run_rpr 顺手建空壳），清空后连壳一起删掉。
    owners = [(r, r.find(qn("w:rPr")), False) for r in _iter_runs(p)]
    pPr = get_pPr(p)
    owners.append((pPr, pPr.find(qn("w:rPr")), True))
    for owner, rpr, is_mark in owners:
        if rpr is None:
            continue
        if is_mark:
            # **段落标记的 rPr 决定自动编号怎么渲染**。用户实测：原本"表1"是加粗的，
            # 改成自动编号后编号仍然加粗、而标题文字不粗——因为加粗留在了段落标记上；
            # 后来又实测到"部分标题序号是红色斜体"，同一个根因（斜体/颜色留在标记上）。
            # 段落标记不是作者内容（它只是那个 ¶ 和自动编号的载体），所以这里把加粗与
            # **全部装饰性属性**一并清掉，让编号严格跟随段落样式：标题样式加粗→编号也粗，
            # 图表标题不加粗→编号也不粗，两边始终一致，且永远是正体、黑色。
            clear_bold = True
            for tag in _MARK_DECORATION_TAGS:
                for el in rpr.findall(qn(tag)):
                    rpr.remove(el)
        rs = rpr.find(qn("w:rStyle"))
        # 段落标记上的**任何**字符样式引用都摘掉（不只"设了字体/字号/加粗"的那些）：
        # 字符样式压过段落样式，标记上留着一个设了斜体/颜色的字符样式，编号照样会红、
        # 会斜。内容 run 仍只摘 `overriding_char_styles` 里那些——那才是作者的强调。
        if rs is not None and (is_mark
                               or rs.get(qn("w:val")) in overriding_char_styles):
            rpr.remove(rs)
        rf = rpr.find(qn("w:rFonts"))
        if rf is not None:
            attrs = []
            if clear_ea:
                attrs += ["w:eastAsia", "w:eastAsiaTheme"]
            if clear_latin:
                attrs += ["w:ascii", "w:asciiTheme", "w:hAnsi", "w:hAnsiTheme",
                          "w:cs", "w:cstheme"]
            for a in attrs:
                if rf.get(qn(a)) is not None:
                    del rf.attrib[qn(a)]
            if not rf.attrib:
                rpr.remove(rf)
        tags = []
        if clear_size:
            tags += ["w:sz", "w:szCs"]
        if clear_bold:
            # 连"显式取消加粗"(w:b val=0) 一起清掉——正是它让标题后半段不粗。
            tags += ["w:b", "w:bCs"]
        for tag in tags:
            el = rpr.find(qn(tag))
            if el is not None:
                rpr.remove(el)
        if len(rpr) == 0 and not rpr.attrib:
            owner.remove(rpr)


def _clear_ppr_governed(pPr, gov):
    """清掉段落 pPr 上被 canonical 样式承载的直接属性（缩进/行距/对齐/大纲级别）。"""
    if gov.get("ind"):
        for ind in pPr.findall(qn("w:ind")):
            pPr.remove(ind)
    sp = pPr.find(qn("w:spacing"))
    if sp is not None:
        attrs = []
        if gov.get("line"):
            attrs += ["w:line", "w:lineRule"]
        if gov.get("space_before_after"):
            attrs += ["w:before", "w:after", "w:beforeLines", "w:afterLines",
                      "w:beforeAutospacing", "w:afterAutospacing"]
        for a in attrs:
            if sp.get(qn(a)) is not None:
                del sp.attrib[qn(a)]
        if not sp.attrib:
            pPr.remove(sp)
    if gov.get("jc"):
        for jc in pPr.findall(qn("w:jc")):
            pPr.remove(jc)
    if gov.get("outline"):
        for ol in pPr.findall(qn("w:outlineLvl")):
            pPr.remove(ol)


# 段落里"没有文字、但有东西"的载体：图片/形状（drawing、VML pict）、嵌入对象、
# 以及 mc:AlternateContent（里面裹着 drawing 的兼容写法）。
_EMBEDDED_TAGS = (
    "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}drawing",
    "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}pict",
    "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}object",
    "{http://schemas.openxmlformats.org/markup-compatibility/2006}AlternateContent",
)


def _is_object_only(p):
    """段落是不是**纯粹的图片/形状/文本框容器**——含嵌入对象、且自身没有可见文字。

    这类段落**一律不指派 canonical 样式**。原因：所有正文类 canonical 样式都带
    **固定行距 28 磅**（`lineRule=exact`），而固定行距**不会随内容长高**——套到一张
    图上，图就被裁成一行高，只露出底部一条（用户实测："图片隐于文字下方，只有底部
    一行的宽度露出来"）。图片段落本来也没有"字体字号行距"可言，放过它才对。

    "自身文字"不算文本框里的（`w:txbxContent` 属于另一个逻辑段落，`_iter_runs` 也
    跳过它）——所以一个只装了文本框的段落同样算 object-only。反之，正文里插了个小
    图标、但确实有文字的段落**仍然要**规范化。"""
    if not any(el.tag in _EMBEDDED_TAGS for el in p.iter()):
        return False
    own_text = "".join((t.text or "") for t in p.iter(qn("w:t"))
                       if not in_textbox(t, stop_at=p))
    return not own_text.strip()


def _blank_style(rec):
    """空行的 canonical 样式：**空行统一套正文样式**（仿宋三号），封面也不例外。

    空行没有可判的格式（`paragraph_role` 对空行返回 None，也不该为它挂批注），但它
    仍然占版面高度：不指派样式的话它跟随 `Normal`，而 `Normal` 已被钉成五号（文档
    网格的前提，陷阱 #12），空行会莫名其妙变矮、和正文行距不一致。

    **封面空行原本是例外**（"版式留白"），2026-08 第六轮验收由用户推翻——"封面的空行
    也应用正文的样式"。代价是封面留白会从五号行高变成正文的固定行距 28 磅、整体变高，
    这是用户知情后的裁决，别再以"封面是版式"为由改回去。

    剩下两处例外：① 落在 **TOC 域/目录内容控件跨度内**的空行，动它有破坏域的风险；
    ② 表格里的空行跟随单元格内容的表格样式，与同格文字保持一致更合理。（图片/文本框
    段落由 `_is_object_only` 更早滤掉——正文样式的固定行距会把图裁成一行，陷阱 #17。）

    例外①只认**域跨度**（`toc_in_field`），不认"套着目录样式"：目录后面紧跟的那几个
    空行往往顺手继承了 `TOC3`/`Contents 3` 样式，它们并不在域里、刷新目录不会碰它们，
    却因为旧的 `is_toc` 判据被一起放过——用户第六轮报的"空行没有应用正文样式"就是这类
    （旧结构 json 没有 `toc_in_field` 字段时回退到 `is_toc`，保持旧行为）。"""
    if not rec.get("is_blank"):
        return None
    if rec.get("toc_in_field", rec.get("is_toc")):
        return None
    if rec.get("auto_num"):
        # 文字为空、却挂着**自动编号**的段落不是空行——Word 会在这里渲染出一个编号。
        # 套上正文样式会把那个编号变成正文格式，还会因为"有编号 + 样式承载缩进"被送进
        # 克隆钳、分到一条私有列表上从头数。原样不动。
        return None
    return STYLE_ID_BY_ROLE["table_body" if rec.get("in_table") else "body"]


def _heading_num_ref(rec, level, heading_num_id, current):
    """一个**原本就自动编号**的标题该挂到哪条编号上：返回 ((numId, ilvl), 是否取消编号)。

    三条分支：

      * **惯例不编号的章节**（摘要/前言/结论/参考文献/附录/致谢…）——取消自动编号。
        判定层（`checks.continuity`）本来就不给它们发号、也不让它们占同级序号，编号层
        却照发不误的话，就会出现"三、结论"、而且它后面的同级标题全被顶掉一位。
      * **序号已经写在文字里**（`num_raw`）——取消自动编号。自动编号叠在手写序号上会
        渲染成"一、一、绪论"。**注意这只是本阶段的保守动作**：2026-08 起绝大多数这类
        标题会由 `autonumber_heading` 那条 fix 在稍后的 fix 循环里把手写序号删掉、再挂
        回 canonical 编号（`_apply_autonumber_heading`），此处写下的 `numId=0` 会被它
        覆盖。留着这一步是**兜底**：转换若失败（比如文字里的序号删不掉），段落就停在
        "手写序号 + 无自动编号"这个自洽状态上，而不是两个编号叠着。
      * 其余——改指注入的 canonical 四级标题列表，`ilvl = 级别-1`。四级同处**一条**多级
        列表是"二级在其一级下从头数"能成立的前提（见
        `canonstyles.heading_numbering_def`）。

    ``current`` 是该段**原有**的 (numId, ilvl)：spec 关掉标题自动编号时原样返回它——
    指派 canonical 样式会丢掉原样式携带的 numPr，不保号等于把"一、"改没了（改原文）。"""
    if is_unnumbered_section(rec.get("text") or "", rec.get("num_raw")):
        return (None, None), True
    if rec.get("num_raw"):
        return (None, None), True
    if not heading_num_id:
        return current, False           # spec 没给标题编号定义：保号、原样不动
    return (heading_num_id, min(max(level, 1), 4) - 1), False


def _assignable(rec):
    """该段的角色是否**可信到可以承载 canonical 样式**（陷阱#5 安全阀）。

    仅凭"图/表+数字"形状认出、没有题注样式撑腰的图表标题（`caption.source == "pattern"`）
    不指派：给它套上"图标题"样式，下一轮它就凭样式变成"已确认"，从而绕过安全阀被自动
    改编号——而它很可能只是一句以"图3 显示了…"开头的正文。这类段落仍走直接属性路径拿到
    格式修复，只是不进样式体系。pattern 标题（heading）已在 `paragraph_role` 里降级成
    正文角色，不必在这里再判一次。"""
    cap = rec.get("caption")
    return not (cap and cap.get("source") == "pattern")


def _set_pstyle(p, style_id):
    """给段落指派样式（w:pStyle 必须是 pPr 的第一个子元素）。"""
    pPr = get_pPr(p)
    for old in pPr.findall(qn("w:pStyle")):
        pPr.remove(old)
    el = etree.Element(qn("w:pStyle"))
    el.set(qn("w:val"), style_id)
    pPr.insert(0, el)


def _assign_canonical_style(p, style_id, gov, num_ref=(None, None),
                            overriding_char_styles=frozenset()):
    """给段落指派 canonical 样式，并清掉该样式承载的直接覆盖。

    ``num_ref`` 是**指派前**解析出的有效 (numId, ilvl)。自动编号常常挂在原样式的
    numPr 上，改指 canonical 样式会把编号一起弄丢（"一、"消失＝改了原文），所以这里
    把它**钉成段落的直接 numPr** 保号。编号层带来的缩进随后由甲法克隆钳中和。"""
    _set_pstyle(p, style_id)
    _clear_ppr_governed(get_pPr(p), gov)
    _clear_run_props(p, gov.get("fonts", False), gov.get("western", False),
                     gov.get("size", False), gov.get("bold", False),
                     overriding_char_styles)
    nid, ilvl = num_ref
    if nid:
        _set_para_numid(p, nid, ilvl or 0)


# 一个 fix 的 set_* 键分别由哪项样式承载 —— 段落已指派 canonical 样式时，这些键不再
# 写成直接属性（样式已经供给了同一个值）。
_FIX_KEY_GOVERNOR = {
    "set_east_asia": "fonts", "set_ascii": "western", "set_size_hp": "size",
    "set_bold": "bold",
    "set_line_exact": "line", "set_line_rule": "line",
    "clear_space_before_after": "space_before_after",
    "set_first_line_chars": "ind", "set_left_chars": "ind",
    "clear_left_indent": "ind", "clear_right_indent": "ind",
    "set_jc": "jc",
}


def _governed(gov, key):
    return bool(gov.get(_FIX_KEY_GOVERNOR.get(key, ""), False))


# ---------------------------------------------------------------------------
# 文档网格 / 表格默认值
# ---------------------------------------------------------------------------
def _apply_document_grid(pkg_dir, spec):
    """把 spec.document_grid 写进 settings.xml（默认制表位/绘图网格/compat 块/语言）。

    行网格 15.6磅/41行 不是只靠 sectPr 的 docGrid——**compat 块（尤其 useFELayout +
    compatibilityMode=15）才是开关**，缺它 Word 按旧版式把行距顶到 21.75磅/29行
    （陷阱 #12，用户实测）。返回写入的子元素数。"""
    children = canonstyles.settings_children_xml(spec)
    if not children:
        return 0
    path = _ensure_part(pkg_dir, "settings.xml", "settings",
                        '<w:settings xmlns:w="%s"/>' % W_NS)
    tree = parse_xml(path)
    root = tree.getroot()
    for tag, xml in children:
        el = _parse_fragment(xml)
        old = root.find(qn("w:" + tag))
        if old is not None:
            old.addprevious(el)
            root.remove(old)
        else:
            ordered_insert(root, el, order_for(root))
    tree.write(path, xml_declaration=True, encoding="UTF-8", standalone=True)
    return len(children)


def _apply_doc_grid_to_sections(doc_root, spec):
    """每个 sectPr 设行网格（docGrid，必须是 sectPr 的最后一个子元素）。"""
    attrs = canonstyles.doc_grid_attrs(spec)
    if not attrs:
        return 0
    n = 0
    for sect in doc_root.iter(qn("w:sectPr")):
        grid = sect.find(qn("w:docGrid"))
        if grid is None:
            # sectPr 里 docGrid 后面还可能有 printerSettings/sectPrChange，按序插
            grid = ordered_insert(sect, etree.Element(qn("w:docGrid")),
                                  order_for(sect))
        for k, v in attrs.items():
            grid.set(qn("w:" + k), v)
        n += 1
    return n


def _apply_table_defaults(doc_root, spec):
    """表格：左缩进 tblInd、表级默认单元格边距 tblCellMar、单元格垂直对齐 vAlign。

    `tblInd`（表格整体左缩进）与 `tblCellMar`（单元格内边距）是两回事，别混
    （陷阱 #12）。返回处理的表格数。"""
    if not (spec.get("table_defaults") or {}):
        return 0
    ind_xml = canonstyles.table_ind_xml(spec)
    mar_xml = canonstyles.cell_margins_xml(spec)
    valign_xml = canonstyles.cell_valign_xml(spec)
    n = 0
    for tbl in doc_root.iter(qn("w:tbl")):
        tblPr = _get_or_make(tbl, "w:tblPr", before_tags=("w:tblGrid", "w:tr"))
        if ind_xml:
            for old in tblPr.findall(qn("w:tblInd")):
                tblPr.remove(old)
            ordered_insert(tblPr, _parse_fragment(ind_xml), order_for(tblPr))
        if mar_xml:
            for old in tblPr.findall(qn("w:tblCellMar")):
                tblPr.remove(old)
            ordered_insert(tblPr, _parse_fragment(mar_xml), order_for(tblPr))
        if valign_xml:
            for tc in tbl.iter(qn("w:tc")):
                tcPr = _get_or_make(tc, "w:tcPr", before_tags=("w:p", "w:tbl"))
                for old in tcPr.findall(qn("w:vAlign")):
                    tcPr.remove(old)
                ordered_insert(tcPr, _parse_fragment(valign_xml), order_for(tcPr))
        n += 1
    return n




# ---------------------------------------------------------------------------
# 图/表标题：Word 自动编号（取代静态"图N/表N"）
# ---------------------------------------------------------------------------
# 静态编号后要连同紧跟的分隔空白一起删掉——编号改由 Word 生成，分隔符由编号定义的
# suff=tab 提供，残留的空格会变成"图1<制表符> 说明"里多出来的那个空格。
STRIP_CAPTION_LABEL = re.compile(r"^\s*(?:图|表)\s*[0-9]+(?:[-\.–][0-9]+)?[ \t　]*")
STRIP_CAPTION_LABEL_RESIDUE = re.compile(r"^\s*(?:图|表)[ \t　]*")


def _ensure_injected_numbering(pkg_dir, spec, cache):
    """确保 numbering.xml 里有本工具的 canonical 编号定义（图%1 / 表%1 / 四级标题），
    返回 {kind: numId}，kind ∈ {"figure","table","heading"}。

    幂等靠 `<w:name>` 标签认领已注入的定义——否则每跑一次就多几条编号定义。级别缩进
    已中和（编号层压过样式层，不中和会把 canonical 样式的缩进盖掉，陷阱 #11）。"""
    if cache:
        return cache
    defs = canonstyles.injected_numbering_defs(spec)
    if not defs:
        return {}
    path = _ensure_part(pkg_dir, "numbering.xml", "numbering",
                        '<w:numbering xmlns:w="%s"/>' % W_NS)
    tree = parse_xml(path)
    root = tree.getroot()

    # 已注入过的：abstractNum 的 w:name 命中标记
    by_marker = {}
    for anum in root.findall(qn("w:abstractNum")):
        nm = anum.find(qn("w:name"))
        val = nm.get(qn("w:val")) if nm is not None else None
        if val:
            by_marker[val] = anum.get(qn("w:abstractNumId"))
    abs2num = {}
    for num in root.findall(qn("w:num")):
        a = num.find(qn("w:abstractNumId"))
        if a is not None:
            abs2num.setdefault(a.get(qn("w:val")), num.get(qn("w:numId")))

    out, changed = {}, False
    new_abs, new_nums = [], []
    last_abstract = None
    for anum in root.findall(qn("w:abstractNum")):
        last_abstract = anum
    for d in defs:
        aid = by_marker.get(d["marker"])
        if aid is not None and aid in abs2num:
            out[d["kind"]] = abs2num[aid]
            continue
        aid = str(_next_int_id(root, "w:abstractNum", "w:abstractNumId", taken=new_abs))
        new_abs.append(aid)
        anum = _parse_fragment(
            '<w:abstractNum w:abstractNumId="%s"><w:multiLevelType w:val="%s"/>'
            '<w:name w:val="%s"/>%s</w:abstractNum>'
            % (aid, d.get("multi_level_type", "singleLevel"), d["marker"], d["lvl_xml"]))
        # schema 要求所有 abstractNum 排在所有 num 之前
        if last_abstract is not None:
            last_abstract.addnext(anum)
        else:
            root.insert(0, anum)
        last_abstract = anum
        nid = str(_next_int_id(root, "w:num", "w:numId", taken=new_nums))
        new_nums.append(nid)
        num_el = ordered_insert(root, etree.Element(qn("w:num")), order_for(root))
        num_el.set(qn("w:numId"), nid)
        a_el = etree.SubElement(num_el, qn("w:abstractNumId"))
        a_el.set(qn("w:val"), aid)
        out[d["kind"]] = nid
        changed = True
    if changed:
        tree.write(path, xml_declaration=True, encoding="UTF-8", standalone=True)
    cache.update(out)
    return out


def _apply_autonumber_caption(p):
    """把图/表标题交给 Word 自动编号：删掉文字里的静态"图N/表N"，并**清掉段落上的直接
    编号覆盖**，让 canonical 题注样式携带的编号生效。

    编号挂在**样式**上而不是段落上（`_inject_canonical_styles` 写进样式的 `numPr`）：
    这样用户在 Word 里插入新图、给标题套上"图标题"样式就直接有编号。段落若残留旧的
    直接 `numPr`，会**压过**样式的编号（直接层优先），所以这里要删掉。

    删静态编号也是必须的——不删就会出现"图1<制表符>图1 系统架构"（Word 生成的编号叠在
    原文字上）。这属于既有的图表重编号内容编辑范畴：渲染出来的编号仍在，只是改由 Word
    维护，插入/删除图表后不会再失序。"""
    pPr = get_pPr(p)
    for npr in pPr.findall(qn("w:numPr")):
        pPr.remove(npr)
    _replace_leading(p, STRIP_CAPTION_LABEL, "",
                     residue_re=STRIP_CAPTION_LABEL_RESIDUE)
    return True


def _drop_leading_separator_tabs(p):
    """删掉段落开头那些"孤儿制表符"——序号被删后残留的分隔符。

    两种形态都要清：① 制表符排在第一个有文字的 `w:t` **之前**（原来它跟在序号后面）；
    ② `_replace_leading` 把全部文字并进了第一个 `w:t`、后面的 `w:t` 全空，此时**任何**
    制表符都是残渣（它们原来的上下文已经不存在了）——不清就会渲染成"一、<制表符><制表符>
    绪论"（编号自己的 `suff=tab` 再加一个）。形态②之外不动标题内部的制表符。"""
    items = para_inline_items(p)
    first_text = next((el for el in items
                       if el.tag == qn("w:t") and (el.text or "").strip()), None)
    if first_text is None:
        drop = [el for el in items if el.tag == qn("w:tab")]
    else:
        after = items[items.index(first_text) + 1:]
        collapsed = all(not (el.text or "") for el in after if el.tag == qn("w:t"))
        drop = []
        for el in items:
            if el.tag != qn("w:tab"):
                continue
            before_text = items.index(el) < items.index(first_text)
            if before_text or collapsed:
                drop.append(el)
    for el in drop:
        parent = el.getparent()
        if parent is not None:
            parent.remove(el)
    return len(drop)


def _apply_autonumber_heading(p, level, num_id, strip):
    """把标题交给 Word 自动编号：删掉文字里的手写序号，并把段落挂到注入的四级标题列表。

    **两件事必须同生共死**（用户："要注意去掉原本的手写编号，这里逻辑需要缜密一点"）：
    只删序号不挂编号 = 把内容删没了；只挂编号不删序号 = 渲染成"一、一、绪论"。所以
    这里先试删、删不掉就整条放弃（返回 False，调用方记 problem 并保持原样——段落此前
    在样式指派阶段已按"手写序号"被写了 `numId=0`，放弃后正好维持旧行为）。

    编号挂在**段落**上而不是标题样式上（与图/表标题相反）：摘要/结论这类"惯例不编号的
    章节"与普通标题共用同一个 canonical 标题样式，编号若挂样式上就没法只让它们不编号
    （陷阱 #19 反例②）。`ilvl = 级别-1`，四级同处一条列表才有逐级归零。"""
    if not num_id:
        return False
    if strip and not _replace_leading(p, STRIP_HEADING, ""):
        return False
    if strip:
        # 序号后面的分隔符是"序号的一部分"，跟着序号一起走，否则会与编号自带的
        # `suff=tab` 叠出两段空白：`_replace_leading` 只在**序号被整段吃掉**时才顺手
        # 削空白（这里替换成的是空串，它会把原分隔空白当正文留下），`w:tab` 元素更是
        # 完全不归它管。两样都要单独清。
        _strip_para_ws(p, "leading")
        _drop_leading_separator_tabs(p)
    _set_para_numid(p, num_id, min(max(level, 1), 4) - 1)
    return True


def _set_update_fields(pkg_dir):
    """Set settings.xml <w:updateFields w:val="true"/> so Word refreshes the TOC
    (renumbered headings + new page numbers) when the document is opened.

    Trade-off (chosen deliberately): Word shows the "该文档包含的域可能引用了其他
    文件。是否更新…" prompt once per open, and clicking 是 rebuilds the TOC to match
    the corrected body. In Word there is no document setting that auto-refreshes
    the TOC WITHOUT this prompt — a per-field w:dirty mark triggers the very same
    dialog — so the only prompt-free alternative is to leave the TOC stale until
    the user presses Ctrl+A then F9. We take the native, exact refresh here.
    If the prompt appears for OTHER reasons (external INCLUDE/LINK/DDE fields,
    external relationships, OLE), run scripts/diagnose_fields.py on the output."""
    path = os.path.join(pkg_dir, "word", "settings.xml")
    if not os.path.exists(path):
        return
    tree = parse_xml(path)
    root = tree.getroot()
    uf = root.find(qn("w:updateFields"))
    if uf is None:
        uf = ordered_insert(root, etree.Element(qn("w:updateFields")),
                            order_for(root))
    uf.set(qn("w:val"), "true")
    tree.write(path, xml_declaration=True, encoding="UTF-8", standalone=True)


# ---------------------------------------------------------------------------
def main():
    if len(sys.argv) < 2:
        print(json.dumps({"status": "error", "error": "usage: 40_apply_fixes.py <workdir>"}))
        sys.exit(1)
    workdir = sys.argv[1]
    meta = json.load(open(os.path.join(workdir, "meta.json"), encoding="utf-8"))
    fixes = json.load(open(os.path.join(workdir, "fixes.json"), encoding="utf-8"))

    working = meta["working_docx"]
    out_pkg = os.path.join(workdir, "out_pkg")
    if os.path.isdir(out_pkg):
        import shutil
        shutil.rmtree(out_pkg)
    unzip_docx(working, out_pkg)

    doc_path = os.path.join(out_pkg, "word", "document.xml")
    tree = parse_xml(doc_path)
    root = tree.getroot()

    # index -> paragraph element (same iterator as extraction)
    para_by_idx = {i: p for i, p in iter_body_paragraphs(root)}

    applied = {"format": 0, "renumber_caption": 0, "renumber_heading": 0,
               "section": 0, "hint": 0, "comments": 0, "skipped": 0}
    try:
        spec = load_default_spec()
    except (OSError, ValueError):
        spec = {}

    # ---- 方案C 阶段2：先注入 canonical 样式 + 文档网格，再指派 ----------------
    # 注入必须在算 char_unit_hp 之前：docDefaults 被钉成五号后，"N 字符"的尺子才是
    # 21 半点，目录制表位/左缩进伴随值都按它换算（陷阱 #12）。
    caption_nums, heading_num_id = {}, None
    if spec:
        # 图/表标题的自动编号定义必须**先**注入：它的 numId 要写进 canonical 题注样式，
        # 这样用户在 Word 里新插一张图、给标题套上"图标题"样式就直接生成编号（只挂在
        # 段落上的话新段落没有编号——用户实测反馈）。标题的四级编号定义同批注入。
        injected_nums = _ensure_injected_numbering(out_pkg, spec, {})
        caption_nums = {k: v for k, v in injected_nums.items()
                        if k in ("figure", "table")}
        heading_num_id = injected_nums.get("heading")
        applied["caption_numbering"] = len(caption_nums)
        applied["canonical_styles"] = _inject_canonical_styles(out_pkg, spec, caption_nums)
        applied["grid_settings"] = _apply_document_grid(out_pkg, spec)
        applied["doc_grid_sections"] = _apply_doc_grid_to_sections(root, spec)
        applied["tables_normalized"] = _apply_table_defaults(root, spec)

    # Character-unit size for converting *Chars indents to absolute twips —
    # the DOCUMENT default font size, matching how Word measures "N 字符".
    char_unit_hp = _default_char_unit_hp(out_pkg)

    # Resolver + numbering map for the 甲法 numbering-layer clamp: an indent fix
    # on an AUTO-NUMBERED paragraph is written char-only, and the inherited
    # hanging is removed at the numbering layer (a cloned abstractNum) instead of
    # being out-muscled by an absolute companion. Parsed AFTER injection so the
    # resolver knows the canonical styles too.
    styles_path = os.path.join(out_pkg, "word", "styles.xml")
    styles_root = parse_xml(styles_path).getroot() if os.path.exists(styles_path) else None
    resolver = StyleResolver(styles_root)
    numbering_path = os.path.join(out_pkg, "word", "numbering.xml")
    numbering_root = (parse_xml(numbering_path).getroot()
                      if os.path.exists(numbering_path) else None)
    numbering_levels = load_numbering_levels(numbering_root)
    clamp_targets = []

    # ---- 指派 pStyle：每个【有角色】的段落都套 canonical 样式 ------------------
    # 为什么是"全部有角色的段落"而不只是违规段落：① §6 用户裁决——"渲染对但用直接
    # 属性表达"不算合规，canonical 值必须由命名样式承载；② 上面把 Normal 钉成五号
    # （网格前提），原本靠 Normal 拿到三号的**合规**段落若不指派样式就会被悄悄改小，
    # 那才是真正的"改了却没提示"盲区。对合规段落而言指派是**渲染中性**的：样式承载
    # 的值本来就等于它的有效值，样式没承载的属性（如密级行的对齐）原样保留。
    # 安全阀（陷阱#5）在 `checks.paragraph_role` 里：pattern 标题不算 heading 角色。
    governs = ({d["id"]: d["governs"] for d in canonstyles.canonical_style_defs(spec)}
               if spec else {})
    canon_caption_style_ids = frozenset(
        STYLE_ID_BY_ROLE[r] for r in canonstyles.CAPTION_ROLE_BY_KIND.values())
    # 设了字体/字号/加粗的**字符样式**：它们压过段落样式，指派时要把 run 上的引用摘掉，
    # 否则同一段里带这种字符样式的 run 不跟随 canonical 样式（半粗半细的根因）。
    overriding_char_styles = char_styles_overriding(styles_root)
    assigned = {}
    structure_path = os.path.join(workdir, "structure.json")
    if spec and os.path.exists(structure_path):
        try:
            with open(structure_path, encoding="utf-8") as f:
                records = json.load(f)["records"]
        except (OSError, ValueError, KeyError):
            records = []
        for rec in records:
            p = para_by_idx.get(rec.get("i"))
            if p is None or not _assignable(rec) or _is_object_only(p):
                continue
            role = paragraph_role(rec, spec) or ""
            sid = STYLE_ID_BY_ROLE.get(role) or _blank_style(rec)
            if sid is None:
                continue
            gov = governs.get(sid, {})
            # 指派前解析编号（改指 canonical 样式会丢掉原样式携带的 numPr，"一、"会消失）。
            # 图/表标题除外：它们的编号由 canonical 题注样式承载，钉住旧编号反而会压过
            # 样式的编号（直接层优先），所以不保号——旧编号本来就要被换掉。
            num_ref = ((None, None) if sid in canon_caption_style_ids
                       else _para_num_ref(p, resolver, numbering_levels))
            # 标题的自动编号统一改挂注入的 canonical 四级列表（逐级归零的前提），
            # 或按 `_heading_num_ref` 的判断整段取消编号。
            suppress_num, on_canonical_num = False, False
            if role.startswith("heading") and num_ref[0]:
                num_ref, suppress_num = _heading_num_ref(
                    rec, int(role[-1]), heading_num_id, num_ref)
                on_canonical_num = bool(num_ref[0]) and num_ref[0] == heading_num_id
                if on_canonical_num:
                    applied["heading_numbering"] = applied.get("heading_numbering", 0) + 1
            _assign_canonical_style(p, sid, gov, num_ref, overriding_char_styles)
            if suppress_num:
                _suppress_auto_numbering(p)
                applied["heading_num_suppressed"] = applied.get(
                    "heading_num_suppressed", 0) + 1
            assigned[rec["i"]] = sid
            # 挂上 canonical 标题编号的段落不进克隆钳：那条定义的级别缩进本来就中和过。
            if gov.get("ind") and num_ref[0] and not on_canonical_num:
                clamp_targets.append(p)
    applied["styles_assigned"] = len(assigned)

    cw = CommentWriter(out_pkg, author="XAgent")
    problems = []
    # idx -> [paragraph_el, [rule_texts]] in first-seen order. Multiple fixes on
    # ONE paragraph (e.g. a font fix + a renumber) are merged into a single
    # XAgent comment at the end, so the document isn't littered with several
    # overlapping comment ranges on the same line.
    pending_comments = {}
    # 需要在**所有 fix 应用完之后**重新分隔"序号 + 制表符 + 标题文字"的段落。
    number_tab_targets = set()
    heading_tab = ((spec.get("heading_numbering") or {}).get("suffix")
                   or "tab") == "tab" if spec else False

    for fix in fixes:
        op = fix.get("op")
        if op == "section":
            applied["section"] += _apply_section(root, fix.get("set_pgmar", {}))
            # section rule has comment=false by spec; nothing to attach
            continue

        idx = fix.get("para_index")
        p = para_by_idx.get(idx)
        if p is None:
            problems.append({"para_index": idx, "op": op, "reason": "paragraph_not_found"})
            applied["skipped"] += 1
            continue

        ok = True
        if op == "format":
            # 段落已指派 canonical 样式时，样式承载的属性**不再写成直接覆盖**
            # （严格-spec §2.2/§6：canonical 值由命名样式供给，直接层要清空）；
            # 样式没承载的属性仍按老路写直接属性（如目录条目——它没有 canonical
            # 样式，缩进由文档自己的 TOC 样式承载）。
            gov = governs.get(assigned.get(idx), {})
            _apply_run_props(
                p,
                None if _governed(gov, "set_east_asia") else fix.get("set_east_asia"),
                None if _governed(gov, "set_ascii") else fix.get("set_ascii"),
                None if _governed(gov, "set_size_hp") else fix.get("set_size_hp"),
                None if _governed(gov, "set_bold") else fix.get("set_bold"))
            pPr = get_pPr(p)
            if (fix.get("set_line_exact") is not None
                    and not _governed(gov, "set_line_exact")):
                _set_line_exact(pPr, fix["set_line_exact"],
                                fix.get("set_line_rule") or "exact")
            if (fix.get("clear_space_before_after")
                    and not _governed(gov, "clear_space_before_after")):
                _clear_space_before_after(pPr)
            if (not _governed(gov, "set_first_line_chars")
                    and (fix.get("set_first_line_chars") is not None
                         or fix.get("clear_left_indent") or fix.get("clear_right_indent")
                         or fix.get("set_left_chars") is not None)):
                _set_first_line_and_clear_left(
                    pPr, fix.get("set_first_line_chars"),
                    bool(fix.get("clear_left_indent")),
                    bool(fix.get("clear_right_indent")),
                    fix.get("set_left_chars"))
                if _para_num_ref(p, resolver, numbering_levels)[0]:
                    clamp_targets.append(p)
            if fix.get("set_jc") is not None and not _governed(gov, "set_jc"):
                _set_jc(pPr, fix["set_jc"])
            if fix.get("strip_text"):
                _strip_para_ws(p, fix["strip_text"])
            if fix.get("set_number_tab"):
                # 实际改动推迟到本轮所有 fix 之后：同一段若还有 renumber_heading，
                # 归位序号会把文字并回一个 w:t、把制表符甩到文字后面。
                number_tab_targets.add(idx)
            applied["format"] += 1
        elif op == "renumber_caption":
            ok = _apply_renumber_caption(p, fix)
            applied["renumber_caption"] += 1 if ok else 0
        elif op == "autonumber_caption":
            if not caption_nums.get(fix.get("kind")):
                ok = False
            else:
                ok = _apply_autonumber_caption(p)
                applied["autonumber_caption"] = applied.get("autonumber_caption", 0) + 1
                # 编号来自 canonical 题注样式（缩进已中和），无需再克隆钳
                clamp_targets = [t for t in clamp_targets if t is not p]
        elif op == "autonumber_heading":
            # 删手写序号 + 挂 canonical 标题编号，两件事同生共死（见函数注释）。
            # 没有注入成功的标题编号定义时整条放弃：宁可留着手写序号，也不能把序号
            # 删掉却没有编号顶上。
            ok = _apply_autonumber_heading(p, fix.get("level") or 1,
                                           heading_num_id, fix.get("strip"))
            if ok:
                applied["autonumber_heading"] = applied.get("autonumber_heading", 0) + 1
                applied["heading_numbering"] = applied.get("heading_numbering", 0) + 1
                # 挂上的是 canonical 那条列表，级别缩进本来就中和过，不必进克隆钳
                clamp_targets = [t for t in clamp_targets if t is not p]
        elif op == "renumber_heading":
            ok = _apply_renumber_heading(p, fix)
            applied["renumber_heading"] += 1 if ok else 0
            # 归位序号会把整段文字并回第一个 w:t，原有的制表符会被落在文字后面——
            # 无论判定层有没有报"缺制表符"，重排过序号的段落都要重新分隔一次。
            if ok:
                number_tab_targets.add(idx)
        elif op == "hint":
            applied["hint"] += 1
        else:
            problems.append({"para_index": idx, "op": op, "reason": "unknown_op"})
            applied["skipped"] += 1
            continue

        if not ok:
            problems.append({"para_index": idx, "op": op, "reason": "text_edit_failed"})

        # queue this fix's rule text for the paragraph's single merged comment
        if fix.get("comment") and fix.get("rule_text"):
            slot = pending_comments.get(idx)
            if slot is None:
                slot = pending_comments[idx] = [p, []]
            slot[1].append(fix["rule_text"])

    # 手写序号与标题文字之间补制表符——放在所有 fix 之后（renumber 会重排文字）、
    # 批注之前（批注区间要把最终内容整个包住）。
    if heading_tab:
        for idx in sorted(number_tab_targets):
            p_el = para_by_idx.get(idx)
            if p_el is not None and _heading_number_tab(p_el):
                applied["heading_number_tab"] = applied.get("heading_number_tab", 0) + 1

    # one XAgent comment per paragraph, joining all its rule texts
    for _idx, (p_el, texts) in pending_comments.items():
        cw.add(p_el, "；".join(texts))
        applied["comments"] += 1

    cw.flush()

    # Patch the TOC entry styles so a refreshed TOC keeps the spec font/size,
    # per-level indent and tab stops (the direct formatting applied above is
    # otherwise wiped when Word rebuilds the field on open).
    applied["toc_styles"] = (_patch_toc_styles(out_pkg, spec.get("toc"), char_unit_hp)
                             if spec else 0)

    # Un-hide auto-generated caption numbers (图N/表N) obscured by a black
    # shading / zero size in the caption numbering definition.
    applied["caption_num_unhidden"] = _clean_caption_numbering(out_pkg)

    # 方案C 甲法：钳住编号层。For every auto-numbered paragraph that received an
    # indent fix (applied char-only above), clone its abstractNum and neutralize
    # the inherited hanging so the char-unit first-line indent renders as "N 字符"
    # without an absolute companion — clone (not in-place) so same-level siblings
    # are untouched (§2.5/#17). Mutates paragraph numPr in `root` (written below)
    # and numbering.xml in place.
    applied["numbering_clamped"] = _clamp_numbering_indent(
        out_pkg, clamp_targets, resolver, numbering_levels)

    # Refresh the TOC on open (renumbered headings + page numbers) via the global
    # updateFields flag. Word prompts once on open; clicking 是 rebuilds the TOC
    # to match the corrected body (see _set_update_fields for the trade-off).
    _set_update_fields(out_pkg)

    tree.write(doc_path, xml_declaration=True, encoding="UTF-8", standalone=True)

    out_docx = os.path.join(workdir, "formatted.docx")
    rezip_docx(out_pkg, out_docx)

    report = {"status": "ok", "formatted_docx": out_docx,
              "applied": applied, "problems": problems}
    json.dump(report, open(os.path.join(workdir, "apply_report.json"), "w",
                           encoding="utf-8"), ensure_ascii=False, indent=1)
    print(json.dumps({"status": "ok", "formatted_docx": out_docx,
                      "applied": applied, "n_problems": len(problems)},
                     ensure_ascii=False))


if __name__ == "__main__":
    main()
