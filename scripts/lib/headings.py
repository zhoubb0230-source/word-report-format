# -*- coding: utf-8 -*-
"""
headings.py — shared heading/caption numbering-label patterns.

Used by extraction (20_extract_structure.py), the model-review-apply stage
(27_apply_review.py), and the text-rewrite stage (40_apply_fixes.py) so the
three never drift out of sync with each other.

Three different jobs are kept deliberately separate here:

  * infer_heading_level() — a best-effort SHAPE/STYLE guess at a paragraph's
    heading level AND how confident that guess is (see "source" below). It
    cannot always be right: a document may use bare "1." for its actual
    level-1 sections, which mechanically looks like the level-3 shape ("1."
    + digit). That is exactly why the skill also does a full-document model
    review pass (see 26_export_review.py / 27_apply_review.py) — the model
    looks at all candidates together and can correct a level this shape-only
    guess gets wrong. Deterministic code alone can't reliably cover every
    document's ground truth here.

  * detection SOURCE ("outline" / "style" / "pattern") — a heading detected
    via a real w:outlineLvl or a heading-ish paragraph style is a much more
    reliable signal than one guessed purely from a short line matching a
    numbering shape (a "pattern" match — plenty of non-heading short lines
    can accidentally look like "3." or "（一）"). Downstream, only
    outline/style-confirmed headings (or ones the model has explicitly
    confirmed via 27_apply_review.py, which upgrades them to
    "model_confirmed") get auto-renumbered; "pattern"-only ones only ever
    get a review hint — never a silent text edit — until a human-equivalent
    (the model, reading full-document context) explicitly confirms them.
    Captions get the same treatment via CAPTION_STYLE_HINTS / caption_source.

  * ANY_LABEL_RE / parse_leading_label() — a LEVEL-INDEPENDENT match for
    "does this paragraph have *some* leading enumeration label, and what is
    its raw text". Continuity checking only ever needs to compare this raw
    text against the position-derived canonical token for whatever level the
    paragraph ended up with (auto-guessed or model-corrected) — it does not
    need to know in advance which of the four shapes was used, so one
    generic pattern replaces what used to be four level-specific ones (and
    correctly handles the common real-world case where the original author's
    punctuation/numeral system doesn't match the shape a given level is
    supposed to use).
"""
import re

CN_NUM = "一二三四五六七八九十百"

# --- shape-based level GUESS (pre-review default only) ---------------------
RE_L1 = re.compile(r"^[%s]+、" % CN_NUM)                     # 一、
RE_L2 = re.compile(r"^[（(]\s*[%s]+\s*[）)]" % CN_NUM)   # （一）
RE_L3 = re.compile(r"^(\d{1,2})\.(?!\d)")                        # 3.  (not 3.1, not 2024.)
RE_L4 = re.compile(r"^[（(]\s*(\d{1,2})\s*[）)]")        # （4）

RE_CAPTION = re.compile(r"^\s*(图|表)\s*([0-9]+(?:[-\.–][0-9]+)?)(.*)$")
RE_TOCTITLE = re.compile(r"^\s*目\s*录\s*$")            # 目录 / 目 录

# TOC entry level from its style ("TOC1"/"toc 2"/"目录 3"): the trailing
# digit is the level, used to pick that level's target indent.
_TOC_LVL_RE = re.compile(r"(?:toc|目录)\s*([1-9])", re.I)

# --- level number carried by a heading STYLE NAME -------------------------
# "标题 1" / "Heading 1" / "一级标题" / "1级标题" all state the level
# explicitly. When a heading style says which level it is, that is far more
# reliable than guessing the level from the numbering SHAPE of the text —
# e.g. a genuine level-1 heading written "1. 项目进展" mechanically looks like
# the level-3 shape ("digit + dot"), and if the shape wins, the heading is
# mis-leveled to 3, its ordinal "1." then looks correct for level 3, and it
# never gets renumbered to the level-1 form "一、". So the style-name level
# takes precedence over the shape guess (an actual w:outlineLvl still wins
# over both — it is the most authoritative signal).
_LVL_NAME_RE = re.compile(r"(?:heading|标题)\s*([1-9])", re.I)
_LVL_AR_RE = re.compile(r"([1-9])\s*级")
_LVL_CN_RE = re.compile(r"([一二三四五六七八九])\s*级")

# Paragraph style name/id substrings that specifically mean "this is a
# figure/table caption", as distinct from the generic "heading" hint below.
# Deliberately checked BEFORE outline/heading-style so a caption paragraph
# that happens to inherit an outlineLvl (template quirk) or whose style name
# also contains "标题" (e.g. a custom "图表标题"/"图标题"/"表标题" caption
# style — "标题" is a substring of all of those) is never misclassified as a
# heading. NOTE: "图标题"/"表标题" must be listed explicitly — the shorter
# "图题"/"表题" are NOT substrings of them ("标" sits in between), so without
# these entries a "表标题"-styled paragraph would fall through to the generic
# "标题"-in-name heading test and be wrongly treated as a level-1 heading.
CAPTION_STYLE_HINTS = ("题注", "caption", "图表标题", "表格标题",
                       "图标题", "表标题", "图题", "表题")

# --- level-independent leading-label matcher --------------------------------
# Any of the canonical shapes, in either numeral system, so a mismatch
# between the ORIGINAL punctuation/numeral system and what the (possibly
# model-corrected) level requires can still be found and replaced wholesale.
#
# The FIRST alternative handles multi-level dotted numbers ("1.1", "1.1.1",
# "2.3", optionally with a trailing dot) which the single-dot alternative below
# deliberately rejects (its (?!\d) guard fails on "1.1"). A confirmed heading
# numbered "1.1" is a real, selectable ordinal that must be renumbered to the
# level's spec token (a level-2 "1.1  目标…" -> "（一）目标…"), not treated as
# "no number".
# The (later) whitespace alternative handles ordinals with NO punctuation,
# separated from
# the title only by whitespace — "5 项目运行管理情况", "2\t背景", "一 概述".
# It is deliberately limited to a 1–2 digit arabic number (or a CN numeral),
# so a title that merely STARTS with a year or long number ("2024 年度总结")
# is NOT mistaken for an ordinal — 2024 is four digits and can never match.
# The trailing whitespace is consumed so replacing the ordinal leaves clean
# text ("5 项目…" -> "二、项目…", not "二、 项目…"). This matcher only ever
# feeds ordinal parsing / the renumber text-strip of an ALREADY-confirmed
# heading, never the decision of whether a paragraph is a heading, so a bare
# "5 " in ordinary body text can't turn that paragraph into a heading.
ANY_LABEL_RE = re.compile(
    r"^(?:"
    r"(?:\d+[.．])+\d+[.．]?[ \t　]*"
    r"|(?:[%s]+|\d{1,4})[、.．](?!\d)"
    r"|[（(]\s*(?:[%s]+|\d{1,3})\s*[）)]"
    r"|(?:[%s]+|\d{1,2})[ \t　]+"
    r")" % (CN_NUM, CN_NUM, CN_NUM)
)


def parse_leading_label(text):
    """Return the raw leading numbering-label text, or None if there isn't one."""
    m = ANY_LABEL_RE.match(text)
    return m.group(0) if m else None


def _style_name(style_id, resolver):
    from docxcommon import qn  # local import: avoid a hard dependency for
                                # any caller that only wants the regexes.
    sid = (style_id or "")
    name = ""
    if resolver and sid in resolver.styles:
        nm = resolver.styles[sid].find(qn("w:name"))
        if nm is not None:
            name = nm.get(qn("w:val")) or ""
    return sid, name


def looks_like_caption_style(style_id, resolver):
    """True if the paragraph style name/id specifically suggests a
    figure/table caption (题注/Caption/图表标题/...), as opposed to a generic
    heading style."""
    sid, name = _style_name(style_id, resolver)
    hint = (sid + " " + name).lower()
    return any(h.lower() in hint for h in CAPTION_STYLE_HINTS)


def toc_level_from_style(style_id, resolver):
    """Level (1..N) of a TOC entry from its style id/name ('TOC1'/'toc 2'/
    '目录 3'), or None if it can't be determined."""
    sid, name = _style_name(style_id, resolver)
    m = _TOC_LVL_RE.search(sid + " " + name)
    return int(m.group(1)) if m else None


def style_is_toc(style_id, resolver):
    """True if a paragraph's style is a TOC ENTRY style — detected by id OR by
    the (possibly localized) style NAME ('TOC1'/'toc 2'/'目录 3').

    Checking the NAME, not just the id, is essential: many 中文/WPS documents
    give their TOC entry styles a non-'TOC…' styleId (numeric or otherwise)
    while the NAME is '目录 N'. Only the FIRST TOC paragraph carries the TOC
    field instruction; the remaining entries are recognized solely by their
    style, so a style test that misses the localized name silently drops every
    entry after the first into the heading branch (wrong font/indent, and a
    renumber may even rewrite the entry text)."""
    sid, name = _style_name(style_id, resolver)
    if sid.lower().startswith("toc"):
        return True
    return bool(_TOC_LVL_RE.search(sid + " " + name))


def caption_kind_from_style(style_id, resolver):
    """Infer 'figure'/'table' from a caption style's NAME when the paragraph
    text itself carries no 图/表 prefix (e.g. a "表标题"-styled line that just
    reads "设备清单"). Returns None when the style name is ambiguous — mentions
    both 图 and 表 (a generic "图表标题" caption style) or neither — in which
    case the kind can only come from the text, not the style."""
    sid, name = _style_name(style_id, resolver)
    s = (sid + " " + name)
    low = s.lower()
    has_fig = ("图" in s) or ("fig" in low)
    has_tbl = ("表" in s) or ("table" in low)
    if has_fig and not has_tbl:
        return "figure"
    if has_tbl and not has_fig:
        return "table"
    return None


def heading_level_from_style_name(style_id, name):
    """Level (1..4) explicitly stated by a heading style NAME/ID
    ('标题 1'/'Heading 1'/'一级标题'/'1级标题'), or None if the name carries
    no level number. Higher-than-4 levels clamp to 4 (the spec only defines
    four heading levels)."""
    s = (style_id or "") + " " + (name or "")
    m = _LVL_NAME_RE.search(s) or _LVL_AR_RE.search(s)
    if m:
        return min(int(m.group(1)), 4)
    m = _LVL_CN_RE.search(s)
    if m:
        return min("一二三四五六七八九".index(m.group(1)) + 1, 4)
    return None


def infer_heading_level(style_id, outline, text, resolver):
    """Return (level, source). level is 1..4 or None; source is
    "outline"/"style"/"pattern" when level is not None, else None. See
    module docstring: this is the PRE-REVIEW default, not a final answer."""
    # 0) a caption-shaped or caption-styled paragraph is NEVER a heading —
    # checked FIRST, before outline/style, because either an inherited
    # outlineLvl or a "标题"-containing caption style name would otherwise
    # override this unambiguous signal (see CAPTION_STYLE_HINTS docstring).
    if RE_CAPTION.match(text) or looks_like_caption_style(style_id, resolver):
        return None, None
    # 1) resolved outline level wins (most authoritative)
    if outline is not None:
        return min(outline + 1, 4), "outline"
    # 2) style whose name/id looks like a heading
    sid, name = _style_name(style_id, resolver)
    hint = (sid + " " + name).lower()
    is_hstyle = ("heading" in hint) or ("标题" in (sid + name))
    # A heading style whose NAME states the level (标题 1 / 一级标题) pins the
    # level directly — this beats the numbering-shape guess below, so a
    # level-1 heading written "1." is not mis-read as level 3 (see
    # _LVL_NAME_RE docstring).
    name_level = heading_level_from_style_name(sid, name) if is_hstyle else None
    # 3) numbering pattern (only trust for short lines when no style backs it)
    short = len(text.strip()) <= 40
    for rx, lvl in ((RE_L1, 1), (RE_L2, 2), (RE_L4, 4), (RE_L3, 3)):
        if rx.match(text):
            if is_hstyle:
                return (name_level or lvl), "style"
            if short:
                return lvl, "pattern"
            return None, None
    if is_hstyle:
        return (name_level or 1), "style"
    return None, None
