# -*- coding: utf-8 -*-
"""
headings.py — shared heading/caption numbering-label patterns.

Used by extraction (20_extract_structure.py), the model-review-apply stage
(27_apply_review.py), and the text-rewrite stage (40_apply_fixes.py) so the
three never drift out of sync with each other.

Two different jobs are kept deliberately separate here:

  * infer_heading_level() — a best-effort SHAPE/STYLE guess at a paragraph's
    heading level, used only as the PRE-REVIEW default. It cannot always be
    right: a document may use bare "1." for its actual level-1 sections,
    which mechanically looks like the level-3 shape ("1." + digit). That is
    exactly why the skill also does a full-document model review pass (see
    26_export_review.py / 27_apply_review.py) — the model looks at all
    candidates together and can correct a level this shape-only guess gets
    wrong. Deterministic code alone can't reliably cover every document's
    ground truth here.

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

# --- level-independent leading-label matcher --------------------------------
# Any of the four canonical shapes, in either numeral system, so a mismatch
# between the ORIGINAL punctuation/numeral system and what the (possibly
# model-corrected) level requires can still be found and replaced wholesale.
ANY_LABEL_RE = re.compile(
    r"^(?:"
    r"(?:[%s]+|\d{1,4})[、.．](?!\d)"
    r"|[（(]\s*(?:[%s]+|\d{1,3})\s*[）)]"
    r")" % (CN_NUM, CN_NUM)
)


def parse_leading_label(text):
    """Return the raw leading numbering-label text, or None if there isn't one."""
    m = ANY_LABEL_RE.match(text)
    return m.group(0) if m else None


def infer_heading_level(style_id, outline, text, resolver):
    """Return a best-effort heading level 1..4 (or None). See module docstring:
    this is the PRE-REVIEW default, not a final answer."""
    # 1) resolved outline level wins
    if outline is not None:
        return min(outline + 1, 4)
    # 2) style whose name/id looks like a heading
    from docxcommon import qn  # local import: avoid a hard dependency for
                                # any caller that only wants the regexes.
    sid = (style_id or "")
    name = ""
    if resolver and sid in resolver.styles:
        nm = resolver.styles[sid].find(qn("w:name"))
        if nm is not None:
            name = nm.get(qn("w:val")) or ""
    hint = (sid + " " + name).lower()
    is_hstyle = ("heading" in hint) or ("标题" in (sid + name))
    short = len(text.strip()) <= 40
    if RE_L1.match(text):
        return 1 if (is_hstyle or short) else None
    if RE_L2.match(text):
        return 2 if (is_hstyle or short) else None
    if RE_L4.match(text):
        return 4 if (is_hstyle or short) else None
    if RE_L3.match(text):
        return 3 if (is_hstyle or short) else None
    # A "图1.../表1..." caption is NEVER a heading, even when its paragraph
    # style name happens to contain "标题" (e.g. a custom "图表标题" caption
    # style) — without this guard the generic is_hstyle fallback below claims
    # it as a level-1 heading, which also hides it from caption continuity
    # checking (that only runs when level is None).
    if RE_CAPTION.match(text):
        return None
    if is_hstyle:
        return 1
    return None
