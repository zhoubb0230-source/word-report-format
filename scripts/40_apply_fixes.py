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
import json
import os
import re
import sys

sys.path.insert(0, os.path.join(os.path.dirname(__file__), "lib"))

from lxml import etree
from docxcommon import (qn, parse_xml, unzip_docx, rezip_docx,
                        iter_body_paragraphs)
from commentwriter import CommentWriter
from headings import ANY_LABEL_RE, CAPTION_STYLE_HINTS

SPEC_PATH = os.path.join(os.path.dirname(__file__), "..", "spec", "format_spec.json")

# Paragraph styles that format the AUTO-GENERATED table-of-contents entries.
# Word regenerates these paragraphs from their style (not from direct
# formatting) whenever the TOC field refreshes — and we set updateFields=true
# so it refreshes on open — so the only way to make the TOC font/size/indent
# actually stick is to patch the styles themselves, not just the current
# entry runs.
RE_TOC_STYLE_ID = re.compile(r"^toc\s*\d+$", re.IGNORECASE)
RE_TOC_STYLE_NAME = re.compile(r"^(?:toc|目录)\s*\d+$", re.IGNORECASE)

# Level-independent: whatever leading enumeration label a heading currently
# has (regardless of numeral system / punctuation, and regardless of whether
# it matches the shape its assigned level "should" use) gets stripped and
# replaced with the canonical token computed by checks.py. Shared with
# 20_extract_structure.py / 27_apply_review.py via lib/headings.py so all
# three never drift out of sync.
STRIP_HEADING = ANY_LABEL_RE
STRIP_CAPTION = re.compile(r"^\s*(?:\u56fe|\u8868)\s*[0-9]+(?:[-\.\u2013][0-9]+)?")
# A bare \u56fe/\u8868 prefix with NO number. Used to strip the residual prefix left
# behind when a caption's number lived in a Word field that gets deleted (the
# static "\u56fe/\u8868" survives the field removal), or a \u8868\u6807\u9898-styled line that
# carries the \u56fe/\u8868 word but no digit -- so writing the new \u56feN/\u8868N never
# doubles the prefix ("\u56fe1\u56fe \u8bf4\u660e").
STRIP_CAPTION_RESIDUE = re.compile(r"^\s*(?:\u56fe|\u8868)")


# ---------------------------------------------------------------------------
# pPr / rPr helpers (create-or-get, keep OOXML child order roughly valid)
# ---------------------------------------------------------------------------
def _get_or_make(parent, tag, before_tags=()):
    el = parent.find(qn(tag))
    if el is not None:
        return el
    el = etree.Element(qn(tag))
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
        anc = r.getparent()
        in_txbx = False
        while anc is not None and anc is not p:
            if anc.tag == qn("w:txbxContent"):
                in_txbx = True
                break
            anc = anc.getparent()
        if in_txbx:
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


def _apply_run_props(p, east_asia, ascii_, size_hp):
    """Apply font/size to every content run + the paragraph-mark rPr."""
    for r in _iter_runs(p):
        rpr = _run_rpr(r)
        if east_asia is not None or ascii_ is not None:
            _set_fonts(rpr, east_asia, ascii_)
        if size_hp is not None:
            _set_size(rpr, size_hp)
    # paragraph mark run properties (pPr/rPr)
    pPr = get_pPr(p)
    mark = _get_or_make(pPr, "w:rPr",
                        before_tags=("w:sectPr",))
    if east_asia is not None or ascii_ is not None:
        _set_fonts(mark, east_asia, ascii_)
    if size_hp is not None:
        _set_size(mark, size_hp)


def _set_line_exact(pPr, line_twips):
    sp = _get_or_make(pPr, "w:spacing")
    sp.set(qn("w:line"), str(line_twips))
    sp.set(qn("w:lineRule"), "exact")
    # remove auto-spacing that would override the fixed value
    for a in ("w:beforeAutospacing", "w:afterAutospacing"):
        if sp.get(qn(a)) is not None:
            del sp.attrib[qn(a)]


def _set_first_line_and_clear_left(pPr, first_line_chars, clear_left, clear_right=False,
                                   set_left_chars=None):
    ind = _get_or_make(pPr, "w:ind")
    if first_line_chars is not None:
        ind.set(qn("w:firstLineChars"), str(first_line_chars))
        # remove absolute first-line / hanging that would conflict
        for a in ("w:firstLine", "w:hanging", "w:hangingChars"):
            if ind.get(qn(a)) is not None:
                del ind.attrib[qn(a)]
    if set_left_chars is not None:
        # Set the left indent to a SPECIFIC character count (TOC per-level
        # indent: 0/200/400). leftChars governs (char-based, East-Asian aware);
        # the absolute w:left is zeroed and the w:start synonyms dropped so
        # nothing overrides it.
        ind.set(qn("w:leftChars"), str(set_left_chars))
        ind.set(qn("w:left"), "0")
        for a in ("w:startChars", "w:start"):
            if ind.get(qn(a)) is not None:
                del ind.attrib[qn(a)]
    elif clear_left:
        # Force the left indent to 0 rather than merely deleting the direct
        # attribute: the indent we need to override is frequently INHERITED
        # from the paragraph style (a title style, or the TOC1/2/3 styles),
        # and deleting a direct attribute that isn't there leaves the style's
        # indent in effect. An explicit direct 0 wins over the inherited value.
        # The w:start/w:startChars synonyms are removed so they can't re-supply
        # a non-zero left indent alongside our 0.
        ind.set(qn("w:leftChars"), "0")
        ind.set(qn("w:left"), "0")
        for a in ("w:startChars", "w:start"):
            if ind.get(qn(a)) is not None:
                del ind.attrib[qn(a)]
    if clear_right:
        # Same reasoning for the right indent. Zeroing this (title/TOC only) is
        # what makes jc=center actually center on the full page width instead
        # of the width left after an un-cleared (often style-inherited) right
        # indent narrows it asymmetrically.
        ind.set(qn("w:rightChars"), "0")
        ind.set(qn("w:right"), "0")
        for a in ("w:endChars", "w:end"):
            if ind.get(qn(a)) is not None:
                del ind.attrib[qn(a)]


def _set_jc(pPr, val):
    jc = _get_or_make(pPr, "w:jc")
    jc.set(qn("w:val"), val)


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
    disturbing existing runs. Used to ADD a caption number (\u56feN/\u8868N) to a
    caption-styled paragraph that has no leading \u56fe/\u8868 label to replace."""
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
    "no numbering" value), so a caption we renumber to static \u56feN/\u8868N text is not
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
        # so the static \u56feN/\u8868N we write below is the only number that renders.
        _suppress_auto_numbering(p)
    kname = "\u56fe" if fix.get("kind") == "figure" else "\u8868"
    new_prefix = kname + str(fix["new_num"])
    # 1) a proper \u56fe/\u8868 + number label (typed, or a field whose result is in
    #    the text) -- replace it; residue_re cleans a bare \u56fe/\u8868 left if the
    #    number was a field that got deleted, so we never double the prefix.
    if _replace_leading(p, STRIP_CAPTION, new_prefix, residue_re=STRIP_CAPTION_RESIDUE):
        return True
    # 2) a bare \u56fe/\u8868 prefix with no number (style-detected caption whose number
    #    was missing) -- replace just that prefix.
    if _replace_leading(p, STRIP_CAPTION_RESIDUE, new_prefix):
        return True
    # 3) no \u56fe/\u8868 label in the text at all (auto-numbered caption whose number
    #    lived only in numPr, or a \u8868\u6807\u9898 line reading just "\u8bbe\u5907\u6e05\u5355") -- add one.
    return _prepend_text(p, new_prefix + " ")


def _heading_insert_prefix(token):
    """Prefix used when INSERTING a missing heading number. Arabic dotted tokens
    ("1.") read better with a trailing space before the title; full-width
    punctuation tokens ("\u4e00\u3001"/"\uff08\u4e00\uff09"/"\uff081\uff09") need none."""
    return token + " " if token.endswith(".") else token


def _apply_renumber_heading(p, fix):
    if fix.get("insert"):
        # Confirmed heading that lost its number entirely -- prepend the
        # position token (there is no existing label to replace).
        return _prepend_text(p, _heading_insert_prefix(fix["new_token"]))
    return _replace_leading(p, STRIP_HEADING, fix["new_token"])


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
    """Trailing digit of a TOC style id/name ('TOC1'/'toc 2') -> its level."""
    m = re.search(r"(?:toc|目录)\s*([1-9])", sid + " " + name, re.IGNORECASE)
    return int(m.group(1)) if m else None


def _patch_toc_styles(pkg_dir, toc_spec):
    """Force the TOC entry styles (toc 1..N) to the spec's font/size and the
    per-level indent (一级0/二级2字符/三级4字符), so a REFRESHED TOC renders
    per spec. Returns the number of styles patched. No-op when styles.xml or
    the toc spec is missing."""
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
        ind = _get_or_make(ppr, "w:ind")
        ind.set(qn("w:leftChars"), str(want_left))
        ind.set(qn("w:left"), "0")
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


def _set_update_fields(pkg_dir):
    path = os.path.join(pkg_dir, "word", "settings.xml")
    if not os.path.exists(path):
        return
    tree = parse_xml(path)
    root = tree.getroot()
    uf = root.find(qn("w:updateFields"))
    if uf is None:
        uf = etree.Element(qn("w:updateFields"))
        root.insert(0, uf)
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

    cw = CommentWriter(out_pkg, author="XAgent")
    applied = {"format": 0, "renumber_caption": 0, "renumber_heading": 0,
               "section": 0, "hint": 0, "comments": 0, "skipped": 0}
    problems = []

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
            _apply_run_props(p, fix.get("set_east_asia"),
                             fix.get("set_ascii"), fix.get("set_size_hp"))
            pPr = get_pPr(p)
            if fix.get("set_line_exact") is not None:
                _set_line_exact(pPr, fix["set_line_exact"])
            if (fix.get("set_first_line_chars") is not None
                    or fix.get("clear_left_indent") or fix.get("clear_right_indent")
                    or fix.get("set_left_chars") is not None):
                _set_first_line_and_clear_left(
                    pPr, fix.get("set_first_line_chars"),
                    bool(fix.get("clear_left_indent")),
                    bool(fix.get("clear_right_indent")),
                    fix.get("set_left_chars"))
            if fix.get("set_jc") is not None:
                _set_jc(pPr, fix["set_jc"])
            applied["format"] += 1
        elif op == "renumber_caption":
            ok = _apply_renumber_caption(p, fix)
            applied["renumber_caption"] += 1 if ok else 0
        elif op == "renumber_heading":
            ok = _apply_renumber_heading(p, fix)
            applied["renumber_heading"] += 1 if ok else 0
        elif op == "hint":
            applied["hint"] += 1
        else:
            problems.append({"para_index": idx, "op": op, "reason": "unknown_op"})
            applied["skipped"] += 1
            continue

        if not ok:
            problems.append({"para_index": idx, "op": op, "reason": "text_edit_failed"})

        # attach XAgent comment (rule text) when requested
        if fix.get("comment") and fix.get("rule_text"):
            cw.add(p, fix["rule_text"])
            applied["comments"] += 1

    cw.flush()

    # Patch the TOC entry styles so a refreshed TOC keeps the spec font/size
    # and no indent (the direct formatting applied above is otherwise wiped
    # when Word rebuilds the field on open).
    try:
        with open(SPEC_PATH, encoding="utf-8") as f:
            _spec = json.load(f)
        applied["toc_styles"] = _patch_toc_styles(out_pkg, _spec.get("toc"))
    except (OSError, ValueError):
        applied["toc_styles"] = 0

    # Un-hide auto-generated caption numbers (图N/表N) obscured by a black
    # shading / zero size in the caption numbering definition.
    applied["caption_num_unhidden"] = _clean_caption_numbering(out_pkg)

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
