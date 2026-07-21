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
from headings import ANY_LABEL_RE

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


def _set_first_line_and_clear_left(pPr, first_line_chars, clear_left, clear_right=False):
    ind = _get_or_make(pPr, "w:ind")
    if first_line_chars is not None:
        ind.set(qn("w:firstLineChars"), str(first_line_chars))
        # remove absolute first-line / hanging that would conflict
        for a in ("w:firstLine", "w:hanging", "w:hangingChars"):
            if ind.get(qn(a)) is not None:
                del ind.attrib[qn(a)]
    if clear_left:
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
def _replace_leading(p, strip_re, new_prefix):
    ts = [t for t in p.iter(qn("w:t"))]
    if not ts:
        return False
    full = "".join(t.text or "" for t in ts)
    m = strip_re.match(full)
    if not m:
        return False
    newfull = new_prefix + full[m.end():]
    ts[0].text = newfull
    ts[0].set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    for t in ts[1:]:
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


def _apply_renumber_caption(p, fix):
    kname = "\u56fe" if fix.get("kind") == "figure" else "\u8868"
    if fix.get("insert"):
        # The caption had no \u56fe/\u8868+\u6570\u5b57 label to replace -- add one (with a
        # trailing space before the existing content).
        return _prepend_text(p, kname + str(fix["new_num"]) + " ")
    new_prefix = kname + str(fix["new_num"])
    return _replace_leading(p, STRIP_CAPTION, new_prefix)


def _apply_renumber_heading(p, fix):
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
def _patch_toc_styles(pkg_dir, toc_spec):
    """Force the TOC entry styles (toc 1..N) to the spec's font/size and no
    indent, so a refreshed TOC renders per spec. Returns the number of styles
    patched. No-op when styles.xml or the toc spec is missing."""
    if not toc_spec or not toc_spec.get("east_asia") or not toc_spec.get("size_hp"):
        return 0
    path = os.path.join(pkg_dir, "word", "styles.xml")
    if not os.path.exists(path):
        return 0
    tree = parse_xml(path)
    root = tree.getroot()
    ea = toc_spec["east_asia"]
    sz = str(toc_spec["size_hp"])
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
        # pPr: no indent at all (entries flush-left per spec)
        ppr = _get_or_make(st, "w:pPr", before_tags=("w:rPr",))
        ind = _get_or_make(ppr, "w:ind")
        for a in ("w:leftChars", "w:left", "w:firstLineChars", "w:firstLine"):
            ind.set(qn(a), "0")
        for a in ("w:startChars", "w:start", "w:hanging", "w:hangingChars"):
            if ind.get(qn(a)) is not None:
                del ind.attrib[qn(a)]
        patched += 1
    if patched:
        tree.write(path, xml_declaration=True, encoding="UTF-8", standalone=True)
    return patched


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
                    or fix.get("clear_left_indent") or fix.get("clear_right_indent")):
                _set_first_line_and_clear_left(
                    pPr, fix.get("set_first_line_chars"),
                    bool(fix.get("clear_left_indent")),
                    bool(fix.get("clear_right_indent")))
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
