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

# --- leading-token strip patterns (MUST mirror 20_extract_structure.py) ----
CN_NUM = "\u4e00\u4e8c\u4e09\u56db\u4e94\u516d\u4e03\u516b\u4e5d\u5341\u767e"
STRIP_L1 = re.compile(r"^[%s]+\u3001" % CN_NUM)
STRIP_L2 = re.compile(r"^[\uff08(]\s*[%s]+\s*[\uff09)]" % CN_NUM)
STRIP_L3 = re.compile(r"^\d{1,2}\.(?!\d)")
STRIP_L4 = re.compile(r"^[\uff08(]\s*\d{1,2}\s*[\uff09)]")
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
    for r in p.iter(qn("w:r")):
        # skip runs living inside the comment reference we may add later; at
        # apply time there are none yet, so all runs are content runs.
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


def _set_first_line_and_clear_left(pPr, first_line_chars, clear_left):
    ind = _get_or_make(pPr, "w:ind")
    if first_line_chars is not None:
        ind.set(qn("w:firstLineChars"), str(first_line_chars))
        # remove absolute first-line / hanging that would conflict
        for a in ("w:firstLine", "w:hanging", "w:hangingChars"):
            if ind.get(qn(a)) is not None:
                del ind.attrib[qn(a)]
    if clear_left:
        for a in ("w:leftChars", "w:left", "w:startChars", "w:start"):
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


def _apply_renumber_caption(p, fix):
    kname = "\u56fe" if fix.get("kind") == "figure" else "\u8868"
    new_prefix = kname + str(fix["new_num"])
    return _replace_leading(p, STRIP_CAPTION, new_prefix)


def _apply_renumber_heading(p, fix):
    lvl = fix.get("level", 1)
    strip = {1: STRIP_L1, 2: STRIP_L2, 3: STRIP_L3, 4: STRIP_L4}.get(lvl, STRIP_L1)
    return _replace_leading(p, strip, fix["new_token"])


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
            if fix.get("set_first_line_chars") is not None or fix.get("clear_left_indent"):
                _set_first_line_and_clear_left(
                    pPr, fix.get("set_first_line_chars"),
                    bool(fix.get("clear_left_indent")))
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
