# -*- coding: utf-8 -*-
"""
Stage 2 — extract a structured IR from the working .docx.

Usage:
    python 20_extract_structure.py <workdir>

Reads   <workdir>/working.docx
Writes  <workdir>/structure.json          (full IR; stays on disk, NOT for model)
Prints  a COMPACT summary JSON to stdout (counts only), so the orchestrating
        model's context stays small even for 3000-page documents.

Key correctness properties (see skill notes):
  * Uses the EFFECTIVE format (docDefaults -> style basedOn chain -> direct),
    so formatting defined only in styles.xml (incl. inherited outlineLvl) is
    seen, not mis-reported as "unset".
  * Detects the TOC region and the cover region and tags them, so目录/封皮
    paragraphs are excluded from heading/caption checks and from continuity
    counting.
  * Blank/whitespace-only paragraphs are tagged is_blank and never classified
    as a title/heading/caption/body — an empty line has no format to judge.
  * Heading LEVEL here is only a best-effort shape/style GUESS (see
    lib/headings.py). It is not the final answer: run 26_export_review.py /
    27_apply_review.py afterwards so the model can correct levels the shape
    heuristic gets wrong before any font/numbering fix is computed from it.
  * Uses the shared iter_body_paragraphs() so indices line up with apply stage.

Deterministic. No model calls.
"""
import json
import os
import re
import sys

sys.path.insert(0, os.path.join(os.path.dirname(__file__), "lib"))
from docxcommon import (
    qn, parse_xml, unzip_docx, iter_body_paragraphs, para_text, in_table,
    StyleResolver, read_ppr, read_rpr, get_pPr, get_style_id, get_mark_rpr,
    iter_text_runs, run_effective_rpr, load_numbering_levels, INDENT_KEYS,
)
from headings import (
    RE_CAPTION, RE_TOCTITLE, infer_heading_level, parse_leading_label,
    looks_like_caption_style, caption_kind_from_style, toc_level_from_style,
    style_is_toc,
)
from structure import tag_regions


def dominant_run_props(p, baseline):
    """Return (east_asia, ascii, size_hp) most used across text runs."""
    ea_count, as_count, sz_count = {}, {}, {}
    total = 0
    for run, txt in iter_text_runs(p):
        eff = run_effective_rpr(run, baseline)
        n = len(txt)
        total += n
        if eff.get("east_asia"):
            ea_count[eff["east_asia"]] = ea_count.get(eff["east_asia"], 0) + n
        if eff.get("ascii"):
            as_count[eff["ascii"]] = as_count.get(eff["ascii"], 0) + n
        if eff.get("size_hp"):
            sz_count[eff["size_hp"]] = sz_count.get(eff["size_hp"], 0) + n

    def top(d, fallback):
        return max(d.items(), key=lambda kv: kv[1])[0] if d else fallback
    ea = top(ea_count, baseline.get("east_asia"))
    asc = top(as_count, baseline.get("ascii"))
    sz = top(sz_count, baseline.get("size_hp"))
    return ea, asc, sz


def has_toc_field(p):
    for it in p.iter(qn("w:instrText")):
        if it.text and "TOC" in it.text.upper():
            return True
    for fs in p.iter(qn("w:fldSimple")):
        instr = fs.get(qn("w:instr")) or ""
        if "TOC" in instr.upper():
            return True
    return False


def _has_page_field(p):
    """True if paragraph p contains a PAGE (or NUMPAGES) field — the page number."""
    for it in p.iter(qn("w:instrText")):
        if it.text and "PAGE" in it.text.upper():
            return True
    for fs in p.iter(qn("w:fldSimple")):
        if "PAGE" in (fs.get(qn("w:instr")) or "").upper():
            return True
    return False


def scan_page_number(unpack):
    """Best-effort: find the page-number run in headers/footers and return its
    western font + size {ascii, size_hp}, or None if no PAGE field is present.

    The page number lives in a header/footer PAGE field, not the body, so it is
    read here and recorded into structure for a HINT-ONLY check downstream
    (doc_hints). A rough read of the direct rPr on the field's paragraph runs is
    enough for a hint; nothing is auto-changed from it."""
    import glob
    word = os.path.join(unpack, "word")
    if not os.path.isdir(word):
        return None
    parts = sorted(glob.glob(os.path.join(word, "header*.xml"))
                   + glob.glob(os.path.join(word, "footer*.xml")))
    for part in parts:
        try:
            root = parse_xml(part).getroot()
        except Exception:
            continue
        for p in root.iter(qn("w:p")):
            if not _has_page_field(p):
                continue
            asc, sz = None, None
            for run in p.iter(qn("w:r")):
                rpr = run.find(qn("w:rPr"))
                if rpr is None:
                    continue
                rf = rpr.find(qn("w:rFonts"))
                if rf is not None and rf.get(qn("w:ascii")):
                    asc = rf.get(qn("w:ascii"))
                s = rpr.find(qn("w:sz"))
                if s is not None and s.get(qn("w:val")):
                    try:
                        sz = int(s.get(qn("w:val")))
                    except ValueError:
                        pass
            return {"ascii": asc, "size_hp": sz}
    return None


def in_toc_sdt(p):
    anc = p.getparent()
    while anc is not None:
        if anc.tag == qn("w:sdt"):
            for g in anc.iter(qn("w:docPartGallery")):
                if (g.get(qn("w:val")) or "").lower().startswith("table of contents"):
                    return True
        anc = anc.getparent()
    return False


def toc_field_para_indices(doc_root):
    """Body-paragraph indices covered by a TOC FIELD span (fldChar begin..end).

    A Word table of contents is ONE field spanning MANY paragraphs, but only the
    first entry carries the 'TOC' instruction text; the continuation entries hold
    just the field RESULT. Detecting the TOC only per-paragraph (has_toc_field)
    therefore tags only the first entry — and any continuation entry that carries
    an inherited outlineLvl or heading-ish style then gets mis-detected as a
    HEADING (wrong 黑体 font + indent, and it corrupts the cover boundary and the
    numbering continuity). Tracking the field span across paragraphs tags every
    entry as TOC. Nested per-entry fields (HYPERLINK / PAGEREF) are handled by a
    depth stack so the outer TOC stays open until its own matching end."""
    inside = set()
    open_fields = []  # stack of {"is_toc": bool, "instr": str} per open field
    for idx, p in iter_body_paragraphs(doc_root):
        touched = any(f["is_toc"] for f in open_fields)
        for el in p.iter():
            tag = el.tag
            if tag == qn("w:fldChar"):
                typ = el.get(qn("w:fldCharType"))
                if typ == "begin":
                    open_fields.append({"is_toc": False, "instr": ""})
                elif typ == "end" and open_fields:
                    open_fields.pop()
            elif tag == qn("w:instrText") and open_fields:
                open_fields[-1]["instr"] += (el.text or "")
                if "TOC" in open_fields[-1]["instr"].upper():
                    open_fields[-1]["is_toc"] = True
            elif tag == qn("w:fldSimple"):
                if "TOC" in (el.get(qn("w:instr")) or "").upper():
                    touched = True
            if any(f["is_toc"] for f in open_fields):
                touched = True
        if touched:
            inside.add(idx)
    return inside


# Font size (half-points) at/above which a cover line is "large" enough to be
# part of the report title. Body text is 三号 (32); the title band starts at
# 小一 (36), so 36 cleanly separates title lines from ordinary cover text.
TITLE_LARGE_HP = 36


def _title_candidate(r):
    """A cover paragraph that could be a title line: large font, short, and not
    a labelled cover field (密级：/项目名称：… — those carry a colon)."""
    if r.get("region") != "cover" or r.get("is_blank"):
        return False
    if (r["eff"].get("size_hp") or 0) < TITLE_LARGE_HP:
        return False
    if (r.get("text_len") or 0) > 40:
        return False
    t = r.get("text") or ""
    return ("：" not in t) and (":" not in t)


def _mark_title_block(records):
    """Tag the report-title paragraph(s) on the cover with is_title=True.

    The title is the largest text on the cover and often WRAPS across two or
    more paragraphs ("先进项目" / "2024年度自评价报告"); a wrapped line may have
    lost its centering, so per-paragraph "looks centered & large" detection
    misses it. Here we have the whole document: take contiguous run(s) of
    large, short, non-field cover lines and, for any run that contains a
    confident anchor (a line at the max title size, or a centered large line),
    tag every line in that run — so a wrapped title is captured whole and each
    line gets centered/re-fonted downstream. Blank lines between title lines do
    not break the run (blanks are already excluded from the candidate list)."""
    cover = [r for r in records if r.get("region") == "cover" and not r.get("is_blank")]
    cands = [r for r in cover if _title_candidate(r)]
    if not cands:
        return
    maxsz = max((r["eff"].get("size_hp") or 0) for r in cands)

    def is_anchor(r):
        return (r["eff"].get("size_hp") or 0) == maxsz or r["eff"].get("jc") == "center"

    candset = {id(r) for r in cands}
    n = len(cover)
    i = 0
    while i < n:
        if id(cover[i]) in candset:
            j = i
            while j + 1 < n and id(cover[j + 1]) in candset:
                j += 1
            run = cover[i:j + 1]
            if any(is_anchor(r) for r in run):
                for r in run:
                    r["is_title"] = True
            i = j + 1
        else:
            i += 1

    # Positional signal: the 密级/文本编号 line sits ABOVE the title (fixed by the
    # template, though not necessarily the very first line). Tag every non-blank
    # cover line before the first title line so cover_role() can default it to
    # "classification" even when the line has no 密级/文本编号 keyword (e.g. just
    # "机密  202501323023"). The AI review can still override.
    first_title_pos = next((idx for idx, r in enumerate(cover) if r.get("is_title")), None)
    if first_title_pos is not None:
        for r in cover[:first_title_pos]:
            r["above_title"] = True


def main():
    if len(sys.argv) < 2:
        print(json.dumps({"status": "error",
                          "error": "usage: 20_extract_structure.py <workdir>"}))
        sys.exit(1)
    workdir = sys.argv[1]

    unpack = os.path.join(workdir, "unpacked_read")
    if os.path.isdir(unpack):
        import shutil as _sh
        _sh.rmtree(unpack)
    unzip_docx(os.path.join(workdir, "working.docx"), unpack)

    doc_tree = parse_xml(os.path.join(unpack, "word", "document.xml"))
    doc_root = doc_tree.getroot()

    styles_path = os.path.join(unpack, "word", "styles.xml")
    styles_root = parse_xml(styles_path).getroot() if os.path.exists(styles_path) else None
    resolver = StyleResolver(styles_root)

    numbering_path = os.path.join(unpack, "word", "numbering.xml")
    numbering_root = parse_xml(numbering_path).getroot() if os.path.exists(numbering_path) else None
    numbering_levels = load_numbering_levels(numbering_root)

    # ---- page setup (first sectPr) ----
    page_setup = None
    body = doc_root.find(qn("w:body"))
    seen = set()
    first_sect = None
    for s in body.iter(qn("w:sectPr")):
        if id(s) in seen:
            continue
        seen.add(id(s))
        first_sect = first_sect or s
    if first_sect is not None:
        pg = first_sect.find(qn("w:pgMar"))
        if pg is not None:
            def gi(a):
                v = pg.get(qn("w:" + a))
                return int(v) if v is not None and v.lstrip("-").isdigit() else None
            page_setup = {
                "top": gi("top"), "bottom": gi("bottom"),
                "left": gi("left"), "right": gi("right"),
                "header": gi("header"), "footer": gi("footer"),
            }

    # ---- iterate paragraphs ----
    # Paragraphs inside a TOC field span (many entries, only the first carries
    # the TOC instruction) — tagged TOC below so continuation entries are never
    # mistaken for headings.
    toc_span = toc_field_para_indices(doc_root)
    records = []
    for i, p in iter_body_paragraphs(doc_root):
        sid = get_style_id(p)
        ppr_el = get_pPr(p)
        mark_rpr = get_mark_rpr(p)
        ppr, rpr = resolver.resolve(sid, ppr_el, mark_rpr)
        text = para_text(p)
        ea, asc, sz = dominant_run_props(p, rpr)

        # Automatic list numbering (w:numPr, possibly inherited from the style
        # chain). numId "0" is the explicit "no numbering" override. When a
        # paragraph IS auto-numbered, its visible number ("一、") is generated
        # by Word and does NOT appear in the text, and its indent frequently
        # comes from the numbering LEVEL definition rather than the paragraph.
        # Pull that level's indent in as a fallback so the effective indent we
        # judge/clear reflects what the reader actually sees.
        num_id = ppr.get("num_id")
        auto_num = bool(num_id and num_id != "0")
        if auto_num:
            try:
                ilvl = int(ppr.get("ilvl") or 0)
            except ValueError:
                ilvl = 0
            lvl_ppr = numbering_levels.get(num_id, {}).get(ilvl, {})
            for k in INDENT_KEYS:
                if ppr.get(k) is None and lvl_ppr.get(k) is not None:
                    ppr[k] = lvl_ppr[k]

        is_blank = not text.strip()
        is_toc = (style_is_toc(sid, resolver) or has_toc_field(p)
                  or in_toc_sdt(p) or (i in toc_span))
        is_toctitle = bool(RE_TOCTITLE.match(text))
        toc_level = toc_level_from_style(sid, resolver) if (is_toc and not is_toctitle) else None

        outline = ppr.get("outline")
        level = None
        level_source = None
        if not is_blank and not is_toc and not is_toctitle:
            level, level_source = infer_heading_level(sid, outline, text, resolver)
        num_raw = parse_leading_label(text) if level else None

        caption = None
        if not is_blank and not is_toc and not is_toctitle and level is None:
            mc = RE_CAPTION.match(text)
            style_says_caption = looks_like_caption_style(sid, resolver)
            if mc:
                rest = mc.group(3).strip()
                caption = {
                    "kind": "figure" if mc.group(1) == "图" else "table",
                    "num_raw": mc.group(2),
                    "has_content": bool(rest),
                    "source": "style" if style_says_caption else "pattern",
                }
            elif style_says_caption:
                # Caption style confirmed (题注/图标题/表标题/...), but the text
                # does not match the "图/表 + 数字" shape -- either it starts
                # with 图/表 yet has no digit, OR it carries no 图/表 prefix at
                # all (e.g. a "表标题"-styled line reading just "设备清单").
                # Either way the paragraph IS a caption whose number is missing,
                # not merely unrecognized, so it is kept as a caption record
                # (num_raw=None) rather than dropped -- continuity() then
                # auto-inserts the correct 图N/表N number (confirmed captions
                # are auto-fixable). The kind comes from the 图/表 prefix when
                # present, else from the style name.
                first = text[:1]
                if first in ("图", "表"):
                    kind = "figure" if first == "图" else "table"
                    content = text[1:].strip()
                else:
                    kind = caption_kind_from_style(sid, resolver)
                    content = text.strip()
                if kind:
                    caption = {
                        "kind": kind,
                        "num_raw": None,
                        "has_content": bool(content),
                        "source": "style",
                    }

        rec = {
            "i": i,
            "style_id": sid,
            "text": text[:60],
            "text_len": len(text),
            "is_blank": is_blank,
            "is_toc": is_toc or is_toctitle,
            "toc_level": toc_level,
            "auto_num": auto_num,
            "is_title": False,     # set by _mark_title_block() after region tagging
            "above_title": False,  # cover line positioned above the title block
            "cover_role": None,    # set only by the step-2.5 AI review (27_apply_review.py)
            "is_heading": level is not None,
            "level": level,
            "level_source": level_source,
            "num_raw": num_raw,
            "caption": caption,
            "in_table": in_table(p),
            # Whether the paragraph actually contains Western text (Latin letters
            # or digits). The western-font rule only applies where such text
            # exists, so a pure-Chinese paragraph is never flagged for its Latin
            # font (computed from the FULL text, before the 60-char truncation).
            "has_western": bool(re.search(r"[A-Za-z0-9]", text)),
            "eff": {
                "east_asia": ea, "ascii": asc, "size_hp": sz,
                "line": ppr.get("line"), "line_rule": ppr.get("line_rule"),
                "first_line_chars": ppr.get("first_line_chars"),
                "first_line": ppr.get("first_line"),
                "left_chars": ppr.get("left_chars"), "left": ppr.get("left"),
                "start_chars": ppr.get("start_chars"), "start": ppr.get("start"),
                "right_chars": ppr.get("right_chars"), "right": ppr.get("right"),
                "end_chars": ppr.get("end_chars"), "end": ppr.get("end"),
                "hanging_chars": ppr.get("hanging_chars"), "hanging": ppr.get("hanging"),
                "jc": ppr.get("jc"), "outline": outline,
            },
        }
        records.append(rec)

    # ---- region tagging: cover / toc / body ----
    tag_regions(records)

    # ---- title block on the cover (may wrap across paragraphs) ----
    _mark_title_block(records)

    # ---- page background (for cover green hint) ----
    bg = doc_root.find(qn("w:background"))
    has_background = bg is not None and (bg.get(qn("w:color")) or bg.get("{http://schemas.microsoft.com/office/word/2010/wordml}color"))

    structure = {
        "n_paragraphs": len(records),
        "page_setup": page_setup,
        "has_background": bool(has_background),
        "page_number": scan_page_number(unpack),
        "records": records,
    }
    with open(os.path.join(workdir, "structure.json"), "w", encoding="utf-8") as f:
        json.dump(structure, f, ensure_ascii=False)

    summary = {
        "status": "ok",
        "n_paragraphs": len(records),
        "regions": {
            "cover": sum(1 for r in records if r["region"] == "cover"),
            "toc": sum(1 for r in records if r["region"] == "toc"),
            "body": sum(1 for r in records if r["region"] == "body"),
        },
        "blank": sum(1 for r in records if r["is_blank"]),
        "headings": sum(1 for r in records if r["is_heading"]),
        "headings_unconfirmed": sum(1 for r in records if r["is_heading"] and r["level_source"] == "pattern"),
        "captions": sum(1 for r in records if r["caption"]),
        "captions_unconfirmed": sum(1 for r in records if r["caption"] and r["caption"].get("source") == "pattern"),
        "page_setup_found": page_setup is not None,
        "next_step": "python 26_export_review.py <workdir>  # full-document heading/caption review before checking format",
    }
    print(json.dumps(summary, ensure_ascii=False))


if __name__ == "__main__":
    main()
