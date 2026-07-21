# -*- coding: utf-8 -*-
"""
Stage 2 — extract a structured IR from the working .docx.

Usage:
    python 20_extract_structure.py <workdir> [--shard-size N]

Reads   <workdir>/working.docx
Writes  <workdir>/structure.json          (full IR; stays on disk, NOT for model)
        <workdir>/shards/shard_000.json ... (paragraph slices for sub-agents)
Prints  a COMPACT summary JSON to stdout (counts only), so the orchestrating
        model's context stays small even for 3000-page documents.

Key correctness properties (see skill notes):
  * Uses the EFFECTIVE format (docDefaults -> style basedOn chain -> direct),
    so formatting defined only in styles.xml (incl. inherited outlineLvl) is
    seen, not mis-reported as "unset".
  * Detects the TOC region and the cover region and tags them, so目录/封皮
    paragraphs are excluded from heading/caption checks and from continuity
    counting.
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
    iter_text_runs, run_effective_rpr,
)

# --- numbering patterns for heading level inference ------------------------
CN_NUM = "\u4e00\u4e8c\u4e09\u56db\u4e94\u516d\u4e03\u516b\u4e5d\u5341\u767e"
RE_L1 = re.compile(r"^[%s]+\u3001" % CN_NUM)                     # 一、
RE_L2 = re.compile(r"^[\uff08(]\s*[%s]+\s*[\uff09)]" % CN_NUM)   # （一）
RE_L3 = re.compile(r"^(\d{1,2})\.(?!\d)")                        # 3.  (not 3.1, not 2024.)
RE_L4 = re.compile(r"^[\uff08(]\s*(\d{1,2})\s*[\uff09)]")        # （4）
RE_CAPTION = re.compile(r"^\s*(\u56fe|\u8868)\s*([0-9]+(?:[-\.\u2013][0-9]+)?)(.*)$")
RE_TOCTITLE = re.compile(r"^\s*\u76ee\s*\u5f55\s*$")            # 目录 / 目 录


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


def in_toc_sdt(p):
    anc = p.getparent()
    while anc is not None:
        if anc.tag == qn("w:sdt"):
            for g in anc.iter(qn("w:docPartGallery")):
                if (g.get(qn("w:val")) or "").lower().startswith("table of contents"):
                    return True
        anc = anc.getparent()
    return False


def infer_heading_level(style_id, outline, text, resolver):
    """Return heading level 1..4 (or None)."""
    # 1) resolved outline level wins
    if outline is not None:
        return min(outline + 1, 4)
    # 2) style whose name/id looks like a heading
    sid = (style_id or "")
    name = ""
    if resolver and sid in resolver.styles:
        nm = resolver.styles[sid].find(qn("w:name"))
        if nm is not None:
            name = nm.get(qn("w:val")) or ""
    hint = (sid + " " + name).lower()
    is_hstyle = ("heading" in hint) or ("\u6807\u9898" in (sid + name))
    # 3) numbering pattern (only trust for short lines)
    short = len(text.strip()) <= 40
    if RE_L1.match(text):
        return 1 if (is_hstyle or short) else None
    if RE_L2.match(text):
        return 2 if (is_hstyle or short) else None
    if RE_L4.match(text):
        return 4 if (is_hstyle or short) else None
    if RE_L3.match(text):
        return 3 if (is_hstyle or short) else None
    # A "\u56fe1.../\u88681..." caption is NEVER a heading, even when its paragraph
    # style name happens to contain "\u6807\u9898" (e.g. a custom "\u56fe\u8868\u6807\u9898" caption
    # style) \u2014 without this guard the generic is_hstyle fallback below claims
    # it as a level-1 heading, which also hides it from caption continuity
    # checking (that only runs when level is None).
    if RE_CAPTION.match(text):
        return None
    if is_hstyle:
        return 1
    return None


# Flexible per-level label regexes used ONLY to *extract the raw leading
# label* once a paragraph's level is already known (from outlineLvl/style/the
# strict RE_L* patterns above). Unlike RE_L1/RE_L3 these also accept the
# "wrong" numeral system (e.g. arabic "3\u3001" for a level-1 heading that should
# read "\u4e09\u3001"), so a numbering-FORMAT mistake is captured (raw != expected
# canonical token) instead of being silently skipped because it failed to
# parse under the strict pattern.
LABEL_RE = {
    1: re.compile(r"^([%s]+|\d{1,3})\u3001" % CN_NUM),
    2: re.compile(r"^[\uff08(]\s*([%s]+|\d{1,2})\s*[\uff09)]" % CN_NUM),
    3: re.compile(r"^([%s]+|\d{1,2})\.(?!\d)" % CN_NUM),
    4: re.compile(r"^[\uff08(]\s*([%s]+|\d{1,2})\s*[\uff09)]" % CN_NUM),
}


def parse_num_token(level, text):
    """Return (raw_label, numeric_value) for continuity/format checks, or
    (None, None) if no leading label for this level can be found at all."""
    rx = LABEL_RE.get(level)
    if rx is None:
        return None, None
    m = rx.match(text)
    if not m:
        return None, None
    raw = m.group(0)
    num_s = m.group(1)
    val = int(num_s) if num_s.isdigit() else _cn2int(num_s)
    return raw, val


_CN_MAP = {c: i + 1 for i, c in enumerate("\u4e00\u4e8c\u4e09\u56db\u4e94\u516d\u4e03\u516b\u4e5d")}
_CN_MAP["\u5341"] = 10


def _cn2int(s):
    s = s.strip()
    if not s:
        return None
    if s == "\u5341":
        return 10
    if "\u5341" in s:  # e.g. 十一, 二十, 二十三
        parts = s.split("\u5341")
        tens = _CN_MAP.get(parts[0], 1) if parts[0] else 1
        ones = _CN_MAP.get(parts[1], 0) if len(parts) > 1 and parts[1] else 0
        return tens * 10 + ones
    total = 0
    for ch in s:
        if ch in _CN_MAP:
            total = total * 10 + _CN_MAP[ch]
        else:
            return None
    return total or None


def main():
    workdir = sys.argv[1]
    shard_size = 400
    if "--shard-size" in sys.argv:
        shard_size = int(sys.argv[sys.argv.index("--shard-size") + 1])

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

    # ---- page setup (first sectPr) ----
    page_setup = None
    body = doc_root.find(qn("w:body"))
    sectPrs = body.findall(qn("w:sectPr")) + [
        s for s in body.iter(qn("w:sectPr"))]
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
    records = []
    for i, p in iter_body_paragraphs(doc_root):
        sid = get_style_id(p)
        ppr_el = get_pPr(p)
        mark_rpr = get_mark_rpr(p)
        ppr, rpr = resolver.resolve(sid, ppr_el, mark_rpr)
        text = para_text(p)
        ea, asc, sz = dominant_run_props(p, rpr)

        is_toc = bool(sid and sid.lower().startswith("toc")) or has_toc_field(p) or in_toc_sdt(p)
        is_toctitle = bool(RE_TOCTITLE.match(text))

        outline = ppr.get("outline")
        level = None
        if not is_toc and not is_toctitle:
            level = infer_heading_level(sid, outline, text, resolver)
        num_raw, num_token = parse_num_token(level, text) if level else (None, None)

        caption = None
        if not is_toc and not is_toctitle and level is None:
            mc = RE_CAPTION.match(text)
            if mc:
                rest = mc.group(3).strip()
                caption = {
                    "kind": "figure" if mc.group(1) == "\u56fe" else "table",
                    "num_raw": mc.group(2),
                    "has_content": bool(rest),
                }

        rec = {
            "i": i,
            "style_id": sid,
            "text": text[:60],
            "text_len": len(text),
            "is_toc": is_toc or is_toctitle,
            "is_heading": level is not None,
            "level": level,
            "num_token": num_token,
            "num_raw": num_raw,
            "caption": caption,
            "in_table": in_table(p),
            "eff": {
                "east_asia": ea, "ascii": asc, "size_hp": sz,
                "line": ppr.get("line"), "line_rule": ppr.get("line_rule"),
                "first_line_chars": ppr.get("first_line_chars"),
                "first_line": ppr.get("first_line"),
                "left_chars": ppr.get("left_chars"), "left": ppr.get("left"),
                "start_chars": ppr.get("start_chars"), "start": ppr.get("start"),
                "hanging_chars": ppr.get("hanging_chars"), "hanging": ppr.get("hanging"),
                "jc": ppr.get("jc"), "outline": outline,
            },
        }
        records.append(rec)

    # ---- region tagging: cover / toc / body ----
    first_toc_idx = next((r["i"] for r in records if r["is_toc"]), None)
    first_heading_idx = next((r["i"] for r in records if r["is_heading"]), None)
    cover_end = first_toc_idx if first_toc_idx is not None else first_heading_idx
    for r in records:
        if r["is_toc"]:
            r["region"] = "toc"
        elif cover_end is not None and r["i"] < cover_end:
            r["region"] = "cover"
        else:
            r["region"] = "body"

    # ---- page background (for cover green hint) ----
    bg = doc_root.find(qn("w:background"))
    has_background = bg is not None and (bg.get(qn("w:color")) or bg.get("{http://schemas.microsoft.com/office/word/2010/wordml}color"))

    structure = {
        "n_paragraphs": len(records),
        "page_setup": page_setup,
        "has_background": bool(has_background),
        "records": records,
    }
    with open(os.path.join(workdir, "structure.json"), "w", encoding="utf-8") as f:
        json.dump(structure, f, ensure_ascii=False)

    # ---- shards (paragraph records only) ----
    shard_dir = os.path.join(workdir, "shards")
    os.makedirs(shard_dir, exist_ok=True)
    for f in os.listdir(shard_dir):
        os.remove(os.path.join(shard_dir, f))
    shard_files = []
    for si, start in enumerate(range(0, len(records), shard_size)):
        chunk = records[start:start + shard_size]
        name = "shard_%03d.json" % si
        with open(os.path.join(shard_dir, name), "w", encoding="utf-8") as f:
            json.dump({"shard_id": si, "range": [chunk[0]["i"], chunk[-1]["i"]],
                       "records": chunk}, f, ensure_ascii=False)
        shard_files.append(name)

    summary = {
        "status": "ok",
        "n_paragraphs": len(records),
        "regions": {
            "cover": sum(1 for r in records if r["region"] == "cover"),
            "toc": sum(1 for r in records if r["region"] == "toc"),
            "body": sum(1 for r in records if r["region"] == "body"),
        },
        "headings": sum(1 for r in records if r["is_heading"]),
        "captions": sum(1 for r in records if r["caption"]),
        "n_shards": len(shard_files),
        "shard_dir": os.path.abspath(shard_dir),
        "shard_files": shard_files,
        "page_setup_found": page_setup is not None,
    }
    print(json.dumps(summary, ensure_ascii=False))


if __name__ == "__main__":
    main()
