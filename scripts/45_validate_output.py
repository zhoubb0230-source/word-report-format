# -*- coding: utf-8 -*-
"""Stage 4b — structural sanity check of the produced formatted.docx.

Usage:
    python 45_validate_output.py <workdir>

Runs AFTER 40_apply_fixes.py and BEFORE 50_finalize.py. Confirms the edited
document is not corrupted before it is delivered to the user — the failure mode
this guards against is real Microsoft Word opening the file and prompting
"内容有问题，需要修复". Checks (all lightweight, pure lxml — no python-docx):

  1. zip integrity        —— ZipFile.testzip() reports no bad CRC.
  2. required parts       —— [Content_Types].xml + word/document.xml present.
  3. XML well-formedness  —— every .xml / .rels part parses.
  4. paragraph parity     —— body-paragraph count in formatted.docx equals the
     pre-edit working.docx. Format / renumber / comment ops never add or remove
     a <w:p>, so a mismatch means the structure was damaged.
  5. numbering-clamp collapse invariant (方案C 甲法) —— every auto-numbered
     paragraph that received an indent fix must NOT still inherit a hanging from
     the NUMBERING layer in the output (that would mean the clone-clamp didn't
     take). Verified with the unified cascade + provenance, so a real leak fails
     hard; a style-inherited hanging (known-unfixed, Word-gated) is a non-fatal
     note only.

Writes <workdir>/validate_report.json and prints a single-line JSON summary.
Exits non-zero (2) if any check fails, so the caller stops before finalizing.
"""
import json
import os
import sys
import zipfile

sys.path.insert(0, os.path.join(os.path.dirname(__file__), "lib"))

from lxml import etree
from docxcommon import (qn, iter_body_paragraphs, StyleResolver,
                        load_numbering_levels, get_style_id, get_pPr, get_mark_rpr)
from cascade import resolve_ppr_with_provenance

REQUIRED_PARTS = ("[Content_Types].xml", "word/document.xml")


def _count_body_paragraphs(zf, name="word/document.xml"):
    """Parse document.xml straight from the zip and count body paragraphs with
    the SAME iterator extraction/apply use, so the numbers are comparable."""
    parser = etree.XMLParser(remove_blank_text=False, huge_tree=True)
    with zf.open(name) as f:
        root = etree.parse(f, parser).getroot()
    return sum(1 for _ in iter_body_paragraphs(root))


def _read_root(zf, name):
    parser = etree.XMLParser(remove_blank_text=False, huge_tree=True)
    with zf.open(name) as f:
        return etree.parse(f, parser).getroot()


def _check_numbering_clamp(zf, indent_fix_indices):
    """方案C 甲法坍缩不变量：在 OUTPUT 上重解析每个"拿到缩进修复的自动编号段"，
    断言其有效 hanging **不再由【编号层】供给**——否则说明克隆钳没生效（numbering.xml
    没去 hanging，或段落 numPr 没改指克隆）。用 cascade provenance 精确定位供给层：

      * hanging 仍由 `numbering` 供给 → LEAK（致命：钳失败，退2）。
      * hanging 由 `style`/`direct` 供给 → note（非致命）：样式层继承 hanging 属**已知
        未修**（待全角色 canonical 样式注入，Word-gated），不该在此硬失败误杀交付。

    只查"拿到缩进修复"的段落（避免对未触碰的合法列表段误报）。返回 (leaks, notes)。"""
    names = set(zf.namelist())
    if "word/document.xml" not in names:
        return [], []
    doc_root = _read_root(zf, "word/document.xml")
    styles_root = _read_root(zf, "word/styles.xml") if "word/styles.xml" in names else None
    numbering_root = (_read_root(zf, "word/numbering.xml")
                      if "word/numbering.xml" in names else None)
    resolver = StyleResolver(styles_root)
    levels = load_numbering_levels(numbering_root)
    para_by_idx = {i: p for i, p in iter_body_paragraphs(doc_root)}

    leaks, notes = [], []
    for idx in sorted(set(indent_fix_indices)):
        p = para_by_idx.get(idx)
        if p is None:
            continue
        eff, owner = resolve_ppr_with_provenance(
            resolver, get_style_id(p), get_pPr(p), get_mark_rpr(p), levels)
        nid = eff.get("num_id")
        if not nid or nid == "0":
            continue  # 非自动编号：不属甲法钳的职责范围
        for key in ("hanging", "hanging_chars"):
            v = eff.get(key)
            if v in (None, 0, "0"):
                continue
            rec = {"para_index": idx, "key": key, "value": v, "owner": owner.get(key)}
            (leaks if owner.get(key) == "numbering" else notes).append(rec)
    return leaks, notes


def validate(formatted_path, reference_path=None, indent_fix_indices=()):
    errors = []
    info = {}

    if not os.path.exists(formatted_path):
        return {"status": "error", "ok": False,
                "errors": ["formatted.docx not found: %s" % formatted_path]}

    # 1. zip integrity + open
    try:
        zf = zipfile.ZipFile(formatted_path)
    except zipfile.BadZipFile as e:
        return {"status": "error", "ok": False,
                "errors": ["not a valid zip/docx: %s" % e]}
    with zf:
        bad = zf.testzip()
        if bad is not None:
            errors.append("corrupted zip entry (bad CRC): %s" % bad)

        names = set(zf.namelist())

        # 2. required parts present
        for req in REQUIRED_PARTS:
            if req not in names:
                errors.append("missing required part: %s" % req)

        # 3. XML well-formedness of every xml/rels part
        parser = etree.XMLParser(remove_blank_text=False, huge_tree=True)
        malformed = []
        for name in names:
            if not (name.endswith(".xml") or name.endswith(".rels")):
                continue
            try:
                with zf.open(name) as f:
                    etree.parse(f, parser)
            except etree.XMLSyntaxError as e:
                malformed.append({"part": name, "error": str(e)})
        if malformed:
            errors.append("malformed XML in %d part(s)" % len(malformed))
            info["malformed"] = malformed

        # 4. paragraph parity vs the pre-edit reference
        if "word/document.xml" in names:
            try:
                n_out = _count_body_paragraphs(zf)
                info["paragraphs_out"] = n_out
            except etree.XMLSyntaxError as e:
                errors.append("cannot parse word/document.xml: %s" % e)
                n_out = None
            if reference_path and os.path.exists(reference_path) and n_out is not None:
                try:
                    with zipfile.ZipFile(reference_path) as rzf:
                        n_ref = _count_body_paragraphs(rzf)
                    info["paragraphs_reference"] = n_ref
                    if n_out != n_ref:
                        errors.append(
                            "paragraph count changed: formatted=%d, original=%d "
                            "(format/renumber/comment ops must not add or remove "
                            "paragraphs — structure likely damaged)" % (n_out, n_ref))
                except (zipfile.BadZipFile, KeyError, etree.XMLSyntaxError) as e:
                    info["paragraph_parity_skipped"] = "reference unreadable: %s" % e

        # 5. 方案C 甲法坍缩不变量：拿到缩进修复的自动编号段，有效 hanging 不得仍由
        #    编号层供给（那意味克隆钳没生效）。样式层继承 hanging 是已知未修，作非致命
        #    note。检查失败不阻断交付；只有"编号层泄漏"才致命。
        if indent_fix_indices:
            try:
                leaks, notes = _check_numbering_clamp(zf, indent_fix_indices)
                if leaks:
                    errors.append(
                        "编号层未钳干净：%d 段有效 hanging 仍来自编号层（克隆钳失败）"
                        % len(leaks))
                    info["numbering_leaks"] = leaks
                if notes:
                    info["style_hanging_notes"] = notes
            except (etree.XMLSyntaxError, KeyError) as e:
                info["clamp_invariant_skipped"] = str(e)

    ok = not errors
    return {"status": "ok" if ok else "error", "ok": ok,
            "formatted_docx": formatted_path, "errors": errors, **info}


def main():
    if len(sys.argv) < 2:
        print(json.dumps({"status": "error",
                          "error": "usage: 45_validate_output.py <workdir>"}))
        sys.exit(1)
    workdir = sys.argv[1]
    formatted = os.path.join(workdir, "formatted.docx")

    reference = None
    meta_path = os.path.join(workdir, "meta.json")
    if os.path.exists(meta_path):
        try:
            meta = json.load(open(meta_path, encoding="utf-8"))
            reference = meta.get("working_docx")
        except (OSError, ValueError):
            reference = None

    # Indices of paragraphs that received an indent fix (same trigger set 40 uses
    # for the numbering clamp) — the collapse invariant only inspects these.
    indent_fix_indices = []
    fixes_path = os.path.join(workdir, "fixes.json")
    if os.path.exists(fixes_path):
        try:
            for fx in json.load(open(fixes_path, encoding="utf-8")):
                if fx.get("op") == "format" and (
                        fx.get("set_first_line_chars") is not None
                        or fx.get("clear_left_indent")
                        or fx.get("clear_right_indent")
                        or fx.get("set_left_chars") is not None):
                    idx = fx.get("para_index")
                    if idx is not None:
                        indent_fix_indices.append(idx)
        except (OSError, ValueError):
            pass

    report = validate(formatted, reference, indent_fix_indices)
    json.dump(report, open(os.path.join(workdir, "validate_report.json"), "w",
                           encoding="utf-8"), ensure_ascii=False, indent=1)
    print(json.dumps(report, ensure_ascii=False))
    sys.exit(0 if report.get("ok") else 2)


if __name__ == "__main__":
    main()
