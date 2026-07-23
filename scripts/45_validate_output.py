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

Writes <workdir>/validate_report.json and prints a single-line JSON summary.
Exits non-zero (2) if any check fails, so the caller stops before finalizing.
"""
import json
import os
import sys
import zipfile

sys.path.insert(0, os.path.join(os.path.dirname(__file__), "lib"))

from lxml import etree
from docxcommon import qn, iter_body_paragraphs

REQUIRED_PARTS = ("[Content_Types].xml", "word/document.xml")


def _count_body_paragraphs(zf, name="word/document.xml"):
    """Parse document.xml straight from the zip and count body paragraphs with
    the SAME iterator extraction/apply use, so the numbers are comparable."""
    parser = etree.XMLParser(remove_blank_text=False, huge_tree=True)
    with zf.open(name) as f:
        root = etree.parse(f, parser).getroot()
    return sum(1 for _ in iter_body_paragraphs(root))


def validate(formatted_path, reference_path=None):
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

    report = validate(formatted, reference)
    json.dump(report, open(os.path.join(workdir, "validate_report.json"), "w",
                           encoding="utf-8"), ensure_ascii=False, indent=1)
    print(json.dumps(report, ensure_ascii=False))
    sys.exit(0 if report.get("ok") else 2)


if __name__ == "__main__":
    main()
