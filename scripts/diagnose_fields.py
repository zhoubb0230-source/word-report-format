# -*- coding: utf-8 -*-
"""Diagnose what makes Word prompt "该文档包含的域可能引用了其他文件。是否更新…"
on open. READ-ONLY — never modifies the file.

Usage:
    python diagnose_fields.py <docx路径>

Scans a .docx for every known trigger of that dialog and prints a report plus a
verdict listing the concrete source(s) found, so a fix can target the real cause
instead of guessing:

  1. settings.xml global <w:updateFields w:val="true"/>   (forces update on open)
  2. any field flagged w:dirty="true"                     (TOC/index etc.)
  3. fields that pull from OTHER FILES: INCLUDETEXT / INCLUDEPICTURE / LINK /
     DDE / DDEAUTO / RD / IMPORT / SUBDOCUMENT / AUTOTEXT
  4. external relationships (.rels TargetMode="External") + OLE link objects

Self-contained: matches elements by local-name so it does not depend on any
namespace-prefix table.
"""
import json
import os
import sys
import zipfile

from lxml import etree

# field commands that may read from a separate file -> "引用了其他文件"
EXTERNAL_FIELDS = {"INCLUDETEXT", "INCLUDEPICTURE", "INCLUDE", "LINK", "DDE",
                   "DDEAUTO", "RD", "IMPORT", "SUBDOCUMENT", "AUTOTEXT",
                   "AUTOTEXTLIST"}
# fields Word COMPILES from the document itself (updated by updateFields/dirty,
# but do not themselves reference an external file)
COMPILED_FIELDS = {"TOC", "TOA", "INDEX"}


def _local(tag):
    return tag.rsplit("}", 1)[-1] if isinstance(tag, str) and "}" in tag else tag


def _attr(el, name):
    for k, v in el.attrib.items():
        if _local(k) == name:
            return v
    return None


def _read_root(zf, name):
    with zf.open(name) as f:
        return etree.parse(f, etree.XMLParser(huge_tree=True)).getroot()


def _field_kind(instr):
    tok = instr.strip().split()[0].upper() if instr.strip() else ""
    if tok in EXTERNAL_FIELDS:
        return tok, "external"
    if tok in COMPILED_FIELDS:
        return tok, "compiled"
    return tok, "other"


def _complex_fields(root):
    """[{instr, dirty, token, kind}] for every complex field (begin..end)."""
    out, stack = [], []
    for el in root.iter():
        ln = _local(el.tag)
        if ln == "fldChar":
            typ = _attr(el, "fldCharType")
            if typ == "begin":
                d = _attr(el, "dirty")
                stack.append({"instr": "",
                              "dirty": d is not None and d not in ("0", "false")})
            elif typ == "end" and stack:
                item = stack.pop()
                item["token"], item["kind"] = _field_kind(item["instr"])
                out.append(item)
        elif ln == "instrText" and stack:
            stack[-1]["instr"] += (el.text or "")
    return out


def _simple_fields(root):
    out = []
    for el in root.iter():
        if _local(el.tag) != "fldSimple":
            continue
        instr = _attr(el, "instr") or ""
        d = _attr(el, "dirty")
        item = {"instr": instr, "dirty": d is not None and d not in ("0", "false")}
        item["token"], item["kind"] = _field_kind(instr)
        out.append(item)
    return out


def _update_fields_val(zf):
    if "word/settings.xml" not in zf.namelist():
        return "absent"
    root = _read_root(zf, "word/settings.xml")
    for el in root.iter():
        if _local(el.tag) == "updateFields":
            return _attr(el, "val") or "true"
    return "absent"


def _external_rels(zf):
    out = []
    for name in zf.namelist():
        if not name.endswith(".rels"):
            continue
        root = _read_root(zf, name)
        for el in root.iter():
            if _local(el.tag) == "Relationship" and _attr(el, "TargetMode") == "External":
                out.append({"part": name,
                            "type": (_attr(el, "Type") or "").rsplit("/", 1)[-1],
                            "target": _attr(el, "Target")})
    return out


def _ole_objects(zf):
    n = 0
    for name in zf.namelist():
        if not (name.endswith(".xml") and name.startswith("word/")):
            continue
        try:
            root = _read_root(zf, name)
        except etree.XMLSyntaxError:
            continue
        for el in root.iter():
            if _local(el.tag) in ("OLEObject", "objectEmbed", "objectLink"):
                n += 1
    return n


def diagnose(docx_path):
    with zipfile.ZipFile(docx_path) as zf:
        doc = _read_root(zf, "word/document.xml")
        complex_fields = _complex_fields(doc)
        simple_fields = _simple_fields(doc)
        uf = _update_fields_val(zf)
        ext_rels = _external_rels(zf)
        ole = _ole_objects(zf)

    all_fields = complex_fields + simple_fields
    dirty = [f for f in all_fields if f["dirty"]]
    external = [f for f in all_fields if f["kind"] == "external"]
    compiled = [f for f in all_fields if f["kind"] == "compiled"]

    sources = []
    if uf == "true":
        sources.append("settings.xml 的全局 updateFields=true（打开时强制更新全部域）")
    if dirty:
        sources.append("%d 个域带 w:dirty=true（如 TOC；打开时会被纳入更新而弹窗）" % len(dirty))
    if external:
        toks = sorted({f["token"] for f in external})
        sources.append("%d 个引用外部文件的域：%s" % (len(external), "、".join(toks)))
    if ext_rels:
        sources.append("%d 个外部关系（TargetMode=External，如链接图片/OLE）" % len(ext_rels))
    if ole:
        sources.append("%d 个 OLE 链接对象" % ole)

    return {
        "docx": docx_path,
        "update_fields": uf,
        "n_fields_total": len(all_fields),
        "n_dirty": len(dirty),
        "n_external_fields": len(external),
        "n_compiled_fields": len(compiled),
        "external_fields": external,
        "compiled_fields": compiled,
        "dirty_fields": dirty,
        "external_rels": ext_rels,
        "n_ole_objects": ole,
        "prompt_sources": sources,
    }


def main():
    if len(sys.argv) < 2:
        print(json.dumps({"status": "error",
                          "error": "usage: diagnose_fields.py <docx路径>"}))
        sys.exit(1)
    path = sys.argv[1]
    if not os.path.exists(path):
        print(json.dumps({"status": "error", "error": "file not found: %s" % path},
                         ensure_ascii=False))
        sys.exit(1)

    rep = diagnose(path)
    # human-readable summary to stderr-free stdout
    print("== 弹窗诊断：%s ==" % os.path.basename(path))
    print("settings.updateFields = %s" % rep["update_fields"])
    print("域总数=%d  其中 dirty=%d  外部引用域=%d  编译域(TOC等)=%d"
          % (rep["n_fields_total"], rep["n_dirty"],
             rep["n_external_fields"], rep["n_compiled_fields"]))
    print("外部关系(External)=%d  OLE链接对象=%d"
          % (len(rep["external_rels"]), rep["n_ole_objects"]))
    if rep["external_fields"]:
        print("-- 外部引用域明细 --")
        for f in rep["external_fields"][:20]:
            print("   %s | %s" % (f["token"], (f["instr"].strip())[:80]))
    if rep["external_rels"]:
        print("-- 外部关系明细 --")
        for r in rep["external_rels"][:20]:
            print("   %s | %s -> %s" % (r["part"], r["type"], r["target"]))
    print("")
    if rep["prompt_sources"]:
        print(">> 疑似弹窗来源：")
        for s in rep["prompt_sources"]:
            print("   - " + s)
    else:
        print(">> 未发现任何已知弹窗来源（updateFields=false、无 dirty 域、无外部"
              "引用域/外部关系/OLE）。若仍弹窗，请确认打开的确是这份文件。")
    print("")
    print("JSON: " + json.dumps(rep, ensure_ascii=False))


if __name__ == "__main__":
    main()
