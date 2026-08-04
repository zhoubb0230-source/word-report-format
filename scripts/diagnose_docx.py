# -*- coding: utf-8 -*-
"""只读排查「Word 提示发现无法读取的内容 / 需要修复」的结构性成因。

用法：
    python3 scripts/diagnose_docx.py <任意 .docx>     # 原件、成品都能查，只读不改

为什么需要它：`45_validate_output.py` 只在流水线的 workdir 上跑，而"Word 打不开"往往
要拿**手上的那个文件**（可能是原件、可能是成品、也可能是用户手改过的）直接查。而且
「XML 良构」检查发现不了这类问题——**良构 ≠ 合法**，OOXML 的复杂类型几乎都是
`xsd:sequence`，子元素顺序错一处 Word 就判定文档损坏（陷阱 #14）。

查这几类（都是 Word 报损坏的常见成因，纯读、不修改）：
  1. **元素顺序**违反 schema sequence（`docxcommon.ELEMENT_ORDER` 登记的容器）。
  2. **样式名重复** —— Word 要求样式名在文档内唯一，重名即报损坏（陷阱 #18）；
     同时查 styleId 重复。
  3. **悬空引用** —— 段落 `pStyle` 指向不存在的样式；`numPr/numId` 指向 numbering.xml
     里没有的编号；`num → abstractNumId` 指向不存在的 abstractNum。
  4. **关系/内容类型** —— document.xml 里用到的 `r:id` 在 rels 里找不到；存在的部件
     没有对应的 Content-Type Override。
  5. **批注** —— 悬空的批注引用、重复的批注 id、不配对的 commentRangeStart/End。
  6. **编号身份链接** —— 两条 abstractNum 抢同一个 `styleLink`/`numStyleLink`，或同一个
     段落样式被两条编号定义的级别 `pStyle` 绑定。

输出单行 JSON（`ok` 为 false 时逐项列出问题）；退出码：0 干净 / 2 发现问题。
"""
import json
import os
import re
import sys
import zipfile

sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), "lib"))

from lxml import etree
from docxcommon import (qn, order_violations, local_name, comment_problems,
                        numbering_link_problems)

CT_NS = "http://schemas.openxmlformats.org/package/2006/content-types"
REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

_XML_PARTS = ("word/document.xml", "word/styles.xml", "word/numbering.xml",
              "word/settings.xml", "word/footnotes.xml", "word/endnotes.xml")


def _root(zf, name):
    parser = etree.XMLParser(remove_blank_text=False, huge_tree=True)
    with zf.open(name) as f:
        return etree.parse(f, parser).getroot()


def _check_order(zf, names, out):
    for part in _XML_PARTS:
        if part not in names:
            continue
        for parent, child, prev in order_violations(_root(zf, part)):
            out.append({"kind": "element_order", "part": part, "container": parent,
                        "element": child, "after": prev,
                        "hint": "OOXML 是 xsd:sequence，顺序错了 Word 会拒绝打开"})


def _check_styles(zf, names, out):
    if "word/styles.xml" not in names:
        return {}
    root = _root(zf, "word/styles.xml")
    by_name, by_id = {}, {}
    for st in root.findall(qn("w:style")):
        sid = st.get(qn("w:styleId"))
        by_id.setdefault(sid, []).append(sid)
        nm = st.find(qn("w:name"))
        if nm is not None and nm.get(qn("w:val")):
            by_name.setdefault(nm.get(qn("w:val")), []).append(sid)
    for nm, ids in by_name.items():
        if len(ids) > 1:
            out.append({"kind": "duplicate_style_name", "name": nm, "styleIds": ids,
                        "hint": "Word 要求样式名唯一，重名会报“发现无法读取的内容”"})
    for sid, ids in by_id.items():
        if len(ids) > 1:
            out.append({"kind": "duplicate_style_id", "styleId": sid})
    return by_id


def _check_numbering(zf, names, out):
    nums, abstracts = set(), set()
    if "word/numbering.xml" in names:
        root = _root(zf, "word/numbering.xml")
        abstracts = {a.get(qn("w:abstractNumId"))
                     for a in root.findall(qn("w:abstractNum"))}
        for num in root.findall(qn("w:num")):
            nums.add(num.get(qn("w:numId")))
            a = num.find(qn("w:abstractNumId"))
            if a is not None and a.get(qn("w:val")) not in abstracts:
                out.append({"kind": "dangling_abstract_num",
                            "numId": num.get(qn("w:numId")),
                            "abstractNumId": a.get(qn("w:val"))})
        out.extend(numbering_link_problems(root))
    return nums


def _check_comments(zf, names, out):
    """批注：悬空引用 / 重复 id / 区间不配对（覆盖式写批注留下的典型残骸）。"""
    if "word/document.xml" not in names:
        return
    comments = (_root(zf, "word/comments.xml")
                if "word/comments.xml" in names else None)
    out.extend(comment_problems(_root(zf, "word/document.xml"), comments))


def _check_references(zf, names, style_ids, num_ids, out):
    if "word/document.xml" not in names:
        return
    root = _root(zf, "word/document.xml")
    bad_styles, bad_nums = set(), set()
    for ps in root.iter(qn("w:pStyle")):
        v = ps.get(qn("w:val"))
        if style_ids and v not in style_ids:
            bad_styles.add(v)
    for ni in root.iter(qn("w:numId")):
        v = ni.get(qn("w:val"))
        if v and v != "0" and num_ids and v not in num_ids:
            bad_nums.add(v)
    if bad_styles:
        out.append({"kind": "dangling_pstyle", "styleIds": sorted(bad_styles),
                    "hint": "段落引用了 styles.xml 里不存在的样式"})
    if bad_nums:
        out.append({"kind": "dangling_numid", "numIds": sorted(bad_nums),
                    "hint": "段落引用了 numbering.xml 里不存在的编号"})

    rels_path = "word/_rels/document.xml.rels"
    if rels_path in names:
        rels = {r.get("Id") for r in _root(zf, rels_path)}
        used = set()
        for el in root.iter():
            for k, v in el.attrib.items():
                if k.startswith("{%s}" % R_NS) and v:
                    used.add(v)
        missing = sorted(used - rels)
        if missing:
            out.append({"kind": "missing_relationship", "rIds": missing,
                        "hint": "document.xml 用到的关系 id 在 rels 里不存在"})


def _check_content_types(zf, names, out):
    if "[Content_Types].xml" not in names:
        out.append({"kind": "missing_content_types"})
        return
    root = _root(zf, "[Content_Types].xml")
    overrides = {ov.get("PartName") for ov in root.findall("{%s}Override" % CT_NS)}
    defaults = {d.get("Extension", "").lower()
                for d in root.findall("{%s}Default" % CT_NS)}
    for part in _XML_PARTS:
        if part in names and ("/" + part) not in overrides and "xml" not in defaults:
            out.append({"kind": "missing_content_type", "part": part})


def diagnose(path):
    problems = []
    try:
        zf = zipfile.ZipFile(path)
    except zipfile.BadZipFile as e:
        return {"status": "error", "ok": False, "file": path,
                "problems": [{"kind": "bad_zip", "error": str(e)}]}
    with zf:
        bad = zf.testzip()
        if bad is not None:
            problems.append({"kind": "bad_crc", "part": bad})
        names = set(zf.namelist())
        for name in names:
            if not (name.endswith(".xml") or name.endswith(".rels")):
                continue
            try:
                _root(zf, name)
            except etree.XMLSyntaxError as e:
                problems.append({"kind": "malformed_xml", "part": name,
                                 "error": str(e)})
        _check_order(zf, names, problems)
        style_ids = set(_check_styles(zf, names, problems) or ())
        num_ids = _check_numbering(zf, names, problems)
        _check_references(zf, names, style_ids, num_ids, problems)
        _check_comments(zf, names, problems)
        _check_content_types(zf, names, problems)
    return {"status": "ok" if not problems else "error", "ok": not problems,
            "file": path, "n_problems": len(problems), "problems": problems[:80]}


def main():
    if len(sys.argv) < 2:
        print(json.dumps({"status": "error",
                          "error": "usage: diagnose_docx.py <docx>"}))
        sys.exit(1)
    report = diagnose(sys.argv[1])
    print(json.dumps(report, ensure_ascii=False, indent=1))
    sys.exit(0 if report["ok"] else 2)


if __name__ == "__main__":
    main()
