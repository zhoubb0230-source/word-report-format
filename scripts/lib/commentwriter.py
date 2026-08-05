# -*- coding: utf-8 -*-
"""Self-contained Word comment writer (lxml only, no python-docx / defusedxml).

Vendored so the skill does not depend on any external docx-tool version.
Accumulates comments against paragraph elements in word/document.xml, then
`flush()` **merges** them into the comment satellite files (never overwrites —
see the CommentWriter docstring) and patches [Content_Types].xml and
word/_rels/document.xml.rels.

Author label for every comment is fixed to "XAgent" (skill requirement).

Marker layout (direct children of w:p, never inside w:r):
    <w:commentRangeStart w:id="N"/>  (inserted right after w:pPr)
    ...existing runs...
    <w:commentRangeEnd w:id="N"/>
    <w:r><w:rPr><w:rStyle w:val="CommentReference"/></w:rPr>
         <w:commentReference w:id="N"/></w:r>
"""
import os
import random
from datetime import datetime, timezone

from lxml import etree

from docxcommon import qn, W, W14, W15, W16CID, parse_xml

CT_NS = "http://schemas.openxmlformats.org/package/2006/content-types"
REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"

CT_MAP = {
    "comments.xml":
        "application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml",
    "commentsExtended.xml":
        "application/vnd.openxmlformats-officedocument.wordprocessingml.commentsExtended+xml",
    "commentsIds.xml":
        "application/vnd.openxmlformats-officedocument.wordprocessingml.commentsIds+xml",
}
REL_TYPE = {
    "comments.xml":
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments",
    "commentsExtended.xml":
        "http://schemas.microsoft.com/office/2011/relationships/commentsExtended",
    "commentsIds.xml":
        "http://schemas.microsoft.com/office/2016/relationships/commentsIds",
}


def _hex8():
    return "%08X" % random.randint(1, 0x7FFFFFFE)


MC_NS = "http://schemas.openxmlformats.org/markup-compatibility/2006"


def _existing_root(pkg_dir, relpath):
    """已有的批注部件根元素（没有/坏了就返回 None）。"""
    path = os.path.join(pkg_dir, relpath)
    if not os.path.exists(path):
        return None
    try:
        return parse_xml(path).getroot()
    except (OSError, etree.XMLSyntaxError):
        return None


def _reroot_with_ns(old_root, nsmap):
    """把已有根元素换成一个**补齐了命名空间声明**的同名根，子元素/属性原样搬过去。

    lxml 不能给已存在的元素补 nsmap，而合并进来的新批注要用 w14/w15/w16cid 前缀的属性；
    根上没有声明时 lxml 会把声明散落到每个子元素上，Word 对这种写法很挑剔。顺带把新前缀
    补进 `mc:Ignorable`（原本就有该属性时），保持与 Word 自己写出来的文件一致。"""
    merged = dict(old_root.nsmap or {})
    merged.update({k: v for k, v in nsmap.items() if merged.get(k) != v})
    if merged == dict(old_root.nsmap or {}):
        return old_root
    new = etree.Element(old_root.tag, nsmap=merged)
    for k, v in old_root.attrib.items():
        new.set(k, v)
    for child in list(old_root):
        new.append(child)
    ign = new.get("{%s}Ignorable" % MC_NS)
    if ign is not None:
        have = ign.split()
        for pfx in nsmap:
            if pfx not in have and pfx != "w":
                have.append(pfx)
        new.set("{%s}Ignorable" % MC_NS, " ".join(have))
    return new


class CommentWriter:
    """把批注**合并**进文档已有的批注里，绝不整体覆盖。

    覆盖写会造成两种确定的损坏（用户实测"发现无法读取的内容"的成因之一）：原文档自带的
    审阅批注被抹掉（那已经是"改原文"），而 document.xml 里它们的
    `commentRangeStart/commentReference` 还在，变成**悬空引用**；同时新批注从 id=0 开始
    分配，与残留的旧标记**撞 id**。所以这里先读一遍已有部件：新 id 从最大值 +1 起排，
    新批注追加到原有根元素上。"""

    def __init__(self, pkg_dir, author="XAgent"):
        self.pkg = pkg_dir
        self.author = author
        self.initials = author[:1].upper() if author else "X"
        self.records = []          # dicts: id, para_id, durable, text
        self._existing = _existing_root(pkg_dir, "word/comments.xml")
        self._next_id = self._first_free_id()

    def _first_free_id(self):
        """新批注 id 的起点：已有批注 id 与文档里已用过的标记 id 的最大值 +1。"""
        used = []
        for root, tag in ((self._existing, qn("w:comment")),
                          (_existing_root(self.pkg, "word/document.xml"), None)):
            if root is None:
                continue
            els = (root.iter(tag) if tag else
                   (e for e in root.iter() if e.tag in (qn("w:commentRangeStart"),
                                                        qn("w:commentRangeEnd"),
                                                        qn("w:commentReference"))))
            for el in els:
                try:
                    used.append(int(el.get(qn("w:id"))))
                except (TypeError, ValueError):
                    continue
        return (max(used) + 1) if used else 0

    # -- public ------------------------------------------------------------
    def add(self, p_el, text):
        """Wrap paragraph p_el with a comment range; queue the comment body."""
        cid = self._next_id
        self._next_id += 1
        para_id = _hex8()
        durable = _hex8()
        self._wrap(p_el, cid)
        self.records.append({"id": cid, "para_id": para_id,
                             "durable": durable, "text": text})
        return cid

    def flush(self):
        if not self.records:
            return
        self._write_comments()
        self._write_extended()
        self._write_ids()
        self._patch_content_types()
        self._patch_rels()

    # -- marker insertion --------------------------------------------------
    def _wrap(self, p, cid):
        start = etree.Element(qn("w:commentRangeStart"))
        start.set(qn("w:id"), str(cid))
        pPr = p.find(qn("w:pPr"))
        if pPr is not None:
            pPr.addnext(start)
        else:
            p.insert(0, start)

        end = etree.Element(qn("w:commentRangeEnd"))
        end.set(qn("w:id"), str(cid))
        p.append(end)

        r = etree.SubElement(p, qn("w:r"))
        rpr = etree.SubElement(r, qn("w:rPr"))
        rstyle = etree.SubElement(rpr, qn("w:rStyle"))
        rstyle.set(qn("w:val"), "CommentReference")
        ref = etree.SubElement(r, qn("w:commentReference"))
        ref.set(qn("w:id"), str(cid))

    # -- satellite writers -------------------------------------------------
    def _nsmap_comments(self):
        return {"w": W, "w14": W14, "w15": W15, "w16cid": W16CID}

    def _write_comments(self):
        root = (_reroot_with_ns(self._existing, {"w": W, "w14": W14})
                if self._existing is not None
                else etree.Element(qn("w:comments"), nsmap={"w": W, "w14": W14}))
        date = datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")
        for rec in self.records:
            c = etree.SubElement(root, qn("w:comment"))
            c.set(qn("w:id"), str(rec["id"]))
            c.set(qn("w:author"), self.author)
            c.set(qn("w:date"), date)
            c.set(qn("w:initials"), self.initials)
            p = etree.SubElement(c, qn("w:p"))
            p.set(qn("w14:paraId"), rec["para_id"])
            p.set(qn("w14:textId"), "77777777")
            # annotationRef run
            r0 = etree.SubElement(p, qn("w:r"))
            rpr0 = etree.SubElement(r0, qn("w:rPr"))
            rs = etree.SubElement(rpr0, qn("w:rStyle"))
            rs.set(qn("w:val"), "CommentReference")
            etree.SubElement(r0, qn("w:annotationRef"))
            # text run
            r1 = etree.SubElement(p, qn("w:r"))
            t = etree.SubElement(r1, qn("w:t"))
            t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
            t.text = rec["text"]
        self._write_tree(root, "word/comments.xml")

    def _write_extended(self):
        old = _existing_root(self.pkg, "word/commentsExtended.xml")
        root = (_reroot_with_ns(old, {"w15": W15}) if old is not None
                else etree.Element(qn("w15:commentsEx"), nsmap={"w15": W15}))
        for rec in self.records:
            e = etree.SubElement(root, qn("w15:commentEx"))
            e.set(qn("w15:paraId"), rec["para_id"])
            e.set(qn("w15:done"), "0")
        self._write_tree(root, "word/commentsExtended.xml")

    def _write_ids(self):
        old = _existing_root(self.pkg, "word/commentsIds.xml")
        root = (_reroot_with_ns(old, {"w16cid": W16CID}) if old is not None
                else etree.Element(qn("w16cid:commentsIds"), nsmap={"w16cid": W16CID}))
        for rec in self.records:
            e = etree.SubElement(root, qn("w16cid:commentId"))
            e.set(qn("w16cid:paraId"), rec["para_id"])
            e.set(qn("w16cid:durableId"), rec["durable"])
        self._write_tree(root, "word/commentsIds.xml")

    def _write_tree(self, root, relpath):
        path = os.path.join(self.pkg, relpath)
        os.makedirs(os.path.dirname(path), exist_ok=True)
        etree.ElementTree(root).write(
            path, xml_declaration=True, encoding="UTF-8", standalone=True)

    # -- content types + rels ---------------------------------------------
    def _patch_content_types(self):
        path = os.path.join(self.pkg, "[Content_Types].xml")
        tree = parse_xml(path)
        root = tree.getroot()
        existing = {ov.get("PartName")
                    for ov in root.findall("{%s}Override" % CT_NS)}
        for fname, ct in CT_MAP.items():
            part = "/word/%s" % fname
            if part in existing:
                continue
            ov = etree.SubElement(root, "{%s}Override" % CT_NS)
            ov.set("PartName", part)
            ov.set("ContentType", ct)
        tree.write(path, xml_declaration=True, encoding="UTF-8", standalone=True)

    def _patch_rels(self):
        path = os.path.join(self.pkg, "word", "_rels", "document.xml.rels")
        tree = parse_xml(path)
        root = tree.getroot()
        rels = root.findall("{%s}Relationship" % REL_NS)
        targets = {r.get("Target") for r in rels}
        maxid = 0
        for r in rels:
            rid = r.get("Id", "")
            if rid.startswith("rId") and rid[3:].isdigit():
                maxid = max(maxid, int(rid[3:]))
        for fname, rtype in REL_TYPE.items():
            if fname in targets:
                continue
            maxid += 1
            rel = etree.SubElement(root, "{%s}Relationship" % REL_NS)
            rel.set("Id", "rId%d" % maxid)
            rel.set("Type", rtype)
            rel.set("Target", fname)
        tree.write(path, xml_declaration=True, encoding="UTF-8", standalone=True)
