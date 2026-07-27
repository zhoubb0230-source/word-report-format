# -*- coding: utf-8 -*-
"""
docxcommon.py  —  Self-contained DOCX helpers for the format-review skill.

VENDORED ON PURPOSE. This module does not import python-docx and does not rely
on any other skill, because the runtime environment may ship a different
python-docx version than expected. It only needs the Python stdlib + lxml.

It provides:
  * WordprocessingML namespaces + qn() qualified-name helper
  * safe unzip / rezip of .docx packages
  * unit conversions (cm->twips, Chinese font size -> half points)
  * a SINGLE canonical body-paragraph iterator (iter_body_paragraphs) that
    BOTH extraction and apply use, so paragraph indices always line up
  * a style-inheritance resolver that computes the EFFECTIVE paragraph/run
    properties (docDefaults -> style basedOn chain -> direct pPr/rPr).

Nothing here calls a language model. All logic is deterministic.
"""

import os
import re
import shutil
import zipfile
from lxml import etree

# ---------------------------------------------------------------------------
# Namespaces
# ---------------------------------------------------------------------------
W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
W14 = "http://schemas.microsoft.com/office/word/2010/wordml"
W15 = "http://schemas.microsoft.com/office/word/2012/wordml"
W16CID = "http://schemas.microsoft.com/office/word/2016/wordml/cid"
W16CEX = "http://schemas.microsoft.com/office/word/2018/wordml/cex"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
CT = "http://schemas.openxmlformats.org/package/2006/content-types"
PR = "http://schemas.openxmlformats.org/package/2006/relationships"

NSMAP = {"w": W, "w14": W14, "w15": W15, "w16cid": W16CID, "w16cex": W16CEX, "r": R}


def qn(tag):
    """'w:pPr' -> '{namespace}pPr'."""
    pfx, local = tag.split(":", 1)
    ns = {"w": W, "w14": W14, "w15": W15, "w16cid": W16CID, "w16cex": W16CEX, "r": R}[pfx]
    return "{%s}%s" % (ns, local)


def parse_xml(path):
    parser = etree.XMLParser(remove_blank_text=False, huge_tree=True)
    return etree.parse(path, parser)


# ---------------------------------------------------------------------------
# Package (zip) helpers  — never mutate the caller's original file
# ---------------------------------------------------------------------------
def unzip_docx(docx_path, dest_dir):
    """Extract a .docx into dest_dir. Strips symlink entries (untrusted input)."""
    os.makedirs(dest_dir, exist_ok=True)
    with zipfile.ZipFile(docx_path) as zf:
        for info in zf.infolist():
            # skip absolute paths / traversal
            name = info.filename
            if name.startswith("/") or ".." in name.split("/"):
                continue
            # skip symlinks
            if (info.external_attr >> 16) & 0o170000 == 0o120000:
                continue
            target = os.path.join(dest_dir, name)
            if name.endswith("/"):
                os.makedirs(target, exist_ok=True)
                continue
            os.makedirs(os.path.dirname(target), exist_ok=True)
            with zf.open(info) as src, open(target, "wb") as out:
                shutil.copyfileobj(src, out)
    return dest_dir


def rezip_docx(src_dir, out_path):
    """Zip an unpacked package back into a .docx. [Content_Types].xml first."""
    if os.path.exists(out_path):
        os.remove(out_path)
    names = []
    for root, _dirs, files in os.walk(src_dir):
        for f in files:
            full = os.path.join(root, f)
            rel = os.path.relpath(full, src_dir).replace(os.sep, "/")
            names.append(rel)
    names.sort(key=lambda n: (n != "[Content_Types].xml", n))
    with zipfile.ZipFile(out_path, "w", zipfile.ZIP_DEFLATED) as zf:
        for rel in names:
            zf.write(os.path.join(src_dir, rel), rel)
    return out_path


# ---------------------------------------------------------------------------
# Unit conversions
# ---------------------------------------------------------------------------
def cm_to_twips(cm):
    return int(round(cm * 1440.0 / 2.54))


# Chinese named sizes -> points
CN_SIZE_PT = {
    "初号": 42, "小初": 36, "一号": 26, "小一": 24, "二号": 22, "小二": 18,
    "三号": 16, "小三": 15, "四号": 14, "小四": 12, "五号": 10.5, "小五": 9,
    "六号": 7.5, "小六": 6.5, "七号": 5.5, "八号": 5,
}


def pt_to_halfpt(pt):
    return int(round(pt * 2))


def name_to_halfpt(name):
    """'三号' -> 32 (half points). Accepts a number (pt) too."""
    if name in CN_SIZE_PT:
        return pt_to_halfpt(CN_SIZE_PT[name])
    return pt_to_halfpt(float(name))


# ---------------------------------------------------------------------------
# Canonical body-paragraph iterator
# ---------------------------------------------------------------------------
def in_textbox(el, stop_at=None):
    """True if el is nested inside a <w:txbxContent> (a textbox), walking up the
    ancestor chain until stop_at (exclusive) or the root.

    Textbox content is not part of the linear reading flow and must not be
    treated as body paragraphs, nor restyled as a side effect of editing the
    host paragraph. Shared by iter_body_paragraphs (stop_at=None, walk to root)
    and the run iterators (stop_at=p, only look above the paragraph)."""
    anc = el.getparent()
    while anc is not None and anc is not stop_at:
        if anc.tag == qn("w:txbxContent"):
            return True
        anc = anc.getparent()
    return False


# BOTH extraction and apply MUST use this so that "para_index" refers to the
# same physical <w:p> in both passes. It walks every <w:p> that is a descendant
# of <w:body> in document order, including paragraphs inside tables, but skips
# paragraphs that live inside a textbox (w:txbxContent) because those are not
# part of the linear reading flow and are rarely subject to these rules.
def iter_body_paragraphs(doc_root):
    body = doc_root.find(qn("w:body"))
    if body is None:
        return
    idx = 0
    for p in body.iter(qn("w:p")):
        if in_textbox(p):  # skip textbox paragraphs
            continue
        yield idx, p
        idx += 1


def para_text(p):
    """Full visible text of a paragraph (concatenate all w:t descendants)."""
    parts = []
    for t in p.iter(qn("w:t")):
        parts.append(t.text or "")
    return "".join(parts)


def in_table(p):
    anc = p.getparent()
    while anc is not None:
        if anc.tag == qn("w:tbl"):
            return True
        anc = anc.getparent()
    return False


# ---------------------------------------------------------------------------
# Style resolver: compute EFFECTIVE properties following inheritance.
# ---------------------------------------------------------------------------
class StyleResolver:
    """
    Resolves effective paragraph + run properties by merging, in order:
        docDefaults (rPrDefault / pPrDefault)
        -> paragraph style chain (basedOn, resolved deepest-first)
        -> direct paragraph pPr and the paragraph-mark rPr
        -> (per run) direct run rPr
    Returns plain dicts of the specific properties this skill cares about.
    """

    def __init__(self, styles_root):
        self.styles = {}            # styleId -> <w:style>
        self.default_para_style = None
        self.doc_rpr = {}
        self.doc_ppr = {}
        if styles_root is not None:
            self._load(styles_root)

    def _load(self, root):
        docdef = root.find(qn("w:docDefaults"))
        if docdef is not None:
            rprd = docdef.find(qn("w:rPrDefault"))
            if rprd is not None:
                rpr = rprd.find(qn("w:rPr"))
                if rpr is not None:
                    self.doc_rpr = read_rpr(rpr)
            pprd = docdef.find(qn("w:pPrDefault"))
            if pprd is not None:
                ppr = pprd.find(qn("w:pPr"))
                if ppr is not None:
                    self.doc_ppr = read_ppr(ppr)
        for st in root.findall(qn("w:style")):
            if st.get(qn("w:type")) != "paragraph":
                continue
            sid = st.get(qn("w:styleId"))
            if sid:
                self.styles[sid] = st
            if st.get(qn("w:default")) == "1":
                self.default_para_style = sid

    def _style_chain(self, style_id):
        """Return list of style elements from root ancestor down to style_id."""
        chain = []
        seen = set()
        cur = style_id
        while cur and cur in self.styles and cur not in seen:
            seen.add(cur)
            st = self.styles[cur]
            chain.append(st)
            based = st.find(qn("w:basedOn"))
            cur = based.get(qn("w:val")) if based is not None else None
        chain.reverse()  # root-most first
        return chain

    def resolve(self, style_id, direct_ppr_el, mark_rpr_el):
        """
        Return (ppr_dict, rpr_dict) of effective paragraph-level properties.
        rpr_dict is the paragraph-mark / style run defaults (the baseline the
        paragraph's runs inherit). Per-run overrides are merged separately.
        """
        ppr = dict(self.doc_ppr)
        rpr = dict(self.doc_rpr)
        sid = style_id or self.default_para_style
        for st in self._style_chain(sid):
            spr = st.find(qn("w:pPr"))
            if spr is not None:
                _merge(ppr, read_ppr(spr))
            srp = st.find(qn("w:rPr"))
            if srp is not None:
                _merge(rpr, read_rpr(srp))
        if direct_ppr_el is not None:
            _merge(ppr, read_ppr(direct_ppr_el))
            mrp = direct_ppr_el.find(qn("w:rPr"))
            if mrp is not None:
                _merge(rpr, read_rpr(mrp))
        if mark_rpr_el is not None:
            _merge(rpr, read_rpr(mark_rpr_el))
        return ppr, rpr

    def resolve_cascade(self, style_id, direct_ppr_el, mark_rpr_el,
                        num_levels=None):
        """Like ``resolve`` but folds the NUMBERING layer into the cascade.

        The real WordprocessingML precedence for a numbered paragraph's
        indentation is::

            docDefaults  <  paragraph style chain  <  numbering level  <  direct

        ``resolve`` above stops at the style chain and never consults the
        numbering level, so an indent (very often a ``w:hanging``) that a
        heading inherits from its list level is INVISIBLE to it — the paragraph
        looks compliant while Word renders a hanging indent. This method inserts
        the numbering level's ``w:pPr`` between the style chain and the direct
        pPr, at the correct precedence (it OVERRIDES the style chain but is
        OVERRIDDEN by any direct value), so the returned ppr reflects the indent
        the reader actually sees. That is the "把编号层折进有效值" requirement
        from the 方案C 契约 (§2.3) and the foundation the collapse-invariant
        check builds on.

        ``num_levels`` is the map returned by :func:`load_numbering_levels`
        (numId -> {ilvl: pPr_dict}). When it is None/empty this degrades to
        exactly ``resolve``'s result, so callers without a numbering.xml are
        unaffected. Only the paragraph pPr is folded (that is where a list level
        carries its indent); the numbering level's own rPr is not merged — a
        deliberate, documented limitation kept to match
        ``load_numbering_levels``.

        The winning numId/ilvl is computed with the same precedence: a direct
        ``w:numPr`` beats one inherited from the style chain. numId "0" is the
        OOXML "no numbering" override and disables the fold.
        """
        # 1) docDefaults + full style chain (no direct yet).
        ppr = dict(self.doc_ppr)
        rpr = dict(self.doc_rpr)
        sid = style_id or self.default_para_style
        for st in self._style_chain(sid):
            spr = st.find(qn("w:pPr"))
            if spr is not None:
                _merge(ppr, read_ppr(spr))
            srp = st.find(qn("w:rPr"))
            if srp is not None:
                _merge(rpr, read_rpr(srp))

        # 2) Resolve the effective numbering reference: a direct numPr wins over
        #    one inherited through the style chain (already in ppr from step 1).
        direct_ppr = read_ppr(direct_ppr_el) if direct_ppr_el is not None else {}
        num_id = direct_ppr.get("num_id")
        if num_id is None:
            num_id = ppr.get("num_id")
        ilvl = direct_ppr.get("ilvl")
        if ilvl is None:
            ilvl = ppr.get("ilvl")

        # 3) Fold the numbering level's pPr in — over the style chain, under the
        #    direct pPr applied next.
        if num_levels and num_id and num_id != "0":
            try:
                il = int(ilvl) if ilvl is not None else 0
            except (ValueError, TypeError):
                il = 0
            lvl_ppr = num_levels.get(num_id, {}).get(il)
            if lvl_ppr:
                _merge(ppr, lvl_ppr)

        # 4) Direct paragraph pPr + paragraph-mark rPr (highest precedence).
        if direct_ppr_el is not None:
            _merge(ppr, direct_ppr)
            mrp = direct_ppr_el.find(qn("w:rPr"))
            if mrp is not None:
                _merge(rpr, read_rpr(mrp))
        if mark_rpr_el is not None:
            _merge(rpr, read_rpr(mark_rpr_el))
        return ppr, rpr


def _merge(base, extra):
    """Overlay non-None values from extra onto base."""
    for k, v in extra.items():
        if v is not None:
            base[k] = v


# --- readers: pull only the fields we care about ---------------------------
def read_rpr(rpr):
    d = {"east_asia": None, "ascii": None, "hansi": None, "size_hp": None,
         "bold": None}
    if rpr is None:
        return d
    rf = rpr.find(qn("w:rFonts"))
    if rf is not None:
        d["east_asia"] = rf.get(qn("w:eastAsia"))
        d["ascii"] = rf.get(qn("w:ascii"))
        d["hansi"] = rf.get(qn("w:hAnsi"))
    sz = rpr.find(qn("w:sz"))
    if sz is not None and sz.get(qn("w:val")):
        try:
            d["size_hp"] = int(sz.get(qn("w:val")))
        except ValueError:
            pass
    b = rpr.find(qn("w:b"))
    if b is not None:
        d["bold"] = b.get(qn("w:val")) not in ("0", "false")
    return d


def read_ppr(ppr):
    d = {"jc": None, "outline": None,
         "line": None, "line_rule": None,
         "space_before": None, "space_after": None,
         "space_before_lines": None, "space_after_lines": None,
         "first_line_chars": None, "first_line": None,
         "left_chars": None, "left": None,
         "start_chars": None, "start": None,
         "right_chars": None, "right": None,
         "end_chars": None, "end": None,
         "hanging_chars": None, "hanging": None,
         "num_id": None, "ilvl": None,
         "style_id": None}
    if ppr is None:
        return d
    ps = ppr.find(qn("w:pStyle"))
    if ps is not None:
        d["style_id"] = ps.get(qn("w:val"))
    npr = ppr.find(qn("w:numPr"))
    if npr is not None:
        ni = npr.find(qn("w:numId"))
        il = npr.find(qn("w:ilvl"))
        if ni is not None and ni.get(qn("w:val")) is not None:
            d["num_id"] = ni.get(qn("w:val"))
        if il is not None and il.get(qn("w:val")) is not None:
            d["ilvl"] = il.get(qn("w:val"))
    jc = ppr.find(qn("w:jc"))
    if jc is not None:
        d["jc"] = jc.get(qn("w:val"))
    ol = ppr.find(qn("w:outlineLvl"))
    if ol is not None and ol.get(qn("w:val")) is not None:
        try:
            d["outline"] = int(ol.get(qn("w:val")))
        except ValueError:
            pass
    sp = ppr.find(qn("w:spacing"))
    if sp is not None:
        if sp.get(qn("w:line")) is not None:
            try:
                d["line"] = int(sp.get(qn("w:line")))
            except ValueError:
                pass
        d["line_rule"] = sp.get(qn("w:lineRule"))
        for attr, key in (("before", "space_before"), ("after", "space_after"),
                          ("beforeLines", "space_before_lines"),
                          ("afterLines", "space_after_lines")):
            v = sp.get(qn("w:" + attr))
            if v is not None:
                try:
                    d[key] = int(v)
                except ValueError:
                    d[key] = v
    ind = ppr.find(qn("w:ind"))
    if ind is not None:
        for attr, key in (("firstLineChars", "first_line_chars"),
                          ("firstLine", "first_line"),
                          ("leftChars", "left_chars"),
                          ("left", "left"),
                          ("startChars", "start_chars"),
                          ("start", "start"),
                          ("rightChars", "right_chars"),
                          ("right", "right"),
                          ("endChars", "end_chars"),
                          ("end", "end"),
                          ("hangingChars", "hanging_chars"),
                          ("hanging", "hanging")):
            v = ind.get(qn("w:" + attr))
            if v is not None:
                try:
                    d[key] = int(v)
                except ValueError:
                    d[key] = v
    return d


def load_numbering_levels(numbering_root):
    """Map numId -> {ilvl(int): pPr_dict} from numbering.xml, so a paragraph
    that gets its number (and often its indent) from an automatic list can be
    resolved. Only the per-level w:pPr is read (that's where the list's indent
    lives). numId -> abstractNumId -> level pPr. w:lvlOverride is ignored
    (rare; the base level's indent is a good enough signal for our purpose)."""
    if numbering_root is None:
        return {}
    abstract = {}   # abstractNumId -> {ilvl: pPr_dict}
    for anum in numbering_root.findall(qn("w:abstractNum")):
        aid = anum.get(qn("w:abstractNumId"))
        levels = {}
        for lvl in anum.findall(qn("w:lvl")):
            il = lvl.get(qn("w:ilvl"))
            ppr_el = lvl.find(qn("w:pPr"))
            if il is not None and ppr_el is not None:
                try:
                    levels[int(il)] = read_ppr(ppr_el)
                except ValueError:
                    pass
        abstract[aid] = levels
    out = {}
    for num in numbering_root.findall(qn("w:num")):
        nid = num.get(qn("w:numId"))
        a = num.find(qn("w:abstractNumId"))
        if nid is not None and a is not None:
            out[nid] = abstract.get(a.get(qn("w:val")), {})
    return out


# Indentation keys shared by extraction / numbering-fallback / apply.
INDENT_KEYS = ("first_line_chars", "first_line", "left_chars", "left",
               "start_chars", "start", "right_chars", "right",
               "end_chars", "end", "hanging_chars", "hanging")


def get_pPr(p):
    return p.find(qn("w:pPr"))


def get_style_id(p):
    ppr = get_pPr(p)
    if ppr is None:
        return None
    ps = ppr.find(qn("w:pStyle"))
    return ps.get(qn("w:val")) if ps is not None else None


def get_mark_rpr(p):
    ppr = get_pPr(p)
    if ppr is None:
        return None
    return ppr.find(qn("w:rPr"))


def iter_text_runs(p):
    """Yield (run_el, text) for runs that carry visible text (skip markers).

    Recurses into wrapper elements (w:hyperlink, w:ins, w:smartTag, ...) via
    p.iter() rather than only direct children — a shallow findall() would see
    ZERO runs for a paragraph whose entire text is wrapped in <w:hyperlink>,
    which is exactly how Word emits every auto-generated TOC entry. That
    previously made font/size detection blind to TOC (and any hyperlinked
    text), silently falling back to the style baseline. para_text() and the
    apply-stage run editor already recurse this way; this now matches them.
    Runs inside a nested textbox (w:txbxContent) are excluded since they
    belong to a different logical paragraph.
    """
    for run in p.iter(qn("w:r")):
        if in_textbox(run, stop_at=p):
            continue
        txt = "".join((t.text or "") for t in run.findall(qn("w:t")))
        if txt:
            yield run, txt


def run_effective_rpr(run, baseline):
    """Merge baseline paragraph rPr with this run's direct rPr."""
    eff = dict(baseline)
    rpr = run.find(qn("w:rPr"))
    if rpr is not None:
        _merge(eff, read_rpr(rpr))
    return eff


def char_styles_overriding(styles_root):
    """styleId 集合：那些**字符样式**（`w:type="character"`）设置了 canonical 段落样式
    要管的属性（字体 / 字号 / 加粗），含经 basedOn 继承来的。

    为什么需要：字符样式挂在 run 上，在 cascade 里**压过段落样式**。一段标题若有几个
    run 带这种字符样式，指派 canonical 标题样式后那几个 run 依然我行我素——用户看到的
    就是"同一个三级标题前半部分加粗、后半部分不加粗，样式好像只应用到了后半部分"。

    `40_apply_fixes.py`（指派时摘掉这些 run 上的 `w:rStyle` 引用）与
    `45_validate_output.py`（坍缩不变量把它算作泄漏）**共用这一份判断**，免得两边对
    "哪些字符样式算抢戏"各有一套、彼此打架。只设颜色/下划线之类的字符样式（超链接、
    批注引用）不在集合里，指派不会动它们。"""
    if styles_root is None:
        return frozenset()
    by_id, based = {}, {}
    for st in styles_root.findall(qn("w:style")):
        if st.get(qn("w:type")) != "character":
            continue
        sid = st.get(qn("w:styleId"))
        if not sid:
            continue
        by_id[sid] = st
        b = st.find(qn("w:basedOn"))
        based[sid] = b.get(qn("w:val")) if b is not None else None

    def sets_governed(sid, seen):
        if sid in seen or sid not in by_id:
            return False
        seen.add(sid)
        rpr = by_id[sid].find(qn("w:rPr"))
        if rpr is not None:
            for tag in ("w:rFonts", "w:sz", "w:szCs", "w:b", "w:bCs"):
                if rpr.find(qn(tag)) is not None:
                    return True
        parent = based.get(sid)
        return sets_governed(parent, seen) if parent else False

    return frozenset(sid for sid in by_id if sets_governed(sid, set()))
