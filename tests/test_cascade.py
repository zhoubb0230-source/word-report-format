# -*- coding: utf-8 -*-
"""对抗性合成 fixtures：把同一竞争值（悬挂缩进/左缩进/字号）分别塞进
样式层 / 编号层 / 直接层，断言统一 cascade resolver 折出**正确优先级**下的有效值：

    docDefaults  <  段落样式链  <  编号层  <  直接属性

这是方案C 契约 §2.1/§2.3 的机器可判化：检测必须"把编号层折进有效值"，否则
把"hanging 藏编号层"的标题误判成合规。同时对照旧 `resolve()`（漏看编号层）以
锁住二者差异，防止有人把新 resolver 悄悄退回旧行为。

仅 stdlib + lxml。手搓最小 styles/numbering，参考 tests/helpers 的 build 思路。"""
import os
import sys
import unittest
from lxml import etree

HERE = os.path.dirname(os.path.abspath(__file__))
LIB = os.path.join(os.path.dirname(HERE), "scripts", "lib")
if LIB not in sys.path:
    sys.path.insert(0, LIB)

from docxcommon import (  # noqa: E402
    StyleResolver, load_numbering_levels, qn,
)
import cascade  # noqa: E402

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"


def styles(*style_blocks, doc_ppr="", doc_rpr=""):
    docdef = ""
    if doc_ppr or doc_rpr:
        docdef = ("<w:docDefaults>"
                  "<w:rPrDefault><w:rPr>%s</w:rPr></w:rPrDefault>"
                  "<w:pPrDefault><w:pPr>%s</w:pPr></w:pPrDefault>"
                  "</w:docDefaults>" % (doc_rpr, doc_ppr))
    return etree.fromstring(
        ('<w:styles xmlns:w="%s">%s%s</w:styles>'
         % (W, docdef, "".join(style_blocks))).encode("utf-8"))


def para_style(style_id, ppr_inner="", rpr_inner="", based_on=None, name=None):
    based = '<w:basedOn w:val="%s"/>' % based_on if based_on else ""
    nm = '<w:name w:val="%s"/>' % name if name else ""
    return ('<w:style w:type="paragraph" w:styleId="%s">%s%s'
            '<w:pPr>%s</w:pPr><w:rPr>%s</w:rPr></w:style>'
            % (style_id, nm, based, ppr_inner, rpr_inner))


def numbering(num_id, abstract_id, levels):
    """levels: {ilvl: ppr_inner_xml}."""
    lvls = "".join(
        '<w:lvl w:ilvl="%d"><w:pPr>%s</w:pPr></w:lvl>' % (il, inner)
        for il, inner in sorted(levels.items()))
    return etree.fromstring(
        ('<w:numbering xmlns:w="%s">'
         '<w:abstractNum w:abstractNumId="%s">%s</w:abstractNum>'
         '<w:num w:numId="%s"><w:abstractNumId w:val="%s"/></w:num>'
         '</w:numbering>'
         % (W, abstract_id, lvls, num_id, abstract_id)).encode("utf-8"))


def ppr_el(inner):
    return etree.fromstring(('<w:pPr xmlns:w="%s">%s</w:pPr>' % (W, inner)).encode("utf-8"))


class TestNumberingFoldedIntoCascade(unittest.TestCase):
    """编号层必须被折进有效值，且优先级正确。"""

    def test_numbering_hanging_beats_style_hanging(self):
        # 样式层 hanging=200，编号层 hanging=420，段落套样式并直接引用该列表。
        st = styles(para_style("H", ppr_inner='<w:ind w:hanging="200"/>'))
        num = numbering("5", "7", {0: '<w:ind w:hanging="420"/>'})
        levels = load_numbering_levels(num)
        resolver = StyleResolver(st)
        direct = ppr_el('<w:pStyle w:val="H"/>'
                        '<w:numPr><w:numId w:val="5"/><w:ilvl w:val="0"/></w:numPr>')

        # 旧 resolve：漏看编号层 → 看到的是样式层的 200（会误判"合规"）。
        old_ppr, _ = resolver.resolve("H", direct, None)
        self.assertEqual(old_ppr["hanging"], 200)

        # 新 resolve_cascade：编号层压过样式层 → 有效 hanging=420。
        ppr, _ = resolver.resolve_cascade("H", direct, None, levels)
        self.assertEqual(ppr["hanging"], 420)

    def test_direct_hanging_beats_numbering(self):
        st = styles(para_style("H", ppr_inner='<w:ind w:hanging="200"/>'))
        num = numbering("5", "7", {0: '<w:ind w:hanging="420"/>'})
        levels = load_numbering_levels(num)
        resolver = StyleResolver(st)
        direct = ppr_el('<w:pStyle w:val="H"/>'
                        '<w:numPr><w:numId w:val="5"/><w:ilvl w:val="0"/></w:numPr>'
                        '<w:ind w:hanging="640"/>')
        ppr, _ = resolver.resolve_cascade("H", direct, None, levels)
        self.assertEqual(ppr["hanging"], 640)   # 直接层最高优先级

    def test_numId_zero_disables_fold(self):
        # numId=0 是 OOXML "无编号"覆盖：即便编号定义里有 hanging，也不应折入。
        st = styles(para_style("H", ppr_inner='<w:ind w:firstLineChars="200"/>'))
        num = numbering("5", "7", {0: '<w:ind w:hanging="420"/>'})
        levels = load_numbering_levels(num)
        resolver = StyleResolver(st)
        direct = ppr_el('<w:pStyle w:val="H"/>'
                        '<w:numPr><w:numId w:val="0"/></w:numPr>')
        ppr, _ = resolver.resolve_cascade("H", direct, None, levels)
        self.assertIsNone(ppr.get("hanging"))
        self.assertEqual(ppr["first_line_chars"], 200)

    def test_num_levels_none_degrades_to_resolve(self):
        # 不传编号层时，行为必须与旧 resolve 完全一致（无 numbering.xml 的输入不受影响）。
        st = styles(para_style("H", ppr_inner='<w:ind w:hanging="200"/>'))
        resolver = StyleResolver(st)
        direct = ppr_el('<w:pStyle w:val="H"/>')
        a_ppr, a_rpr = resolver.resolve("H", direct, None)
        b_ppr, b_rpr = resolver.resolve_cascade("H", direct, None, None)
        self.assertEqual(a_ppr, b_ppr)
        self.assertEqual(a_rpr, b_rpr)

    def test_numbering_inherited_via_style_chain(self):
        # numPr 由样式链继承（不在直接 pPr 上）时，编号层仍应折入。
        st = styles(para_style(
            "H", ppr_inner='<w:numPr><w:numId w:val="5"/><w:ilvl w:val="0"/></w:numPr>'))
        num = numbering("5", "7", {0: '<w:ind w:hanging="420"/>'})
        levels = load_numbering_levels(num)
        resolver = StyleResolver(st)
        direct = ppr_el('<w:pStyle w:val="H"/>')
        ppr, _ = resolver.resolve_cascade("H", direct, None, levels)
        self.assertEqual(ppr["hanging"], 420)


class TestProvenanceReport(unittest.TestCase):
    """provenance = 方案C 的量尺：如实报告每个缩进键的有效值由哪一层供给。"""

    def test_each_layer_owns_its_contested_key(self):
        # docDefaults 供 left、样式链供 firstLineChars、编号层供 hanging、直接供 leftChars。
        st = styles(
            para_style("H", ppr_inner='<w:ind w:firstLineChars="200"/>'),
            doc_ppr='<w:ind w:left="100"/>')
        num = numbering("5", "7", {0: '<w:ind w:hanging="420"/>'})
        levels = load_numbering_levels(num)
        resolver = StyleResolver(st)
        direct = ppr_el('<w:pStyle w:val="H"/>'
                        '<w:numPr><w:numId w:val="5"/><w:ilvl w:val="0"/></w:numPr>'
                        '<w:ind w:leftChars="0"/>')
        report = cascade.contested_indent_report(resolver, "H", direct, None, levels)
        self.assertEqual(report.get("left"), "docDefaults")
        self.assertEqual(report.get("first_line_chars"), "style")
        self.assertEqual(report.get("hanging"), "numbering")
        self.assertEqual(report.get("leftChars") or report.get("left_chars"), "direct")

    def test_direct_override_shifts_ownership(self):
        # 编号层给 hanging，但直接层也给 hanging → 归属应记为 direct（钳干净的信号）。
        st = styles(para_style("H"))
        num = numbering("5", "7", {0: '<w:ind w:hanging="420"/>'})
        levels = load_numbering_levels(num)
        resolver = StyleResolver(st)
        direct = ppr_el('<w:pStyle w:val="H"/>'
                        '<w:numPr><w:numId w:val="5"/><w:ilvl w:val="0"/></w:numPr>'
                        '<w:ind w:hanging="0"/>')
        report = cascade.contested_indent_report(resolver, "H", direct, None, levels)
        self.assertEqual(report.get("hanging"), "direct")


class TestFontSizeCascade(unittest.TestCase):
    """字号竞争（样式层 vs 直接层）；编号层 rPr 不折入（已知取舍）。"""

    def test_direct_size_beats_style(self):
        st = styles(para_style("H", rpr_inner='<w:sz w:val="32"/>'))
        resolver = StyleResolver(st)
        # 段落标记 rPr 直接给 24（小三），压过样式链的 32（三号）。
        direct = ppr_el('<w:pStyle w:val="H"/><w:rPr><w:sz w:val="24"/></w:rPr>')
        _, rpr = resolver.resolve_cascade("H", direct, None, {})
        self.assertEqual(rpr["size_hp"], 24)

    def test_style_chain_basedon_resolves_deepest_first(self):
        # basedOn：子样式 H 覆盖父样式 Base 的字号。
        st = styles(
            para_style("Base", rpr_inner='<w:sz w:val="21"/>'),
            para_style("H", based_on="Base", rpr_inner='<w:sz w:val="32"/>'))
        resolver = StyleResolver(st)
        _, rpr = resolver.resolve_cascade("H", None, None, {})
        self.assertEqual(rpr["size_hp"], 32)


if __name__ == "__main__":
    unittest.main()
