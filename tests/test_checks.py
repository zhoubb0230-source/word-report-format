# -*- coding: utf-8 -*-
"""Unit tests for the deterministic judgment layer (scripts/lib/checks.py) and
heading inference (scripts/lib/headings.py).

Each test pins one decision recorded in references/已知陷阱与设计决策.md so a
future edit that "reverts to intuition" fails here instead of in production.
Pure functions, no docx needed.
"""
import os
import tempfile
import unittest

import helpers  # noqa: F401  (sets up sys.path for lib imports)
import checks
import headings
from docxcommon import StyleResolver, qn
from lxml import etree

SPEC = checks.load_spec(helpers.SPEC_PATH)


def eff(**kw):
    """Effective-format dict defaulting to a COMPLIANT body paragraph
    (仿宋 / 三号 / 固定行距28磅 / 首行缩进2字符); override what a test needs."""
    d = dict(east_asia="仿宋", ascii=None, size_hp=32, line=560, line_rule="exact",
             first_line_chars=200, first_line=None, left_chars=None, left=None,
             start_chars=None, start=None, right_chars=None, right=None,
             end_chars=None, end=None, hanging_chars=None, hanging=None,
             jc=None, outline=None)
    d.update(kw)
    return d


def rec(i=0, region="body", text="", eff_=None, **kw):
    r = dict(i=i, region=region, text=text, is_blank=False, is_toc=False,
             toc_level=None, auto_num=False, is_title=False, above_title=False,
             cover_role=None, is_heading=False, level=None, level_source=None,
             num_raw=None, caption=None, in_table=False, eff=eff_ or eff())
    r.update(kw)
    return r


class TestCoverTitle(unittest.TestCase):
    """陷阱 #2: 题目删首尾空格后再居中；左右缩进一起清零。"""

    def test_strip_both_center_and_clear_both_indents(self):
        r = rec(region="cover", is_title=True, text="  先进项目报告  ",
                eff_=eff(east_asia="宋体", size_hp=32, jc="left", left_chars=200,
                         right_chars=100))
        fix = checks.check_paragraph(r, SPEC)
        self.assertEqual(fix["strip_text"], "both")
        self.assertEqual(fix["set_jc"], "center")
        self.assertTrue(fix["clear_left_indent"])
        self.assertTrue(fix["clear_right_indent"], "右缩进必须与左缩进一起清零")
        self.assertEqual(fix["set_east_asia"], "方正小标宋")


class TestCoverField(unittest.TestCase):
    """陷阱 #1: 封面字段行首空格删除并左对齐（行内填空空格保留）。"""

    def test_leading_strip_and_left_align(self):
        r = rec(region="cover", text="  项目名称：先进项目",
                eff_=eff(east_asia="宋体", size_hp=32, jc="center"))
        fix = checks.check_paragraph(r, SPEC)
        # leading-only strip protects internal fill-in spaces
        self.assertEqual(fix["strip_text"], "leading")
        self.assertEqual(fix["set_jc"], "left")
        self.assertEqual(fix["set_east_asia"], "方正黑体")


class TestTocExcluded(unittest.TestCase):
    """陷阱 #4: 目录段落走目录字体校验，不落入标题分支，也不计连续性。"""

    def test_toc_uses_toc_font_not_heading(self):
        r = rec(region="toc", is_toc=True, toc_level=1, text="绪论\t1",
                eff_=eff(east_asia="黑体", size_hp=32))
        fix = checks.check_paragraph(r, SPEC)
        self.assertEqual(fix["set_east_asia"], "仿宋")  # toc font, not 黑体
        self.assertEqual(fix["set_size_hp"], 28)         # 14磅

    def test_toc_not_counted_in_continuity(self):
        recs = [rec(i=0, region="toc", is_toc=True, is_heading=False,
                    text="一、绪论")]
        self.assertEqual(checks.continuity(recs, SPEC), [])


class TestPatternHeadingNeverAutoEdited(unittest.TestCase):
    """陷阱 #5: 仅形状(pattern)标题只批注，绝不自动改文字。"""

    def test_pattern_heading_gets_hint_not_renumber(self):
        recs = [rec(i=0, is_heading=True, level=3, level_source="pattern",
                    num_raw="3.", text="3. 概述")]
        fixes = checks.continuity(recs, SPEC)
        self.assertEqual(len(fixes), 1)
        self.assertEqual(fixes[0]["op"], "hint")
        self.assertEqual(fixes[0]["rule_id"], "heading.unconfirmed")

    def test_confirmed_heading_gets_heading_font(self):
        r = rec(is_heading=True, level=1, level_source="outline",
                num_raw="一、", text="一、绪论", eff_=eff(east_asia="宋体"))
        fix = checks.check_paragraph(r, SPEC)
        self.assertEqual(fix["set_east_asia"], "黑体")  # level-1 heading font


class TestUnnumberedSection(unittest.TestCase):
    """惯例不编号章节（参考文献等）不占用同级序号计数。"""

    def test_unnumbered_does_not_consume_ordinal(self):
        recs = [
            rec(i=0, is_heading=True, level=1, level_source="outline",
                num_raw="一、", text="一、项目概况"),
            rec(i=1, is_heading=True, level=1, level_source="outline",
                num_raw=None, text="参考文献"),
            rec(i=2, is_heading=True, level=1, level_source="outline",
                num_raw=None, text="结果分析"),
        ]
        fixes = checks.continuity(recs, SPEC)
        by_idx = {f["para_index"]: f for f in fixes}
        self.assertNotIn(0, by_idx)          # 一、绪论 already correct
        self.assertNotIn(1, by_idx)          # 参考文献 untouched (not numbered)
        self.assertEqual(by_idx[2]["op"], "renumber_heading")
        self.assertEqual(by_idx[2]["new_token"], "二、")  # not 三、


class TestCaptionGrouping(unittest.TestCase):
    """图/表编号：平铺 vs 章-序分组，各自计数。"""

    def _cap(self, i, num_raw):
        return rec(i=i, caption={"kind": "figure", "num_raw": num_raw,
                                 "has_content": True, "source": "style"})

    def test_flat_renumber(self):
        recs = [self._cap(0, "1"), self._cap(1, "3")]
        fixes = checks.continuity(recs, SPEC)
        by_idx = {f["para_index"]: f for f in fixes}
        self.assertNotIn(0, by_idx)
        self.assertEqual(by_idx[1]["new_num"], "2")

    def test_chapter_based_renumber(self):
        recs = [self._cap(0, "1-1"), self._cap(1, "1-3")]
        fixes = checks.continuity(recs, SPEC)
        by_idx = {f["para_index"]: f for f in fixes}
        self.assertNotIn(0, by_idx)
        self.assertEqual(by_idx[1]["new_num"], "1-2")


class TestHeadingInference(unittest.TestCase):
    """层级判定优先级：样式名写明的级别 > 序号形状。"""

    def test_style_name_level_beats_shape(self):
        w_ns = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
        styles = etree.fromstring(
            ('<w:styles xmlns:w="%s"><w:style w:type="paragraph" '
             'w:styleId="H1"><w:name w:val="标题 1"/></w:style></w:styles>'
             % w_ns).encode("utf-8"))
        resolver = StyleResolver(styles)
        level, source = headings.infer_heading_level("H1", None, "1. 项目进展", resolver)
        self.assertEqual(level, 1)      # 标题 1 wins over "1." looking like level 3
        self.assertEqual(source, "style")

    def test_bare_pattern_is_unconfirmed(self):
        level, source = headings.infer_heading_level(None, None, "1. 概述", None)
        self.assertEqual(level, 3)
        self.assertEqual(source, "pattern")


class TestPageSetup(unittest.TestCase):
    def test_wrong_margins_flagged(self):
        fix = checks.check_page_setup(
            {"top": 1000, "bottom": 1814, "left": 1616, "right": 1616,
             "header": 850, "footer": 992}, SPEC)
        self.assertIsNotNone(fix)
        self.assertEqual(fix["set_pgmar"]["top"], 1984)

    def test_compliant_margins_no_fix(self):
        fix = checks.check_page_setup(
            {"top": 1984, "bottom": 1814, "left": 1616, "right": 1616,
             "header": 850, "footer": 992}, SPEC)
        self.assertIsNone(fix)


class TestWesternFont(unittest.TestCase):
    """P0-1: 西文字体 Times New Roman 在正文/标题/表格上被规范化。"""

    def test_body_wrong_western_flagged(self):
        fix = checks.check_paragraph(rec(eff_=eff(ascii="Calibri")), SPEC)
        self.assertEqual(fix["set_ascii"], "Times New Roman")

    def test_body_missing_western_flagged(self):
        # default eff has ascii=None (unset western) -> normalized to TNR
        fix = checks.check_paragraph(rec(), SPEC)
        self.assertEqual(fix["set_ascii"], "Times New Roman")

    def test_fully_compliant_body_no_fix(self):
        fix = checks.check_paragraph(rec(eff_=eff(ascii="Times New Roman")), SPEC)
        self.assertIsNone(fix)

    def test_heading_gets_western(self):
        r = rec(is_heading=True, level=1, level_source="outline",
                num_raw="一、", text="一、绪论", eff_=eff(ascii="Arial"))
        fix = checks.check_paragraph(r, SPEC)
        self.assertEqual(fix["set_ascii"], "Times New Roman")

    def test_cover_field_not_forced_western(self):
        # cover layout lines deliberately opt out of western enforcement:
        # an otherwise-compliant field with a wrong Latin font gets NO fix.
        r = rec(region="cover", text="项目名称：先进",
                eff_=eff(east_asia="方正黑体", size_hp=30, ascii="Calibri",
                         first_line_chars=None))
        fix = checks.check_paragraph(r, SPEC)
        self.assertIsNone(fix)  # font/size ok, western not enforced here


class TestDocHints(unittest.TestCase):
    """P0-1: 目录超三级 / 页码字体 提示项。"""

    def test_toc_depth_hint(self):
        structure = {"records": [
            rec(i=0, region="toc", is_toc=True, toc_level=1, text="a"),
            rec(i=1, region="toc", is_toc=True, toc_level=4, text="b"),
        ], "page_number": None}
        ids = {h["rule_id"] for h in checks.doc_hints(structure, SPEC)}
        self.assertIn("toc.depth", ids)

    def test_toc_within_depth_no_hint(self):
        structure = {"records": [
            rec(i=0, region="toc", is_toc=True, toc_level=1, text="a"),
            rec(i=1, region="toc", is_toc=True, toc_level=3, text="b"),
        ], "page_number": None}
        ids = {h["rule_id"] for h in checks.doc_hints(structure, SPEC)}
        self.assertNotIn("toc.depth", ids)

    def test_page_number_font_hint(self):
        structure = {"records": [rec(i=0, region="cover", text="x")],
                     "page_number": {"ascii": "宋体", "size_hp": 21}}
        ids = {h["rule_id"] for h in checks.doc_hints(structure, SPEC)}
        self.assertIn("page_number.font", ids)

    def test_page_number_ok_no_hint(self):
        structure = {"records": [rec(i=0, region="cover", text="x")],
                     "page_number": {"ascii": "Times New Roman", "size_hp": 24}}
        ids = {h["rule_id"] for h in checks.doc_hints(structure, SPEC)}
        self.assertNotIn("page_number.font", ids)

    def test_page_number_unread_no_false_positive(self):
        structure = {"records": [rec(i=0, region="cover", text="x")],
                     "page_number": {"ascii": None, "size_hp": None}}
        ids = {h["rule_id"] for h in checks.doc_hints(structure, SPEC)}
        self.assertNotIn("page_number.font", ids)


class TestScanPageNumber(unittest.TestCase):
    """P0-1: 20 阶段从页眉/页脚扫描页码字体。"""

    def test_reads_footer_page_field(self):
        mod = helpers.load_script("20_extract_structure.py")
        d = tempfile.mkdtemp()
        try:
            os.makedirs(os.path.join(d, "word"))
            footer = (
                '<?xml version="1.0" encoding="UTF-8"?>'
                '<w:ftr xmlns:w="%s"><w:p>'
                '<w:r><w:rPr><w:rFonts w:ascii="宋体"/><w:sz w:val="21"/></w:rPr>'
                '<w:t>第</w:t></w:r>'
                '<w:fldSimple w:instr=" PAGE "><w:r><w:t>1</w:t></w:r></w:fldSimple>'
                '</w:p></w:ftr>' % helpers.W_NS)
            with open(os.path.join(d, "word", "footer1.xml"), "w", encoding="utf-8") as f:
                f.write(footer)
            info = mod.scan_page_number(d)
            self.assertEqual(info, {"ascii": "宋体", "size_hp": 21})
        finally:
            import shutil
            shutil.rmtree(d, ignore_errors=True)

    def test_no_page_field_returns_none(self):
        mod = helpers.load_script("20_extract_structure.py")
        d = tempfile.mkdtemp()
        try:
            os.makedirs(os.path.join(d, "word"))
            with open(os.path.join(d, "word", "footer1.xml"), "w", encoding="utf-8") as f:
                f.write('<?xml version="1.0"?><w:ftr xmlns:w="%s"><w:p><w:r>'
                        '<w:t>页脚</w:t></w:r></w:p></w:ftr>' % helpers.W_NS)
            self.assertIsNone(mod.scan_page_number(d))
        finally:
            import shutil
            shutil.rmtree(d, ignore_errors=True)


class TestStripParaWsInternalPreserved(unittest.TestCase):
    """apply 层 _strip_para_ws：只清首尾，保留行内填空空格。"""

    def test_both_preserves_internal(self):
        mod = helpers.load_script("40_apply_fixes.py")
        p = etree.Element(qn("w:p"))
        r = etree.SubElement(p, qn("w:r"))
        t = etree.SubElement(r, qn("w:t"))
        t.text = "  20   年   月  "
        mod._strip_para_ws(p, "both")
        self.assertEqual(t.text, "20   年   月")


if __name__ == "__main__":
    unittest.main()
