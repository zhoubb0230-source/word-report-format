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
    d = dict(east_asia="仿宋", ascii=None, size_hp=32, bold=True,
             line=560, line_rule="exact",
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
        self.assertEqual(fix["set_east_asia"], "方正小标宋_GBK")


class TestCoverField(unittest.TestCase):
    """陷阱 #1: 封面字段——两端对齐、首行缩进2字符、删行首空格、清左侧缩进
    （行内填空空格保留）。"""

    def test_justify_first_line_indent_and_leading_strip(self):
        r = rec(region="cover", text="  项目名称：先进项目",
                eff_=eff(east_asia="宋体", size_hp=32, jc="center",
                         first_line_chars=None, left_chars=200))
        fix = checks.check_paragraph(r, SPEC)
        # leading-only strip protects internal fill-in spaces
        self.assertEqual(fix["strip_text"], "leading")
        self.assertEqual(fix["set_jc"], "both")            # 两端对齐
        self.assertEqual(fix["set_first_line_chars"], 200)  # 首行缩进2字符
        self.assertTrue(fix["clear_left_indent"])           # 清左侧缩进
        self.assertEqual(fix["set_east_asia"], "方正黑体_GBK")

    def test_field_font_size_and_line_spacing(self):
        # 报告题目下的要素：方正黑体_GBK / 小三（30半点）/ 2倍行距（auto/480）。
        r = rec(region="cover", text="项目名称：先进项目",
                eff_=eff(east_asia="宋体", size_hp=32, jc="both",
                         first_line_chars=200, line=560, line_rule="exact"))
        fix = checks.check_paragraph(r, SPEC)
        self.assertEqual(fix["set_east_asia"], "方正黑体_GBK")
        self.assertEqual(fix["set_size_hp"], 30)            # 小三
        self.assertEqual(fix["set_line_exact"], 480)        # 2倍行距
        self.assertEqual(fix["set_line_rule"], "auto")

    def test_project_number_is_field_not_classification(self):
        # 问题1：项目编号是报告题目下的要素（方正黑体_GBK 小三），不是密级/文本
        # 编号那类 classification（仿宋 16磅）。位于标题下方、无 above_title。
        r = rec(region="cover", text="项目编号：2025010203",
                eff_=eff(east_asia="仿宋", size_hp=32, jc="left",
                         first_line_chars=None))
        self.assertEqual(checks.cover_role(r, SPEC), "field")
        fix = checks.check_paragraph(r, SPEC)
        self.assertEqual(fix["set_east_asia"], "方正黑体_GBK")  # 非仿宋
        self.assertEqual(fix["set_size_hp"], 30)               # 小三，非三号

    def test_compliant_field_no_fix(self):
        r = rec(region="cover", text="项目名称：先进项目",
                eff_=eff(east_asia="方正黑体_GBK", size_hp=30, jc="both",
                         first_line_chars=200, line=480, line_rule="auto"))
        self.assertIsNone(checks.check_paragraph(r, SPEC))


class TestTableBody(unittest.TestCase):
    """表格内容：仿宋 四号（28半点）、行距固定值28磅、无任何缩进。"""

    def test_table_size_is_sihao(self):
        r = rec(in_table=True, text="设备名称123", has_western=True,
                eff_=eff(east_asia="仿宋", size_hp=32, ascii="仿宋",
                         first_line_chars=None))
        fix = checks.check_paragraph(r, SPEC)
        self.assertEqual(fix["set_size_hp"], 28)   # 四号，非三号

    def test_table_sihao_compliant_no_fix(self):
        # 合规表格内容：四号 + 固定28磅 + 无缩进
        r = rec(in_table=True, text="设备名称", has_western=False,
                eff_=eff(east_asia="仿宋", size_hp=28, ascii="仿宋",
                         first_line_chars=None, line=560, line_rule="exact"))
        self.assertIsNone(checks.check_paragraph(r, SPEC))

    def test_table_indent_and_line_spacing_flagged(self):
        # 表格内容带缩进 + 行距非28磅固定 -> 都要被清/改
        r = rec(in_table=True, text="设备名称", has_western=False,
                eff_=eff(east_asia="仿宋", size_hp=28, ascii="仿宋",
                         first_line_chars=200, left_chars=100,
                         line=480, line_rule="auto"))
        fix = checks.check_paragraph(r, SPEC)
        self.assertEqual(fix["set_line_exact"], 560)
        self.assertEqual(fix["set_line_rule"], "exact")
        self.assertTrue(fix["clear_left_indent"])
        self.assertEqual(fix["set_first_line_chars"], 0)


class TestCoverTextNumberVsProjectNumber(unittest.TestCase):
    """文本编号（题目上，仿宋三号）与 项目编号（题目下，field）是两个不同的必备项，
    不得互相顶替。"""

    def test_project_number_does_not_satisfy_text_number(self):
        # 只有项目编号、没有文本编号 -> 仍应报缺文本编号（两者不是别名）。
        structure = {"records": [
            rec(i=0, region="cover", text="项目编号：2025010203"),
        ]}
        missing = checks._cover_missing_fields(structure, SPEC)
        self.assertIn("文本编号", missing)

    def test_both_are_required_fields(self):
        required = SPEC["cover"]["required_fields"]
        self.assertIn("文本编号", required)
        self.assertIn("项目编号", required)

    def test_no_field_aliases_conflation(self):
        # 明确不再有把项目编号并到文本编号的别名机制。
        self.assertNotIn("field_aliases", SPEC["cover"])


class TestTocExcluded(unittest.TestCase):
    """陷阱 #4: 目录段落走目录字体校验，不落入标题分支，也不计连续性。"""

    def test_toc_uses_toc_font_not_heading(self):
        r = rec(region="toc", is_toc=True, toc_level=1, text="绪论\t1",
                eff_=eff(east_asia="黑体", size_hp=28))
        fix = checks.check_paragraph(r, SPEC)
        self.assertEqual(fix["set_east_asia"], "仿宋")  # toc font, not 黑体
        self.assertEqual(fix["set_size_hp"], 30)         # 小三

    def test_toc_not_counted_in_continuity(self):
        recs = [rec(i=0, region="toc", is_toc=True, is_heading=False,
                    text="一、绪论")]
        self.assertEqual(checks.continuity(recs, SPEC), [])

    def test_toc_fix_carries_no_comment(self):
        # 目录条目的格式 fix 不挂批注：updateFields 刷新目录会孤儿化批注区间，
        # 显示成空白批注；目录格式靠样式回写保证。
        r = rec(region="toc", is_toc=True, toc_level=2, text="研究方法\t3",
                eff_=eff(east_asia="黑体", size_hp=28, left_chars=0))
        fix = checks.check_paragraph(r, SPEC)
        self.assertIsNotNone(fix)               # 仍产出直接格式修正
        self.assertFalse(fix["comment"])        # 但不挂批注


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
    """图/表编号：平铺走 Word 自动编号；章-序分组仍走静态重编号。"""

    def _cap(self, i, num_raw, **kw):
        return rec(i=i, caption={"kind": "figure", "num_raw": num_raw,
                                 "has_content": True, "source": "style"}, **kw)

    def test_flat_captions_go_to_word_autonumbering(self):
        # 静态编号插/删图后会整体失序，用户阶段2 验收否掉了它：平铺编号一律改成
        # Word 自动编号（注入 图%1 定义、编号后带制表符），文字里的"图N"删掉。
        recs = [self._cap(0, "1"), self._cap(1, "3")]
        fixes = checks.continuity(recs, SPEC)
        by_idx = {f["para_index"]: f for f in fixes}
        self.assertEqual({f["op"] for f in fixes}, {"autonumber_caption"})
        self.assertEqual(by_idx[0]["kind"], "figure")
        # 连"编号看着是对的"那条也要转（静态编号本身就不合规）
        self.assertIn(0, by_idx)

    def test_already_canonical_autonumbered_caption_is_left_alone(self):
        # 幂等：上一轮处理过的产物（自动编号 + 文字里无编号 + 注入的图标题样式）
        # 不再产生 fix，重跑收敛。
        r = self._cap(0, None, auto_num=True, style_id="FGWCanonCaptionFig")
        self.assertEqual(checks.continuity([r], SPEC), [])

    def test_foreign_autonumbering_still_converted(self):
        # 文档自带的一套外来编号定义（未承载在我们注入的题注样式上）仍要转成规范编号，
        # 否则"Figure 1"这类会被当成合规放过。
        r = self._cap(0, None, auto_num=True, style_id="SomeOtherCaption")
        fixes = checks.continuity([r], SPEC)
        self.assertEqual([f["op"] for f in fixes], ["autonumber_caption"])

    def test_chapter_based_renumber(self):
        # 章-序分组（图1-1）平铺自动编号表达不了，仍走静态重编号——别顺手也改掉。
        recs = [self._cap(0, "1-1"), self._cap(1, "1-3")]
        fixes = checks.continuity(recs, SPEC)
        by_idx = {f["para_index"]: f for f in fixes}
        self.assertNotIn(0, by_idx)
        self.assertEqual(by_idx[1]["op"], "renumber_caption")
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


class TestLibreOfficeTocStyle(unittest.TestCase):
    """LibreOffice .doc→.docx 兼容：目录条目样式被命名为 'Contents N'（Word 用
    'TOC N'/'目录 N'）。若不识别，toc_level 取不到而退化成不缩进（目录被压平）、
    且 40 的样式回写会跳过它们，刷新后目录彻底丢缩进。"""

    def _resolver(self):
        w_ns = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
        styles = etree.fromstring(
            ('<w:styles xmlns:w="%s">'
             '<w:style w:type="paragraph" w:styleId="Contents2">'
             '<w:name w:val="Contents 2"/></w:style></w:styles>'
             % w_ns).encode("utf-8"))
        return StyleResolver(styles)

    def test_contents_style_recognized_as_toc(self):
        resolver = self._resolver()
        self.assertTrue(headings.style_is_toc("Contents2", resolver))
        self.assertEqual(headings.toc_level_from_style("Contents2", resolver), 2)


class TestTocStyleLeftCharsOnly(unittest.TestCase):
    """目录样式回写：_patch_toc_styles 钉 font/size/leftChars(按级 0/200/400) **以及
    制表位**（左制表位 + 右点线制表位）。

    别再删制表位：2026-07 曾以"Word updateFields 会自建、写死会被盖掉"为由删掉，
    阶段0 的 Word 逐项验收**推翻**了它——目录条目的"编号→标题→点线→页码"排布正是
    靠样式制表位撑起来的，缺了点线会跑到编号与标题之间、页码换行（陷阱 #12）。
    制表位没有字符单位形式，只能按【文档字符单位字号】把字符数换算成绝对 twips：
    Normal=五号(21) → 二级左 1050、右点线 8665。判定层仍只按 leftChars 判合规。"""

    W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

    def _make_pkg(self, tmp):
        word = os.path.join(tmp, "word")
        os.makedirs(word)
        with open(os.path.join(word, "styles.xml"), "wb") as f:
            f.write(('<w:styles xmlns:w="%s">'
                     '<w:style w:type="paragraph" w:styleId="TOC2">'
                     '<w:name w:val="toc 2"/></w:style></w:styles>' % self.W
                     ).encode("utf-8"))
        with open(os.path.join(word, "document.xml"), "wb") as f:
            f.write(('<w:document xmlns:w="%s"><w:body><w:sectPr>'
                     '<w:pgSz w:w="11906" w:h="16838"/>'
                     '<w:pgMar w:left="1616" w:right="1616" w:top="1984" w:bottom="1814"/>'
                     '</w:sectPr></w:body></w:document>' % self.W).encode("utf-8"))
        return tmp

    def test_toc_style_gets_leftchars_and_tabs(self):
        mod = helpers.load_script("40_apply_fixes.py")
        with tempfile.TemporaryDirectory() as tmp:
            self._make_pkg(tmp)
            n = mod._patch_toc_styles(tmp, SPEC["toc"], char_unit_hp=21)
            self.assertEqual(n, 1)
            root = etree.parse(os.path.join(tmp, "word", "styles.xml")).getroot()
            ind = root.find(".//{%s}ind" % self.W)
            self.assertIsNotNone(ind)
            # 二级 leftChars = 200（2字符）
            self.assertEqual(ind.get("{%s}leftChars" % self.W), "200")
            # 制表位：左 5字符=1050、右点线 41.26字符=8665（均按五号 210/字符换算）
            tabs = root.findall(".//{%s}tabs/{%s}tab" % (self.W, self.W))
            got = [(t.get("{%s}val" % self.W), t.get("{%s}pos" % self.W),
                    t.get("{%s}leader" % self.W)) for t in tabs]
            self.assertEqual(got, [("left", "1050", None), ("right", "8665", "dot")])

    def test_toc_tab_positions_follow_char_unit_size(self):
        # 字符单位随 Normal 字号走：Normal=三号(32) 时同样的字符数换算成 1600/13203。
        mod = helpers.load_script("40_apply_fixes.py")
        with tempfile.TemporaryDirectory() as tmp:
            self._make_pkg(tmp)
            mod._patch_toc_styles(tmp, SPEC["toc"], char_unit_hp=32)
            root = etree.parse(os.path.join(tmp, "word", "styles.xml")).getroot()
            pos = [t.get("{%s}pos" % self.W)
                   for t in root.findall(".//{%s}tabs/{%s}tab" % (self.W, self.W))]
            self.assertEqual(pos, ["1600", "13203"])


class TestInt2Cn(unittest.TestCase):
    def test_basic_range(self):
        self.assertEqual(checks.int2cn(1), "一")
        self.assertEqual(checks.int2cn(10), "十")
        self.assertEqual(checks.int2cn(23), "二十三")
        self.assertEqual(checks.int2cn(99), "九十九")

    def test_out_of_range_falls_back_no_crash(self):
        self.assertEqual(checks.int2cn(100), "100")  # was IndexError before guard
        self.assertEqual(checks.int2cn(0), "0")


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
    """P0-1 + Bug2: 西文字体只在【含西文】的正文/标题/表格上规范化为 TNR；
    纯中文段落绝不因西文字体被批注。"""

    def test_body_with_western_wrong_font_flagged(self):
        r = rec(text="正文ABC", has_western=True, eff_=eff(ascii="Calibri"))
        self.assertEqual(checks.check_paragraph(r, SPEC)["set_ascii"], "Times New Roman")

    def test_body_with_western_unset_font_flagged(self):
        # ascii unset but the line has western chars -> normalized to TNR
        r = rec(text="正文123", has_western=True)
        self.assertEqual(checks.check_paragraph(r, SPEC)["set_ascii"], "Times New Roman")

    def test_pure_chinese_body_no_western_fix(self):
        # Bug2: pure-Chinese line whose Latin font is 仿宋 must NOT be flagged.
        r = rec(text="纯中文正文。", has_western=False, eff_=eff(ascii="仿宋"))
        self.assertIsNone(checks.check_paragraph(r, SPEC))

    def test_pure_chinese_table_no_western_fix(self):
        # Bug2 exact report: 表格内容纯中文 误判 "西文应为 TNR（实际：仿宋）".
        r = rec(in_table=True, text="设备名称", has_western=False,
                eff_=eff(east_asia="仿宋", size_hp=28, ascii="仿宋",
                         first_line_chars=None, line=560, line_rule="exact"))
        self.assertIsNone(checks.check_paragraph(r, SPEC))

    def test_heading_with_western_gets_tnr(self):
        r = rec(is_heading=True, level=1, level_source="outline",
                num_raw="一、", text="一、Overview A", has_western=True,
                eff_=eff(ascii="Arial"))
        self.assertEqual(checks.check_paragraph(r, SPEC)["set_ascii"], "Times New Roman")

    def test_cover_field_western_enforced_when_line_has_digits(self):
        # 阶段0 Word 验收推翻了陷阱#7 在封面的范围：封面要素/密级行里的数字与西文
        # 也要走 Times（spec 的 cover_field.enforce_western）。含西文的行才校验。
        r = rec(region="cover", text="项目编号：KJ-2024-001", has_western=True,
                eff_=eff(east_asia="方正黑体_GBK", size_hp=30, ascii="Calibri",
                         jc="both", first_line_chars=200, line=480,
                         line_rule="auto"))
        self.assertEqual(checks.check_paragraph(r, SPEC)["set_ascii"], "Times New Roman")

    def test_pure_chinese_cover_field_no_western_fix(self):
        # 但纯中文的封面行没有西文可规范，仍然一个字都不动（has_western 保护）。
        r = rec(region="cover", text="承担单位（盖章）：某某研究院", has_western=False,
                eff_=eff(east_asia="方正黑体_GBK", size_hp=30, ascii="仿宋",
                         jc="both", first_line_chars=200, line=480,
                         line_rule="auto"))
        self.assertIsNone(checks.check_paragraph(r, SPEC))

    def test_toc_still_not_forced_western(self):
        # 目录仍**不**套西文（_check_toc 不传 western）——只有封面被推翻，别顺手扩大。
        r = rec(region="toc", is_toc=True, toc_level=1, text="一、概述 3",
                has_western=True,
                eff_=eff(east_asia="仿宋", size_hp=30, ascii="Calibri"))
        fix = checks.check_paragraph(r, SPEC)
        self.assertIsNone(fix if fix is None else fix.get("set_ascii"))


class TestNewFormatRules2026(unittest.TestCase):
    """2026-07 批次新增/调整的规则（见 HANDOFF §12 分诊 A/B）。每条锁一个决策。"""

    def test_classification_font_is_fangzheng_heiti_gbk(self):
        # #1 密级/文本编号 -> 方正黑体_GBK（三号不变）
        r = rec(region="cover", above_title=True, text="机密　2025013",
                eff_=eff(east_asia="仿宋", size_hp=32))
        self.assertEqual(checks.cover_role(r, SPEC), "classification")
        fix = checks.check_paragraph(r, SPEC)
        self.assertEqual(fix["set_east_asia"], "方正黑体_GBK")
        self.assertIsNone(fix["set_size_hp"])  # 三号(32) 不变

    def test_title_font_gbk_and_digits_western(self):
        # #2 题目字体 方正小标宋_GBK；题目内数字用西文(TNR)
        r = rec(region="cover", is_title=True, text="2024年度自评价报告",
                has_western=True,
                eff_=eff(east_asia="宋体", size_hp=40, jc="center", ascii="宋体",
                         first_line_chars=None))
        fix = checks.check_paragraph(r, SPEC)
        self.assertEqual(fix["set_east_asia"], "方正小标宋_GBK")
        self.assertEqual(fix["set_ascii"], "Times New Roman")

    def test_title_pure_chinese_no_western_flag(self):
        # 纯中文题目（无数字）不因西文被误报
        r = rec(region="cover", is_title=True, text="先进项目自评价报告",
                has_western=False,
                eff_=eff(east_asia="方正小标宋_GBK", size_hp=40, jc="center",
                         first_line_chars=None))
        self.assertIsNone(checks.check_paragraph(r, SPEC))

    def test_heading_line_spacing_28pt(self):
        # #6 一~四级标题行距固定值28磅(560/exact)
        r = rec(is_heading=True, level=2, level_source="outline", num_raw="（一）",
                text="（一）研究背景",
                eff_=eff(east_asia="楷体", size_hp=32, line=None, line_rule=None))
        fix = checks.check_paragraph(r, SPEC)
        self.assertEqual(fix["set_line_exact"], 560)
        self.assertEqual(fix["set_line_rule"], "exact")

    def test_caption_font_size_line(self):
        # #8 图/表标题 仿宋 三号 + 行距固定28磅 + 居中 + 无缩进
        r = rec(text="图1 系统结构", caption={"kind": "figure", "num_raw": "1",
                "has_content": True, "source": "style"},
                eff_=eff(east_asia="黑体", size_hp=40, jc="left",
                         first_line_chars=200, line=None, line_rule=None))
        fix = checks.check_paragraph(r, SPEC)
        self.assertEqual(fix["set_east_asia"], "仿宋")
        self.assertEqual(fix["set_size_hp"], 32)
        self.assertEqual(fix["set_line_exact"], 560)
        self.assertEqual(fix["set_jc"], "center")
        self.assertEqual(fix["set_first_line_chars"], 0)

    def test_body_removes_space_before_after(self):
        # #15 正文去段前/段后间距
        r = rec(text="正文一段。",
                eff_=eff(space_before=100, space_after=50))
        fix = checks.check_paragraph(r, SPEC)
        self.assertTrue(fix["clear_space_before_after"])

    def test_body_compliant_with_zero_spacing_no_fix(self):
        r = rec(text="正文一段。",
                eff_=eff(space_before=None, space_after=None))
        self.assertIsNone(checks.check_paragraph(r, SPEC))

    def test_headings_must_be_bold(self):
        # 2026-07 新增规范：一~四级标题加粗。
        r = rec(is_heading=True, level=2, level_source="outline", num_raw="（一）",
                text="（一）研究方法",
                eff_=eff(east_asia="楷体", size_hp=32, line=560, line_rule="exact",
                         first_line_chars=200, bold=False))
        fix = checks.check_paragraph(r, SPEC)
        self.assertTrue(fix["set_bold"])
        self.assertIn("加粗", fix["rule_text"])

    def test_bold_heading_compliant_no_bold_fix(self):
        r = rec(is_heading=True, level=2, level_source="outline", num_raw="（一）",
                text="（一）研究方法",
                eff_=eff(east_asia="楷体", size_hp=32, line=560, line_rule="exact",
                         first_line_chars=200, bold=True))
        self.assertIsNone(checks.check_paragraph(r, SPEC))

    def test_half_bold_heading_is_flagged(self):
        # "前半加粗、后半不加粗"：eff.bold 走【全体一致】语义（20 抽取时只要有一个
        # 文字 run 不粗就是 False），所以这种半粗半细会被判不合规而不是被多数票放过。
        r = rec(is_heading=True, level=3, level_source="outline", num_raw="1.",
                text="1. 数据来源与口径说明",
                eff_=eff(east_asia="仿宋", size_hp=32, line=560, line_rule="exact",
                         first_line_chars=200, bold=False))
        self.assertTrue(checks.check_paragraph(r, SPEC)["set_bold"])

    def test_body_bold_is_not_touched(self):
        # 规范没规定正文加粗 → 正文样式不承载 bold，作者的行内加粗保持原样。
        r = rec(text="正文内容。",
                eff_=eff(east_asia="仿宋", size_hp=32, line=560, line_rule="exact",
                         first_line_chars=200, space_before=0, space_after=0,
                         bold=True))
        self.assertIsNone(checks.check_paragraph(r, SPEC))

    def test_direct_hanging_forces_first_line_rewrite(self):
        # #12/#14 段落带直接 hanging，即使首行缩进值看似已对，也强制重写首行缩进
        # （apply 写 firstLine 时清直接 hanging）
        r = rec(text="正文一段。",
                eff_=eff(first_line_chars=200, hanging=420))
        fix = checks.check_paragraph(r, SPEC)
        self.assertEqual(fix["set_first_line_chars"], 200)


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
