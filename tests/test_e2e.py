# -*- coding: utf-8 -*-
"""End-to-end smoke test: run the real pipeline scripts (05→10→20→30→40→45) on a
minimal generated .docx and assert the output validates and preserves paragraph
count. Exercises the shared helpers (tag_regions / in_textbox) through their
actual entry points.

Skipped automatically if lxml is not importable (the pipeline's one dependency).
"""
import json
import os
import shutil
import subprocess
import sys
import tempfile
import unittest

import helpers

try:
    import lxml  # noqa: F401
    HAVE_LXML = True
except ImportError:
    HAVE_LXML = False


def run(*args):
    """Run a pipeline script, assert exit 0, return parsed last-line JSON."""
    proc = subprocess.run([sys.executable] + list(args),
                          capture_output=True, text=True)
    assert proc.returncode == 0, "cmd %s failed (%d): %s" % (args, proc.returncode, proc.stderr)
    last = [ln for ln in proc.stdout.strip().splitlines() if ln.strip()][-1]
    return json.loads(last)


@unittest.skipUnless(HAVE_LXML, "lxml not installed")
class TestPipelineEndToEnd(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.mkdtemp(prefix="wrf_e2e_")

    def tearDown(self):
        shutil.rmtree(self.tmp, ignore_errors=True)

    def _script(self, name):
        return os.path.join(helpers.SCRIPTS, name)

    def test_full_pipeline(self):
        # cover title (large, centered), then an outline heading ends the cover,
        # then a non-compliant body paragraph. Wrong margins trigger a section fix.
        body = (
            helpers.para("先进项目2024年度自评价报告", east_asia="宋体",
                         size_hp=44, jc="center")
            + helpers.para("一、绪论", east_asia="宋体", size_hp=32, outline=0)
            + helpers.para("这是正文内容需要被规范化。", east_asia="宋体", size_hp=32)
        )
        src = os.path.join(self.tmp, "input.docx")
        helpers.build_docx(src, body, pg_mar=(1000, 1814, 1616, 1616, 850, 992))

        base = os.path.join(self.tmp, "work_base")
        wd = run(self._script("05_new_workdir.py"), base)["workdir"]
        run(self._script("10_prepare_input.py"), src, wd)
        extracted = run(self._script("20_extract_structure.py"), wd)
        self.assertEqual(extracted["n_paragraphs"], 3)
        # sharding is retired: extraction must not emit shard artifacts/fields
        self.assertNotIn("n_shards", extracted)
        self.assertFalse(os.path.isdir(os.path.join(wd, "shards")))

        checked = run(self._script("30_check_format.py"), wd)
        self.assertGreater(checked["n_fixes"], 0)

        applied = run(self._script("40_apply_fixes.py"), wd)
        self.assertEqual(applied["status"], "ok")

        validated = run(self._script("45_validate_output.py"), wd)
        self.assertTrue(validated["ok"], validated.get("errors"))
        self.assertEqual(validated["paragraphs_out"],
                         validated["paragraphs_reference"])

    def test_comments_merged_per_paragraph(self):
        # A heading with BOTH a wrong font (format fix) and a wrong ordinal
        # (renumber fix) must end up with ONE merged comment, not two.
        body = (helpers.para("三、项目概况", east_asia="宋体", size_hp=32, outline=0)
                + helpers.para("正文内容。", east_asia="宋体", size_hp=32))
        src = os.path.join(self.tmp, "in3.docx")
        helpers.build_docx(src, body)
        base = os.path.join(self.tmp, "wb3")
        wd = run(self._script("05_new_workdir.py"), base)["workdir"]
        run(self._script("10_prepare_input.py"), src, wd)
        run(self._script("20_extract_structure.py"), wd)
        run(self._script("30_check_format.py"), wd)
        run(self._script("40_apply_fixes.py"), wd)

        import collections
        from lxml import etree
        with open(os.path.join(wd, "fixes.json"), encoding="utf-8") as f:
            fixes = json.load(f)
        counts = collections.Counter(
            fx["para_index"] for fx in fixes
            if fx.get("comment") and fx.get("rule_text") and fx.get("para_index") is not None)
        # the heading paragraph really does carry >=2 commentable fixes
        self.assertGreaterEqual(counts[0], 2)

        w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
        root = etree.parse(os.path.join(wd, "out_pkg", "word", "comments.xml")).getroot()
        n_comments = len(root.findall("{%s}comment" % w))
        # one comment per distinct commented paragraph (merge worked)
        self.assertEqual(n_comments, len(counts))

    def test_toc_field_span_tags_continuation_entries(self):
        # Bug1: a TOC field spans many paragraphs but only the first carries the
        # TOC instruction; continuation entries with an inherited outlineLvl were
        # mis-detected as 黑体 headings. The field-span tagging must catch them.
        def entry(text, first=False, last=False):
            ppr = '<w:pPr><w:outlineLvl w:val="0"/></w:pPr>'
            r = ""
            if first:
                r += ('<w:r><w:fldChar w:fldCharType="begin"/></w:r>'
                      '<w:r><w:instrText xml:space="preserve"> TOC \\o "1-3" \\h </w:instrText></w:r>'
                      '<w:r><w:fldChar w:fldCharType="separate"/></w:r>')
            r += ('<w:r><w:rPr><w:rFonts w:eastAsia="黑体"/></w:rPr>'
                  '<w:t xml:space="preserve">%s</w:t></w:r>' % text)
            if last:
                r += '<w:r><w:fldChar w:fldCharType="end"/></w:r>'
            return "<w:p>%s%s</w:p>" % (ppr, r)

        body = (helpers.para("封面标题", east_asia="宋体", size_hp=44)
                + helpers.para("目录", east_asia="黑体")
                + entry("第一章 概述\t1", first=True)
                + entry("第二章 方法\t2")
                + entry("第三章 结果\t3", last=True)
                + helpers.para("一、概述", east_asia="宋体", outline=0)
                + helpers.para("正文。", east_asia="宋体"))
        src = os.path.join(self.tmp, "toc.docx")
        helpers.build_docx(src, body)
        base = os.path.join(self.tmp, "wbtoc")
        wd = run(self._script("05_new_workdir.py"), base)["workdir"]
        run(self._script("10_prepare_input.py"), src, wd)
        run(self._script("20_extract_structure.py"), wd)
        with open(os.path.join(wd, "structure.json"), encoding="utf-8") as f:
            recs = json.load(f)["records"]
        toc_texts = [r["text"] for r in recs if r["region"] == "toc"]
        for t in ("第一章", "第二章", "第三章"):
            self.assertTrue(any(t in x for x in toc_texts), "%s 未标记为目录" % t)
        heads = [r["text"] for r in recs if r["is_heading"]]
        self.assertFalse(any("章" in h for h in heads),
                         "目录条目被误判为标题：%s" % heads)

    def test_apply_review_recomputes_regions(self):
        # A tiny doc still round-trips through 27 with an empty overrides file
        # (a no-op review): it must rewrite structure.json with region tags on
        # every record and NOT create any shard artifacts.
        body = (helpers.para("一、绪论", east_asia="宋体", outline=0)
                + helpers.para("正文。", east_asia="宋体"))
        src = os.path.join(self.tmp, "in2.docx")
        helpers.build_docx(src, body)
        base = os.path.join(self.tmp, "wb2")
        wd = run(self._script("05_new_workdir.py"), base)["workdir"]
        run(self._script("10_prepare_input.py"), src, wd)
        run(self._script("20_extract_structure.py"), wd)
        ov = os.path.join(self.tmp, "ov.json")
        with open(ov, "w") as f:
            f.write("{}")
        res = run(self._script("27_apply_review.py"), wd, ov)
        self.assertEqual(res["status"], "ok")
        self.assertFalse(os.path.isdir(os.path.join(wd, "shards")))
        with open(os.path.join(wd, "structure.json"), encoding="utf-8") as f:
            recs = json.load(f)["records"]
        self.assertTrue(all(r.get("region") in ("cover", "toc", "body") for r in recs))


@unittest.skipUnless(HAVE_LXML, "lxml not installed")
class TestApplyAbsoluteIndentCompanion(unittest.TestCase):
    """严格-spec §2.2：直接层缩进**只写字符单位**，不补非零的绝对伴随值。

    历史（别照直觉加回来）：40 曾给每个字符单位缩进补一个绝对 firstLine/left，用来
    压过 LibreOffice 转换出的、继承自样式/编号层的绝对 hanging（陷阱 #10）。方案C
    阶段2 之后两条继承路径都被根治——编号层由甲法克隆钳中和、样式层由 canonical 命名
    样式注入+指派接管——伴随值失去理由且**自身违规**（Word 会把缩进显示成厘米而不是
    "2 字符"），已删除。要清零的方向仍写显式 0（零值无单位歧义，用来挡继承）。

    `_char_twips` / `_default_char_unit_hp` 仍在：目录样式的制表位与左缩进伴随值
    必须写绝对 twips，按【文档字符单位字号】换算（陷阱 #12）。"""

    def setUp(self):
        self.mod = helpers.load_script("40_apply_fixes.py")
        from lxml import etree
        self.etree = etree

    def _ppr(self, ind_xml=""):
        w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
        return self.etree.fromstring(
            ('<w:pPr xmlns:w="%s">%s</w:pPr>' % (w, ind_xml)).encode("utf-8"))

    def _ind(self, pPr):
        return pPr.find(self.mod.qn("w:ind"))

    def test_char_twips_uses_char_unit_size(self):
        # 2 chars at the DEFAULT char-unit 五号(21 半点=10.5pt) = 420 twips
        # (=0.74cm) — the value Word actually renders. At 三号(32) it would be
        # 640 (=1.13cm), which is the over-indent bug we must NOT produce.
        self.assertEqual(self.mod._char_twips(200, 21), 420)
        self.assertEqual(self.mod._char_twips(400, 21), 840)  # 4 chars = 1.49cm
        self.assertEqual(self.mod._char_twips(200, 32), 640)  # 16pt over-indents
        self.assertEqual(self.mod._char_twips(0, 21), 0)

    def test_default_char_unit_hp_reads_docdefaults_else_21(self):
        import tempfile as _tf
        import os as _os
        w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
        d = _tf.mkdtemp()
        _os.makedirs(_os.path.join(d, "word"))
        # no docDefaults size -> fall back to 21 (五号)
        with open(_os.path.join(d, "word", "styles.xml"), "w", encoding="utf-8") as f:
            f.write('<w:styles xmlns:w="%s"></w:styles>' % w)
        self.assertEqual(self.mod._default_char_unit_hp(d), 21)
        # explicit docDefaults size wins
        with open(_os.path.join(d, "word", "styles.xml"), "w", encoding="utf-8") as f:
            f.write('<w:styles xmlns:w="%s"><w:docDefaults><w:rPrDefault>'
                    '<w:rPr><w:sz w:val="24"/></w:rPr></w:rPrDefault>'
                    '</w:docDefaults></w:styles>' % w)
        self.assertEqual(self.mod._default_char_unit_hp(d), 24)
        shutil.rmtree(d, ignore_errors=True)

    def test_first_line_is_char_only(self):
        pPr = self._ppr()
        self.mod._set_first_line_and_clear_left(pPr, 200, False)
        ind = self._ind(pPr)
        self.assertEqual(ind.get(self.mod.qn("w:firstLineChars")), "200")
        # 无绝对伴随值——有它 Word 就把首行缩进显示成厘米而不是"2 字符"
        self.assertIsNone(ind.get(self.mod.qn("w:firstLine")))

    def test_direct_hanging_removed_when_setting_first_line(self):
        # 段落带【绝对 hanging】（LibreOffice 形态）时，首行缩进写字符单位、hanging 清掉
        pPr = self._ppr('<w:ind w:hanging="420" w:left="420"/>')
        self.mod._set_first_line_and_clear_left(pPr, 200, True)
        ind = self._ind(pPr)
        self.assertIsNone(ind.get(self.mod.qn("w:hanging")))
        self.assertIsNone(ind.get(self.mod.qn("w:firstLine")))
        self.assertEqual(ind.get(self.mod.qn("w:firstLineChars")), "200")
        # 要清零的左缩进写显式 0（挡住继承值），零值不涉及单位歧义
        self.assertEqual(ind.get(self.mod.qn("w:leftChars")), "0")
        self.assertEqual(ind.get(self.mod.qn("w:left")), "0")

    def test_clear_no_indent_zeroes_all_directions(self):
        pPr = self._ppr('<w:ind w:hanging="420"/>')
        self.mod._set_first_line_and_clear_left(pPr, 0, True, True)
        ind = self._ind(pPr)
        self.assertEqual(ind.get(self.mod.qn("w:firstLineChars")), "0")
        self.assertIsNone(ind.get(self.mod.qn("w:hanging")))
        self.assertEqual(ind.get(self.mod.qn("w:left")), "0")
        self.assertEqual(ind.get(self.mod.qn("w:right")), "0")

    def test_left_chars_is_char_only(self):
        # 目录条目的按级左缩进：只写 leftChars，绝对 left 被删掉
        pPr = self._ppr('<w:ind w:left="999"/>')
        self.mod._set_first_line_and_clear_left(pPr, None, False, False, 200)
        ind = self._ind(pPr)
        self.assertEqual(ind.get(self.mod.qn("w:leftChars")), "200")
        self.assertIsNone(ind.get(self.mod.qn("w:left")))


# Docx parts for the LibreOffice-shaped heading test (numbering-level hanging).
_CT_LO = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
    '<Default Extension="xml" ContentType="application/xml"/>'
    '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
    '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>'
    '<Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/>'
    '</Types>'
)
_DRELS_LO = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>'
    '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/>'
    '</Relationships>'
)
_STYLES_LO = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
    '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>'
    '<w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/>'
    '<w:basedOn w:val="Normal"/><w:pPr><w:outlineLvl w:val="0"/></w:pPr></w:style>'
    '</w:styles>'
)
# The heading number comes from an automatic list whose indent is an ABSOLUTE
# hanging (420 twips ≈ 0.74cm) — exactly the shape a LibreOffice conversion
# emits, and the source of the "-0.74cm 悬挂缩进" symptom.
_NUM_LO = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
    '<w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0"><w:start w:val="1"/>'
    '<w:numFmt w:val="decimal"/><w:lvlText w:val="%1、"/>'
    '<w:pPr><w:ind w:left="420" w:hanging="420"/></w:pPr></w:lvl></w:abstractNum>'
    '<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>'
    '</w:numbering>'
)


@unittest.skipUnless(HAVE_LXML, "lxml not installed")
class TestLibreOfficeHeadingIndentPipeline(unittest.TestCase):
    """整链回归（方案C 甲法）：一个从编号层继承【绝对 hanging】的自动编号标题，跑完
    20→30→40 后，编号层的 hanging 被【克隆钳住】——段落 numPr 改指一个新克隆的
    numId、克隆级别不再有 hanging——于是段落自身写【纯字符单位首行缩进】
    （firstLineChars，无绝对伴随值，严格-spec §2.2）即可显示"2字符"而非 -0.74cm 悬挂。
    原共享 abstractNum 不被原地改（§2.5/#17）。"""

    def setUp(self):
        self.tmp = tempfile.mkdtemp(prefix="wrf_lo_")

    def tearDown(self):
        shutil.rmtree(self.tmp, ignore_errors=True)

    def _script(self, name):
        return os.path.join(helpers.SCRIPTS, name)

    def _build(self, path, second_heading=False):
        import zipfile
        extra = ""
        if second_heading:
            # 第二个自动编号标题，共享同一 numId=1（考验：钳一个别溢到另一个/共享安全）
            extra = (
                '<w:p><w:pPr><w:pStyle w:val="Heading1"/>'
                '<w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr></w:pPr>'
                '<w:r><w:rPr><w:rFonts w:eastAsia="宋体"/><w:sz w:val="32"/></w:rPr>'
                '<w:t xml:space="preserve">背景</w:t></w:r></w:p>')
        doc = (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>'
            '<w:p><w:pPr><w:pStyle w:val="Heading1"/>'
            '<w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr></w:pPr>'
            '<w:r><w:rPr><w:rFonts w:eastAsia="宋体"/><w:sz w:val="32"/></w:rPr>'
            '<w:t xml:space="preserve">概述</w:t></w:r></w:p>'
            + extra +
            '<w:p><w:r><w:rPr><w:rFonts w:eastAsia="宋体"/><w:sz w:val="32"/></w:rPr>'
            '<w:t xml:space="preserve">正文内容。</w:t></w:r></w:p>'
            '</w:body></w:document>'
        )
        with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
            z.writestr("[Content_Types].xml", _CT_LO)
            z.writestr("_rels/.rels", helpers.ROOT_RELS)
            z.writestr("word/_rels/document.xml.rels", _DRELS_LO)
            z.writestr("word/document.xml", doc)
            z.writestr("word/styles.xml", _STYLES_LO)
            z.writestr("word/numbering.xml", _NUM_LO)
        return path

    def _w(self):
        w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
        return w, (lambda t: "{%s}%s" % (w, t.split(":")[1]))

    def _run_pipeline(self, src):
        base = os.path.join(self.tmp, "wb")
        wd = run(self._script("05_new_workdir.py"), base)["workdir"]
        run(self._script("10_prepare_input.py"), src, wd)
        run(self._script("20_extract_structure.py"), wd)
        run(self._script("30_check_format.py"), wd)
        run(self._script("40_apply_fixes.py"), wd)
        return wd

    def test_numbering_hanging_clamped_via_clone(self):
        from lxml import etree
        src = self._build(os.path.join(self.tmp, "lo.docx"))
        wd = self._run_pipeline(src)
        _w, qn = self._w()

        root = etree.parse(os.path.join(wd, "out_pkg", "word", "document.xml")).getroot()
        heading_p = root.iter(qn("w:p")).__next__()  # first body paragraph = 概述 heading

        # 阶段2：缩进由【注入的 canonical 标题样式】承载，段落上不留直接 w:ind
        # （严格-spec §6：canonical 值必须由命名样式供给、直接覆盖清掉）。
        pstyle = heading_p.find(qn("w:pPr") + "/" + qn("w:pStyle"))
        self.assertIsNotNone(pstyle, "标题段落未指派 canonical 样式")
        self.assertEqual(pstyle.get(qn("w:val")), "FGWCanonH1")
        self.assertIsNone(heading_p.find(qn("w:pPr") + "/" + qn("w:ind")),
                          "canonical 样式承载缩进后，段落不应再有直接 w:ind")

        styles_root = etree.parse(os.path.join(wd, "out_pkg", "word", "styles.xml")).getroot()
        h1 = [s for s in styles_root.findall(qn("w:style"))
              if s.get(qn("w:styleId")) == "FGWCanonH1"][0]
        sind = h1.find(qn("w:pPr") + "/" + qn("w:ind"))
        # 首行缩进2字符：纯字符单位
        self.assertEqual(sind.get(qn("w:firstLineChars")), "200")
        # 严格-spec：不写非零绝对伴随值（编号层已钳，无需绝对值压制）；
        # 左右显式归零只是挡继承，零值没有单位歧义。
        self.assertIsNone(sind.get(qn("w:firstLine")))
        self.assertEqual(sind.get(qn("w:left")), "0")
        self.assertIsNone(sind.get(qn("w:hanging")))

        # 段落 numPr 已改指一个【新克隆的 numId】(不再是原 numId=1)
        numid = heading_p.find(qn("w:pPr") + "/" + qn("w:numPr") + "/" + qn("w:numId"))
        self.assertIsNotNone(numid)
        new_numid = numid.get(qn("w:val"))
        self.assertNotEqual(new_numid, "1")

        # numbering.xml：原 abstractNum 0 未被原地改（hanging 仍在）；克隆级别无 hanging
        num_root = etree.parse(os.path.join(wd, "out_pkg", "word", "numbering.xml")).getroot()
        abs0 = [a for a in num_root.findall(qn("w:abstractNum"))
                if a.get(qn("w:abstractNumId")) == "0"][0]
        ind0 = abs0.find(qn("w:lvl") + "/" + qn("w:pPr") + "/" + qn("w:ind"))
        self.assertEqual(ind0.get(qn("w:hanging")), "420", "共享 abstractNum 被原地改了（#17 回退）")
        # 新 numId → 其 abstractNum，克隆级别 hanging 已去除
        num2abs = {n.get(qn("w:numId")): n.find(qn("w:abstractNumId")).get(qn("w:val"))
                   for n in num_root.findall(qn("w:num"))}
        clone_aid = num2abs[new_numid]
        clone = [a for a in num_root.findall(qn("w:abstractNum"))
                 if a.get(qn("w:abstractNumId")) == clone_aid][0]
        cind = clone.find(qn("w:lvl") + "/" + qn("w:pPr") + "/" + qn("w:ind"))
        self.assertIsNone(cind.get(qn("w:hanging")))

    def test_extraction_surfaces_numbering_indent_over_style_indent(self):
        """检测接入 resolve_cascade（关键：竞争同一个键）：标题样式自身设了 left=200，
        编号层同样设了 left=420（+hanging=420）。正确优先级下编号层压过样式 →
        eff.left=420；旧 gap-fill 只兜 None，样式的 left=200 非 None 就把编号层的 left
        挡掉、eff.left 停在 200，从而漏看真实缩进（方案C §2.3）。断言 eff.left=420 才
        证明编号层被折进有效值、且压过了样式。"""
        import zipfile
        styles = (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
            '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>'
            '<w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/>'
            '<w:basedOn w:val="Normal"/>'
            # 样式自身带 left=200（与编号层的 left 竞争同一个键）
            '<w:pPr><w:outlineLvl w:val="0"/><w:ind w:left="200"/></w:pPr></w:style>'
            '</w:styles>')
        src = os.path.join(self.tmp, "lo_style_indent.docx")
        doc = (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>'
            '<w:p><w:pPr><w:pStyle w:val="Heading1"/>'
            '<w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr></w:pPr>'
            '<w:r><w:rPr><w:rFonts w:eastAsia="宋体"/><w:sz w:val="32"/></w:rPr>'
            '<w:t xml:space="preserve">概述</w:t></w:r></w:p>'
            '<w:p><w:r><w:rPr><w:rFonts w:eastAsia="宋体"/><w:sz w:val="32"/></w:rPr>'
            '<w:t xml:space="preserve">正文内容。</w:t></w:r></w:p>'
            '</w:body></w:document>')
        with zipfile.ZipFile(src, "w", zipfile.ZIP_DEFLATED) as z:
            z.writestr("[Content_Types].xml", _CT_LO)
            z.writestr("_rels/.rels", helpers.ROOT_RELS)
            z.writestr("word/_rels/document.xml.rels", _DRELS_LO)
            z.writestr("word/document.xml", doc)
            z.writestr("word/styles.xml", styles)
            z.writestr("word/numbering.xml", _NUM_LO)   # abstractNum 0: left=420 hanging=420

        base = os.path.join(self.tmp, "wbse")
        wd = run(self._script("05_new_workdir.py"), base)["workdir"]
        run(self._script("10_prepare_input.py"), src, wd)
        run(self._script("20_extract_structure.py"), wd)
        with open(os.path.join(wd, "structure.json"), encoding="utf-8") as f:
            recs = json.load(f)["records"]
        heading = recs[0]
        self.assertTrue(heading["auto_num"])
        # 编号层压过样式：竞争键 left = 420（非样式的 200）；hanging 也 surface
        self.assertEqual(heading["eff"]["left"], 420)
        self.assertEqual(heading["eff"]["hanging"], 420)

    def test_shared_abstractnum_not_mutated(self):
        """两个共享 numId=1 的标题都被钳：都改指同一新 numId，原 abstractNum 0 不变。"""
        from lxml import etree
        src = self._build(os.path.join(self.tmp, "lo2.docx"), second_heading=True)
        wd = self._run_pipeline(src)
        _w, qn = self._w()

        root = etree.parse(os.path.join(wd, "out_pkg", "word", "document.xml")).getroot()
        paras = list(root.iter(qn("w:p")))
        h1, h2 = paras[0], paras[1]  # 概述 / 背景
        n1 = h1.find(qn("w:pPr") + "/" + qn("w:numPr") + "/" + qn("w:numId")).get(qn("w:val"))
        n2 = h2.find(qn("w:pPr") + "/" + qn("w:numPr") + "/" + qn("w:numId")).get(qn("w:val"))
        # 同组共享同一克隆（只克隆一次），且都不是原 numId
        self.assertEqual(n1, n2)
        self.assertNotEqual(n1, "1")

        num_root = etree.parse(os.path.join(wd, "out_pkg", "word", "numbering.xml")).getroot()
        abs0 = [a for a in num_root.findall(qn("w:abstractNum"))
                if a.get(qn("w:abstractNumId")) == "0"][0]
        ind0 = abs0.find(qn("w:lvl") + "/" + qn("w:pPr") + "/" + qn("w:ind"))
        self.assertEqual(ind0.get(qn("w:hanging")), "420")
        # 原 abstractNum 0 只被克隆【一次】：两个标题共享一条 abstractNum，就该合成
        # 一组、共用一个新 numId——分开克隆会把一个多级列表劈成几条独立列表，二/三/
        # 四级标题的编号从中间重新计数（用户实测的"标题编号顺序出错"）。
        # 另外两条是图/表标题的编号定义（无条件注入，靠 w:name 认领）。
        names = [a.find(qn("w:name")).get(qn("w:val")) if a.find(qn("w:name")) is not None
                 else None for a in num_root.findall(qn("w:abstractNum"))]
        self.assertEqual(sorted(n for n in names if n),
                         ["FGWCaptionFigure", "FGWCaptionTable"])
        self.assertEqual(len([n for n in names if n is None]), 2,
                         "原 abstractNum + 一份克隆 = 2 条（克隆多于一份即分组错了）")


@unittest.skipUnless(HAVE_LXML, "lxml not installed")
class TestCollapseInvariant(unittest.TestCase):
    """45 的方案C 甲法坍缩不变量：拿到缩进修复的自动编号段，有效 hanging 不得仍由
    编号层供给。用 provenance 区分致命泄漏（numbering）与非致命提示（style）。"""

    W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

    def setUp(self):
        self.tmp = tempfile.mkdtemp(prefix="wrf_ci_")
        self.mod = helpers.load_script("45_validate_output.py")

    def tearDown(self):
        shutil.rmtree(self.tmp, ignore_errors=True)

    def _pkg(self, name, styles, numbering):
        import zipfile
        doc = (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<w:document xmlns:w="%s"><w:body>'
            '<w:p><w:pPr><w:pStyle w:val="Heading1"/>'
            '<w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr></w:pPr>'
            '<w:r><w:t xml:space="preserve">概述</w:t></w:r></w:p>'
            '<w:p><w:r><w:t xml:space="preserve">正文。</w:t></w:r></w:p>'
            '</w:body></w:document>' % self.W)
        path = os.path.join(self.tmp, name)
        with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
            z.writestr("[Content_Types].xml", _CT_LO)
            z.writestr("_rels/.rels", helpers.ROOT_RELS)
            z.writestr("word/_rels/document.xml.rels", _DRELS_LO)
            z.writestr("word/document.xml", doc)
            z.writestr("word/styles.xml", styles)
            z.writestr("word/numbering.xml", numbering)
        return path

    _STYLES_MIN = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
        '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>'
        '<w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/>'
        '<w:pPr><w:outlineLvl w:val="0"/></w:pPr></w:style></w:styles>')

    def _num(self, lvl_ind):
        return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                '<w:numbering xmlns:w="%s">'
                '<w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0"><w:start w:val="1"/>'
                '<w:numFmt w:val="decimal"/><w:lvlText w:val="%%1、"/>'
                '<w:pPr><w:ind %s/></w:pPr></w:lvl></w:abstractNum>'
                '<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num></w:numbering>'
                % (self.W, lvl_ind))

    def _leaks(self, path):
        import zipfile
        with zipfile.ZipFile(path) as zf:
            return self.mod._check_numbering_clamp(zf, [0])

    def test_numbering_hanging_is_a_leak(self):
        # 编号层仍有 hanging，段落 numPr 仍指它 → 有效 hanging 由 numbering 供给 = 致命泄漏
        path = self._pkg("leak.docx", self._STYLES_MIN, self._num('w:left="420" w:hanging="420"'))
        leaks, notes = self._leaks(path)
        self.assertEqual(len(leaks), 1)
        self.assertEqual(leaks[0]["owner"], "numbering")
        self.assertEqual(leaks[0]["key"], "hanging")

    def test_clamped_numbering_no_leak(self):
        # 编号层已中和（无 hanging，left=0）→ 无泄漏
        path = self._pkg("ok.docx", self._STYLES_MIN, self._num('w:left="0" w:leftChars="0"'))
        leaks, notes = self._leaks(path)
        self.assertEqual(leaks, [])
        self.assertEqual(notes, [])

    def test_style_hanging_is_note_not_leak(self):
        # hanging 来自【样式】而非编号层 → 非致命 note（已知未修，不阻断交付）
        styles = (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<w:styles xmlns:w="%s">'
            '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>'
            '<w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/>'
            '<w:pPr><w:outlineLvl w:val="0"/><w:ind w:hanging="300"/></w:pPr></w:style></w:styles>' % self.W)
        path = self._pkg("note.docx", styles, self._num('w:left="0" w:leftChars="0"'))
        leaks, notes = self._leaks(path)
        self.assertEqual(leaks, [])
        self.assertEqual(len(notes), 1)
        self.assertEqual(notes[0]["owner"], "style")


W_NS_ = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"


@unittest.skipUnless(HAVE_LXML, "lxml not installed")
class TestCanonicalStyleInjection(unittest.TestCase):
    """方案C 阶段2 主体：全角色 canonical 样式**注入 + 指派 + 清直接覆盖**，
    外加文档网格 / 表格默认值。每条锁一个"别回退"的决策。"""

    W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

    def setUp(self):
        self.tmp = tempfile.mkdtemp(prefix="wrf_canon_")
        with open(helpers.SPEC_PATH, encoding="utf-8") as f:
            self.spec = json.load(f)

    def tearDown(self):
        shutil.rmtree(self.tmp, ignore_errors=True)

    def _script(self, name):
        return os.path.join(helpers.SCRIPTS, name)

    def qn(self, t):
        return "{%s}%s" % (self.W, t.split(":")[1])

    def _run(self, body, pg_mar=(1000, 1814, 1616, 1616, 850, 992)):
        src = os.path.join(self.tmp, "in.docx")
        helpers.build_docx(src, body, pg_mar=pg_mar)
        wd = run(self._script("05_new_workdir.py"), os.path.join(self.tmp, "wb"))["workdir"]
        run(self._script("10_prepare_input.py"), src, wd)
        run(self._script("20_extract_structure.py"), wd)
        run(self._script("30_check_format.py"), wd)
        applied = run(self._script("40_apply_fixes.py"), wd)
        return wd, applied

    def _doc(self, wd):
        from lxml import etree
        return etree.parse(os.path.join(wd, "out_pkg", "word", "document.xml")).getroot()

    def _styles(self, wd):
        from lxml import etree
        return etree.parse(os.path.join(wd, "out_pkg", "word", "styles.xml")).getroot()

    _MIXED_BODY = (
        helpers.para("先进项目2024年度自评价报告", east_asia="宋体", size_hp=44, jc="center")
        + helpers.para("项目编号：KJ-2024-001", east_asia="宋体", size_hp=30)
        + helpers.para("一、绪论", east_asia="宋体", size_hp=32, outline=0)
        + helpers.para("这是正文内容含数字123。", east_asia="宋体", size_hp=32)
        + helpers.para("图1 系统架构", east_asia="宋体", size_hp=32, style="图标题")
        + '<w:tbl><w:tblPr/><w:tr><w:tc><w:tcPr/>'
        + helpers.para("单元格", east_asia="宋体", size_hp=32) + '</w:tc></w:tr></w:tbl>')

    def test_each_role_gets_its_canonical_style(self):
        wd, _ = self._run(self._MIXED_BODY)
        got = [p.find(self.qn("w:pPr") + "/" + self.qn("w:pStyle"))
               for p in self._doc(wd).iter(self.qn("w:p"))]
        got = [g.get(self.qn("w:val")) if g is not None else None for g in got]
        self.assertEqual(got, ["FGWCanonTitle", "FGWCanonCoverField", "FGWCanonH1",
                               "FGWCanonBody", "FGWCanonCaptionFig",
                               "FGWCanonTableBody"])

    def test_pattern_caption_is_not_assigned_a_style(self):
        """安全阀（陷阱#5）：仅凭"图+数字"形状认出、没有题注样式撑腰的图表标题**不**
        指派 canonical 样式——否则下一轮它就凭样式变成"已确认"，绕过安全阀被自动改
        编号，而它可能只是一句以"图3 显示了…"开头的正文。"""
        body = (helpers.para("一、绪论", east_asia="宋体", size_hp=32, outline=0)
                + helpers.para("图3 显示了系统架构。", east_asia="宋体", size_hp=32))
        wd, _ = self._run(body)
        paras = list(self._doc(wd).iter(self.qn("w:p")))
        pstyle = paras[1].find(self.qn("w:pPr") + "/" + self.qn("w:pStyle"))
        self.assertIsNone(pstyle)
        self.assertIsNone(paras[1].find(self.qn("w:pPr") + "/" + self.qn("w:numPr")))

    def test_direct_overrides_are_cleared_not_rewritten(self):
        # 严格-spec §6：canonical 值由样式承载，段落上不再留同一属性的直接覆盖
        # （"渲染对但用直接属性表达"不算合规）。
        wd, _ = self._run(self._MIXED_BODY)
        body_p = [p for p in self._doc(wd).iter(self.qn("w:p"))
                  if (p.find(self.qn("w:pPr") + "/" + self.qn("w:pStyle")) is not None
                      and p.find(self.qn("w:pPr") + "/" + self.qn("w:pStyle")).get(
                          self.qn("w:val")) == "FGWCanonBody")][0]
        self.assertIsNone(body_p.find(self.qn("w:pPr") + "/" + self.qn("w:ind")))
        self.assertIsNone(body_p.find(self.qn("w:pPr") + "/" + self.qn("w:spacing")))
        for r in body_p.iter(self.qn("w:r")):
            rpr = r.find(self.qn("w:rPr"))
            if rpr is None:
                continue
            self.assertIsNone(rpr.find(self.qn("w:sz")))
            rf = rpr.find(self.qn("w:rFonts"))
            if rf is not None:
                self.assertIsNone(rf.get(self.qn("w:eastAsia")))

    def test_style_carries_the_spec_values(self):
        wd, _ = self._run(self._MIXED_BODY)
        styles = {s.get(self.qn("w:styleId")): s
                  for s in self._styles(wd).findall(self.qn("w:style"))}
        body = styles["FGWCanonBody"]
        rf = body.find(self.qn("w:rPr") + "/" + self.qn("w:rFonts"))
        self.assertEqual(rf.get(self.qn("w:eastAsia")), self.spec["body"]["east_asia"])
        self.assertEqual(rf.get(self.qn("w:ascii")), self.spec["western_font"])
        self.assertEqual(body.find(self.qn("w:rPr") + "/" + self.qn("w:sz")).get(
            self.qn("w:val")), str(self.spec["body"]["size_hp"]))
        sp = body.find(self.qn("w:pPr") + "/" + self.qn("w:spacing"))
        self.assertEqual(sp.get(self.qn("w:line")),
                         str(self.spec["line_spacing"]["line_twips"]))
        self.assertEqual(sp.get(self.qn("w:lineRule")), "exact")

    def test_normal_pinned_to_wuhao_for_document_grid(self):
        # 陷阱#12：Word 的【文档网格字体】＝Normal 样式字号，必须是五号(21)，
        # 否则行网格 15.6磅/41行 被顶成 21.75磅/29行。正文三号由 FGW正文 承载。
        wd, _ = self._run(self._MIXED_BODY)
        root = self._styles(wd)
        normal = [s for s in root.findall(self.qn("w:style"))
                  if s.get(self.qn("w:styleId")) == "Normal"][0]
        self.assertEqual(normal.find(self.qn("w:rPr") + "/" + self.qn("w:sz")).get(
            self.qn("w:val")), str(self.spec["document_grid"]["normal_size_hp"]))
        dd = root.find(self.qn("w:docDefaults") + "/" + self.qn("w:rPrDefault")
                       + "/" + self.qn("w:rPr") + "/" + self.qn("w:sz"))
        self.assertEqual(dd.get(self.qn("w:val")),
                         str(self.spec["document_grid"]["doc_defaults_size_hp"]))

    def test_compliant_paragraph_still_assigned_so_normal_change_cant_shrink_it(self):
        """反回退：把 Normal 钉成五号会让"原本靠 Normal 拿到三号的合规段落"悄悄变小。

        所以指派覆盖**全部有角色的段落**，不只违规段落——否则就出现"改了却没提示"的
        盲区。这里造一个各项都合规、字号来自 Normal 的正文段：它没有 fix，但必须被
        指派 FGW正文（三号）而不是跟着 Normal 掉到五号。"""
        import zipfile
        styles = (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<w:styles xmlns:w="%s">'
            '<w:style w:type="paragraph" w:default="1" w:styleId="Normal">'
            '<w:name w:val="Normal"/><w:pPr>'
            '<w:spacing w:line="560" w:lineRule="exact" w:before="0" w:after="0"/>'
            '<w:ind w:firstLineChars="200"/></w:pPr>'
            '<w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" '
            'w:eastAsia="仿宋"/><w:sz w:val="32"/><w:szCs w:val="32"/></w:rPr>'
            '</w:style></w:styles>' % self.W)
        doc = helpers.document_xml(
            '<w:p><w:r><w:t xml:space="preserve">完全合规的正文段落。</w:t></w:r></w:p>',
            pg_mar=(1984, 1814, 1616, 1616, 850, 992))
        src = os.path.join(self.tmp, "compliant.docx")
        with zipfile.ZipFile(src, "w", zipfile.ZIP_DEFLATED) as z:
            z.writestr("[Content_Types].xml", _CT_LO)
            z.writestr("_rels/.rels", helpers.ROOT_RELS)
            z.writestr("word/_rels/document.xml.rels", _DRELS_LO)
            z.writestr("word/document.xml", doc)
            z.writestr("word/styles.xml", styles)
            z.writestr("word/numbering.xml",
                       '<w:numbering xmlns:w="%s"/>' % self.W)
        wd = run(self._script("05_new_workdir.py"), os.path.join(self.tmp, "wb2"))["workdir"]
        run(self._script("10_prepare_input.py"), src, wd)
        run(self._script("20_extract_structure.py"), wd)
        checked = run(self._script("30_check_format.py"), wd)
        self.assertEqual(checked["by_op"].get("format", 0), 0, "该段本应完全合规")
        run(self._script("40_apply_fixes.py"), wd)
        p = list(self._doc(wd).iter(self.qn("w:p")))[0]
        pstyle = p.find(self.qn("w:pPr") + "/" + self.qn("w:pStyle"))
        self.assertIsNotNone(pstyle, "合规段落也必须被指派 canonical 样式")
        self.assertEqual(pstyle.get(self.qn("w:val")), "FGWCanonBody")

    def test_document_grid_and_compat_written(self):
        # 网格 15.6磅/41行 靠 sectPr 的 docGrid ＋ settings 的 compat 块共同生效，
        # 缺 compat（尤其 useFELayout/compatibilityMode=15）Word 会退回旧版式。
        wd, applied = self._run(self._MIXED_BODY)
        from lxml import etree
        grid = self.spec["document_grid"]
        sect = self._doc(wd).find(".//" + self.qn("w:sectPr") + "/" + self.qn("w:docGrid"))
        self.assertEqual(sect.get(self.qn("w:type")), grid["grid_type"])
        self.assertEqual(sect.get(self.qn("w:linePitch")), str(grid["line_pitch_twips"]))
        st = etree.parse(os.path.join(wd, "out_pkg", "word", "settings.xml")).getroot()
        self.assertEqual(st.find(self.qn("w:defaultTabStop")).get(self.qn("w:val")),
                         str(grid["default_tab_stop_twips"]))
        compat = st.find(self.qn("w:compat"))
        self.assertIsNotNone(compat.find(self.qn("w:useFELayout")))
        modes = {c.get(self.qn("w:name")): c.get(self.qn("w:val"))
                 for c in compat.findall(self.qn("w:compatSetting"))}
        self.assertEqual(modes.get("compatibilityMode"),
                         grid["compat_settings"]["compatibilityMode"])

    def test_table_defaults_applied(self):
        wd, _ = self._run(self._MIXED_BODY)
        td = self.spec["table_defaults"]
        tbl = self._doc(wd).find(".//" + self.qn("w:tbl"))
        self.assertEqual(tbl.find(self.qn("w:tblPr") + "/" + self.qn("w:tblInd")).get(
            self.qn("w:w")), str(td["tbl_ind_twips"]))
        mar = tbl.find(self.qn("w:tblPr") + "/" + self.qn("w:tblCellMar"))
        self.assertEqual(mar.find(self.qn("w:left")).get(self.qn("w:w")),
                         str(td["cell_margin_twips"]["left"]))
        self.assertEqual(
            tbl.find(".//" + self.qn("w:tc") + "/" + self.qn("w:tcPr") + "/"
                     + self.qn("w:vAlign")).get(self.qn("w:val")), td["cell_valign"])

    def test_injection_is_idempotent(self):
        # 重跑收敛：同 styleId 整体替换，不会越注入越多。
        wd, applied = self._run(self._MIXED_BODY)
        n1 = len(self._styles(wd).findall(self.qn("w:style")))
        run(self._script("40_apply_fixes.py"), wd)
        self.assertEqual(len(self._styles(wd).findall(self.qn("w:style"))), n1)

    def test_full_collapse_invariant_passes_on_real_output(self):
        # 45 的全坍缩不变量：真实产物上，canonical 属性必须全部由样式层供给。
        wd, _ = self._run(self._MIXED_BODY)
        validated = run(self._script("45_validate_output.py"), wd)
        self.assertTrue(validated["ok"], validated.get("errors"))
        self.assertNotIn("canonical_leaks", validated)

    def test_full_collapse_invariant_catches_a_direct_leak(self):
        # 人为把一个直接缩进塞回 canonical 段落 → 必须硬失败（退2），不能放行。
        import zipfile
        sys_path = os.path.join(helpers.SCRIPTS, "lib")
        if sys_path not in sys.path:
            sys.path.insert(0, sys_path)
        from docxcommon import rezip_docx
        wd, _ = self._run(self._MIXED_BODY)
        doc_path = os.path.join(wd, "out_pkg", "word", "document.xml")
        with open(doc_path, encoding="utf-8") as f:
            s = f.read().replace(
                '<w:pStyle w:val="FGWCanonBody"/>',
                '<w:pStyle w:val="FGWCanonBody"/><w:ind w:hanging="420"/>', 1)
        with open(doc_path, "w", encoding="utf-8") as f:
            f.write(s)
        rezip_docx(os.path.join(wd, "out_pkg"), os.path.join(wd, "formatted.docx"))
        proc = subprocess.run(
            [sys.executable, self._script("45_validate_output.py"), wd],
            capture_output=True, text=True)
        self.assertEqual(proc.returncode, 2, proc.stdout)
        report = json.loads(proc.stdout.strip().splitlines()[-1])
        self.assertTrue(any("坍缩" in e for e in report["errors"]), report["errors"])

    # 一段"半粗半细"的三级标题：前半 run 直接加粗 + 挂了一个设了字体/加粗的**字符样式**，
    # 后半 run 显式取消加粗。这是用户阶段2 验收发现的真实形态——字符样式压过段落样式，
    # 于是 canonical 标题样式看起来"只应用到了后半部分"。
    _SPLIT_BOLD_STYLES = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<w:styles xmlns:w="%s">'
        '<w:style w:type="paragraph" w:default="1" w:styleId="Normal">'
        '<w:name w:val="Normal"/></w:style>'
        '<w:style w:type="paragraph" w:styleId="Heading3"><w:name w:val="heading 3"/>'
        '<w:pPr><w:outlineLvl w:val="2"/></w:pPr></w:style>'
        '<w:style w:type="character" w:styleId="StrongCS"><w:name w:val="Strong"/>'
        '<w:rPr><w:rFonts w:eastAsia="黑体"/><w:b/></w:rPr></w:style>'
        '<w:style w:type="character" w:styleId="LinkCS"><w:name w:val="Hyperlink"/>'
        '<w:rPr><w:color w:val="0563C1"/><w:u w:val="single"/></w:rPr></w:style>'
        '</w:styles>' % W_NS_)

    def _build_split_bold(self, name="split.docx"):
        import zipfile
        h3 = ('<w:p><w:pPr><w:pStyle w:val="Heading3"/></w:pPr>'
              '<w:r><w:rPr><w:rStyle w:val="StrongCS"/><w:rFonts w:eastAsia="宋体"/>'
              '<w:b/><w:sz w:val="32"/></w:rPr>'
              '<w:t xml:space="preserve">1. 数据来源</w:t></w:r>'
              '<w:r><w:rPr><w:rFonts w:eastAsia="宋体"/><w:b w:val="0"/>'
              '<w:sz w:val="32"/></w:rPr>'
              '<w:t xml:space="preserve">与口径说明</w:t></w:r></w:p>')
        # 只设颜色/下划线的字符样式（超链接）必须**留着**——样式没管的作者格式不动
        body_p = ('<w:p><w:r><w:rPr><w:rStyle w:val="LinkCS"/><w:rFonts w:eastAsia="宋体"/>'
                  '<w:sz w:val="32"/></w:rPr>'
                  '<w:t xml:space="preserve">正文里的一个链接。</w:t></w:r></w:p>')
        src = os.path.join(self.tmp, name)
        with zipfile.ZipFile(src, "w", zipfile.ZIP_DEFLATED) as z:
            z.writestr("[Content_Types].xml", _CT_LO)
            z.writestr("_rels/.rels", helpers.ROOT_RELS)
            z.writestr("word/_rels/document.xml.rels", _DRELS_LO)
            z.writestr("word/document.xml", helpers.document_xml(h3 + body_p))
            z.writestr("word/styles.xml", self._SPLIT_BOLD_STYLES)
            z.writestr("word/numbering.xml", '<w:numbering xmlns:w="%s"/>' % self.W)
        return src

    def _run_src(self, src, wbname):
        wd = run(self._script("05_new_workdir.py"),
                 os.path.join(self.tmp, wbname))["workdir"]
        run(self._script("10_prepare_input.py"), src, wd)
        run(self._script("20_extract_structure.py"), wd)
        run(self._script("30_check_format.py"), wd)
        run(self._script("40_apply_fixes.py"), wd)
        return wd

    def test_split_run_heading_becomes_uniformly_bold(self):
        """标题加粗由样式承载，run 上的直接加粗/取消加粗/抢戏字符样式全部清掉。

        反回退：别只清 `w:b`——真正让后半段不粗的是 `w:b w:val="0"`（显式取消），
        还有那个设了字体+加粗的**字符样式**（压过段落样式）。三者都得清，标题才统一。"""
        wd = self._run_src(self._build_split_bold(), "wbsb")
        h3 = list(self._doc(wd).iter(self.qn("w:p")))[0]
        for r in h3.iter(self.qn("w:r")):
            if not any((t.text or "").strip() for t in r.findall(self.qn("w:t"))):
                continue          # 批注引用 run 不算正文
            rpr = r.find(self.qn("w:rPr"))
            if rpr is None:
                continue
            self.assertIsNone(rpr.find(self.qn("w:b")), "run 上仍有直接加粗设置")
            self.assertIsNone(rpr.find(self.qn("w:bCs")))
            self.assertIsNone(rpr.find(self.qn("w:sz")))
            rs = rpr.find(self.qn("w:rStyle"))
            self.assertNotEqual(rs.get(self.qn("w:val")) if rs is not None else None,
                                "StrongCS", "抢戏的字符样式引用没摘掉")
        style = [s for s in self._styles(wd).findall(self.qn("w:style"))
                 if s.get(self.qn("w:styleId")) == "FGWCanonH3"][0]
        self.assertIsNotNone(style.find(self.qn("w:rPr") + "/" + self.qn("w:b")))
        self.assertIsNotNone(style.find(self.qn("w:rPr") + "/" + self.qn("w:bCs")))

    def test_benign_char_style_is_kept(self):
        # 只设颜色/下划线的字符样式（超链接）不抢 canonical 样式的属性，必须保留——
        # 别把"清直接覆盖"扩大成"把 run 上的东西一律抹掉"。
        wd = self._run_src(self._build_split_bold("split2.docx"), "wbsb2")
        body_p = list(self._doc(wd).iter(self.qn("w:p")))[1]
        styles = [rs.get(self.qn("w:val"))
                  for rs in body_p.iter(self.qn("w:rStyle"))]
        self.assertIn("LinkCS", styles)

    def _build_captions(self, name="cap.docx"):
        import zipfile
        caps = (
            '<w:p><w:pPr><w:pStyle w:val="图标题"/></w:pPr>'
            '<w:r><w:rPr><w:rFonts w:eastAsia="宋体"/><w:sz w:val="32"/></w:rPr>'
            '<w:t xml:space="preserve">图1 系统总体架构</w:t></w:r></w:p>'
            '<w:p><w:pPr><w:pStyle w:val="表标题"/></w:pPr>'
            '<w:r><w:rPr><w:rFonts w:eastAsia="宋体"/><w:sz w:val="32"/></w:rPr>'
            '<w:t xml:space="preserve">表1 主要指标</w:t></w:r></w:p>')
        src = os.path.join(self.tmp, name)
        with zipfile.ZipFile(src, "w", zipfile.ZIP_DEFLATED) as z:
            z.writestr("[Content_Types].xml", _CT_LO)
            z.writestr("_rels/.rels", helpers.ROOT_RELS)
            z.writestr("word/_rels/document.xml.rels", _DRELS_LO)
            z.writestr("word/document.xml", helpers.document_xml(caps))
            z.writestr("word/styles.xml", self._SPLIT_BOLD_STYLES)
            z.writestr("word/numbering.xml", '<w:numbering xmlns:w="%s"/>' % self.W)
        return src

    def test_captions_use_word_autonumbering_with_tab(self):
        """图/表标题改用 Word 自动编号：静态"图N/表N"从文字里删掉、段落挂到注入的
        图%1/表%1 序列上、编号后是**制表符**（suff=tab）。

        别退回静态编号：插入/删除一张图之后所有后续编号都要重排，而 Word 自己维护
        序列——这是用户阶段2 验收明确否掉静态编号的理由。"""
        from lxml import etree
        wd = self._run_src(self._build_captions(), "wbcap")
        paras = list(self._doc(wd).iter(self.qn("w:p")))
        texts = ["".join(t.text or "" for t in p.iter(self.qn("w:t")))
                 for p in paras]
        self.assertEqual(texts[0], "系统总体架构", "静态“图1”没删掉，会和自动编号叠加")
        self.assertEqual(texts[1], "主要指标")

        num_root = etree.parse(os.path.join(wd, "out_pkg", "word",
                                            "numbering.xml")).getroot()
        abs_by_id = {a.get(self.qn("w:abstractNumId")): a
                     for a in num_root.findall(self.qn("w:abstractNum"))}
        num2abs = {n.get(self.qn("w:numId")):
                   n.find(self.qn("w:abstractNumId")).get(self.qn("w:val"))
                   for n in num_root.findall(self.qn("w:num"))}
        # 编号挂在【样式】上，不是段落上——用户才能在 Word 里给新插入的图表标题
        # 套上"图标题"样式就直接生成编号（只挂段落的话新段落没编号）。
        styles = {s.get(self.qn("w:styleId")): s
                  for s in self._styles(wd).findall(self.qn("w:style"))}
        got = {}
        for p, kind, sid in zip(paras, ("图", "表"),
                                ("FGWCanonCaptionFig", "FGWCanonCaptionTbl")):
            self.assertIsNone(
                p.find(self.qn("w:pPr") + "/" + self.qn("w:numPr")),
                "段落上不该留直接 numPr（会压过样式携带的编号）")
            nid = styles[sid].find(self.qn("w:pPr") + "/" + self.qn("w:numPr")
                                   + "/" + self.qn("w:numId"))
            self.assertIsNotNone(nid, "题注样式没带自动编号")
            lvl = abs_by_id[num2abs[nid.get(self.qn("w:val"))]].find(self.qn("w:lvl"))
            got[kind] = (lvl.find(self.qn("w:lvlText")).get(self.qn("w:val")),
                         lvl.find(self.qn("w:numFmt")).get(self.qn("w:val")),
                         lvl.find(self.qn("w:suff")).get(self.qn("w:val")))
        self.assertEqual(got["图"], ("图%1", "decimal", "tab"))
        self.assertEqual(got["表"], ("表%1", "decimal", "tab"))
        # 编号级别缩进必须中和：编号层压过样式层，不中和会盖掉样式的"无缩进"
        for a in abs_by_id.values():
            ind = a.find(self.qn("w:lvl") + "/" + self.qn("w:pPr") + "/"
                         + self.qn("w:ind"))
            self.assertEqual(ind.get(self.qn("w:left")), "0")
            self.assertIsNone(ind.get(self.qn("w:hanging")))

    def test_caption_style_numbering_survives_for_newly_inserted_captions(self):
        # 用户实测："插入新图后我添加标题，应用图标题的样式，没有生成编号"。
        # 编号写进样式后，一个**只套了样式、没有任何直接 numPr** 的新段落也会有编号。
        wd = self._run_src(self._build_captions("cap3.docx"), "wbcap3")
        styles = {s.get(self.qn("w:styleId")): s
                  for s in self._styles(wd).findall(self.qn("w:style"))}
        for sid in ("FGWCanonCaptionFig", "FGWCanonCaptionTbl"):
            numpr = styles[sid].find(self.qn("w:pPr") + "/" + self.qn("w:numPr"))
            self.assertIsNotNone(numpr, "%s 样式必须自带编号" % sid)
            self.assertEqual(numpr.find(self.qn("w:ilvl")).get(self.qn("w:val")), "0")

    def test_caption_numbering_injection_is_idempotent(self):
        # 靠 abstractNum 的 w:name 标记认领已注入的定义，重跑不会越注入越多。
        from lxml import etree
        wd = self._run_src(self._build_captions("cap2.docx"), "wbcap2")

        def n_abs():
            root = etree.parse(os.path.join(wd, "out_pkg", "word",
                                            "numbering.xml")).getroot()
            return len(root.findall(self.qn("w:abstractNum")))
        first = n_abs()
        self.assertEqual(first, 2)
        run(self._script("40_apply_fixes.py"), wd)
        self.assertEqual(n_abs(), first)

    def test_output_has_no_element_order_violations(self):
        """产物不得违反 OOXML 的 xsd:sequence——乱序的良构 XML 会让真实 Word 提示
        "发现无法读取的内容，是否恢复此文档的内容"（用户实测踩过：`w:suff` 写在
        `w:numFmt` 前、`w:jc` 写在 `w:ind` 前、`w:outlineLvl` 写在 `w:spacing` 前、
        往只有 `w:sz` 的 Normal 里 append `w:rFonts`）。"""
        from lxml import etree
        sys.path.insert(0, os.path.join(helpers.SCRIPTS, "lib"))
        from docxcommon import order_violations
        wd, _ = self._run(self._MIXED_BODY)
        for part in ("document.xml", "styles.xml", "numbering.xml", "settings.xml"):
            path = os.path.join(wd, "out_pkg", "word", part)
            if not os.path.exists(path):
                continue
            self.assertEqual(order_violations(etree.parse(path).getroot()), [],
                             "%s 元素顺序违反 schema" % part)

    def test_shared_abstractnum_across_numids_stays_one_list(self):
        """两个 `w:num` 指向同一 `abstractNum`（Word 里极常见）时，克隆钳必须把它们
        **合成一组、只克隆一次**——共享 abstractNum 就是共享计数器，分开克隆会把一个
        多级列表劈成几条独立列表，二/三/四级标题编号从中间重新计数（用户实测的
        "标题编号顺序出错"）。"""
        import zipfile
        from lxml import etree
        lvls = "".join(
            '<w:lvl w:ilvl="%d"><w:start w:val="1"/><w:numFmt w:val="decimal"/>'
            '<w:lvlText w:val="%%%d."/><w:lvlJc w:val="left"/>'
            '<w:pPr><w:ind w:left="420" w:hanging="420"/></w:pPr></w:lvl>' % (i, i + 1)
            for i in range(3))
        numbering = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                     '<w:numbering xmlns:w="%s"><w:abstractNum w:abstractNumId="0">%s'
                     '</w:abstractNum>'
                     '<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>'
                     '<w:num w:numId="2"><w:abstractNumId w:val="0"/></w:num>'
                     '</w:numbering>' % (self.W, lvls))
        styles = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                  '<w:styles xmlns:w="%s">'
                  '<w:style w:type="paragraph" w:default="1" w:styleId="Normal">'
                  '<w:name w:val="Normal"/></w:style>'
                  + "".join('<w:style w:type="paragraph" w:styleId="Heading%d">'
                            '<w:name w:val="heading %d"/><w:pPr>'
                            '<w:outlineLvl w:val="%d"/></w:pPr></w:style>'
                            % (i, i, i - 1) for i in (1, 2, 3))
                  + '</w:styles>') % self.W

        def para(style, text, ilvl, numid):
            return ('<w:p><w:pPr><w:pStyle w:val="%s"/><w:numPr>'
                    '<w:ilvl w:val="%d"/><w:numId w:val="%d"/></w:numPr></w:pPr>'
                    '<w:r><w:rPr><w:rFonts w:eastAsia="宋体"/><w:sz w:val="32"/></w:rPr>'
                    '<w:t xml:space="preserve">%s</w:t></w:r></w:p>'
                    % (style, ilvl, numid, text))
        # 一级用 numId=1，二/三级用 numId=2 —— 同一 abstractNum
        doc = helpers.document_xml(
            para("Heading1", "绪论", 0, 1) + para("Heading2", "研究方法", 1, 2)
            + para("Heading3", "数据来源", 2, 2))
        src = os.path.join(self.tmp, "shared.docx")
        with zipfile.ZipFile(src, "w", zipfile.ZIP_DEFLATED) as z:
            z.writestr("[Content_Types].xml", _CT_LO)
            z.writestr("_rels/.rels", helpers.ROOT_RELS)
            z.writestr("word/_rels/document.xml.rels", _DRELS_LO)
            z.writestr("word/document.xml", doc)
            z.writestr("word/styles.xml", styles)
            z.writestr("word/numbering.xml", numbering)
        wd = self._run_src(src, "wbshared")

        paras = list(self._doc(wd).iter(self.qn("w:p")))
        nids = [p.find(self.qn("w:pPr") + "/" + self.qn("w:numPr") + "/"
                       + self.qn("w:numId")).get(self.qn("w:val")) for p in paras]
        self.assertEqual(len(set(nids)), 1,
                         "共享 abstractNum 的标题被拆到了不同 numId（计数器会分裂）")
        ilvls = [p.find(self.qn("w:pPr") + "/" + self.qn("w:numPr") + "/"
                        + self.qn("w:ilvl")).get(self.qn("w:val")) for p in paras]
        self.assertEqual(ilvls, ["0", "1", "2"], "级别没保持原样")
        num_root = etree.parse(os.path.join(wd, "out_pkg", "word",
                                            "numbering.xml")).getroot()
        unnamed = [a for a in num_root.findall(self.qn("w:abstractNum"))
                   if a.find(self.qn("w:name")) is None]
        self.assertEqual(len(unnamed), 2, "原 abstractNum 应只被克隆一次")

    def test_caption_number_not_bold_when_text_is_not(self):
        """自动编号的渲染取自**段落标记的 rPr**：原表标题整体加粗时，若不清掉段落标记
        上的加粗，就会出现"编号粗、标题文字不粗"（用户实测）。段落标记不是作者内容，
        指派 canonical 样式时一律清掉它的加粗，编号才跟随样式。"""
        import zipfile
        cap = ('<w:p><w:pPr><w:pStyle w:val="表标题"/><w:rPr><w:b/></w:rPr></w:pPr>'
               '<w:r><w:rPr><w:rFonts w:eastAsia="宋体"/><w:b/><w:sz w:val="32"/></w:rPr>'
               '<w:t xml:space="preserve">表1 主要指标</w:t></w:r></w:p>')
        src = os.path.join(self.tmp, "boldcap.docx")
        with zipfile.ZipFile(src, "w", zipfile.ZIP_DEFLATED) as z:
            z.writestr("[Content_Types].xml", _CT_LO)
            z.writestr("_rels/.rels", helpers.ROOT_RELS)
            z.writestr("word/_rels/document.xml.rels", _DRELS_LO)
            z.writestr("word/document.xml", helpers.document_xml(cap))
            z.writestr("word/styles.xml", self._SPLIT_BOLD_STYLES)
            z.writestr("word/numbering.xml", '<w:numbering xmlns:w="%s"/>' % self.W)
        wd = self._run_src(src, "wbboldcap")
        p = list(self._doc(wd).iter(self.qn("w:p")))[0]
        mark = p.find(self.qn("w:pPr") + "/" + self.qn("w:rPr"))
        if mark is not None:
            self.assertIsNone(mark.find(self.qn("w:b")), "段落标记仍加粗→编号会粗")

    def test_heading_and_caption_ends_are_stripped(self):
        # 首尾空格会顶偏居中/缩进，图表标题还会把自动编号顶开；行内空格保留。
        body = (helpers.para("  一、绪论  ", east_asia="宋体", size_hp=32, outline=0)
                + helpers.para("  图1 系统架构  ", east_asia="宋体", size_hp=32,
                               style="图标题"))
        wd, _ = self._run(body)
        texts = ["".join(t.text or "" for t in p.iter(self.qn("w:t")))
                 for p in self._doc(wd).iter(self.qn("w:p"))]
        self.assertEqual(texts[0], "一、绪论")
        self.assertEqual(texts[1], "系统架构")

    def test_blank_lines_outside_cover_get_body_style(self):
        """新增规范：封面外的空行统一套正文样式（仿宋三号）。

        不这么做的话空行跟随 Normal——而 Normal 已被钉成五号（文档网格的前提），
        空行会莫名变矮、与正文行距不一致。封面空行属版式留白，不在规则内。"""
        body = (helpers.para("先进项目2024年度自评价报告", east_asia="宋体",
                             size_hp=44, jc="center")
                + '<w:p/>'                       # 封面空行：不动
                + helpers.para("一、绪论", east_asia="宋体", size_hp=32, outline=0)
                + '<w:p/>'                       # 正文区空行：套正文样式
                + helpers.para("正文内容。", east_asia="宋体", size_hp=32))
        wd, _ = self._run(body)
        paras = list(self._doc(wd).iter(self.qn("w:p")))
        cover_blank = paras[1].find(self.qn("w:pPr") + "/" + self.qn("w:pStyle"))
        self.assertIsNone(cover_blank, "封面空行不该被指派样式")
        body_blank = paras[3].find(self.qn("w:pPr") + "/" + self.qn("w:pStyle"))
        self.assertIsNotNone(body_blank, "正文区空行应套正文样式")
        self.assertEqual(body_blank.get(self.qn("w:val")), "FGWCanonBody")

    _IMAGE_P = ('<w:p><w:r><w:drawing><wp:inline xmlns:wp="http://schemas.openxmlformats'
                '.org/drawingml/2006/wordprocessingDrawing"><wp:extent cx="5000000" '
                'cy="3000000"/></wp:inline></w:drawing></w:r></w:p>')
    _TEXTBOX_P = ('<w:p><w:r><w:pict><v:shape xmlns:v="urn:schemas-microsoft-com:vml">'
                  '<v:textbox><w:txbxContent><w:p><w:r><w:rPr>'
                  '<w:rFonts w:eastAsia="黑体"/><w:sz w:val="21"/></w:rPr>'
                  '<w:t>文本框文字</w:t></w:r></w:p></w:txbxContent></v:textbox>'
                  '</v:shape></w:pict></w:r></w:p>')

    def test_picture_only_paragraph_keeps_no_style(self):
        """图片段落**不指派** canonical 样式。

        所有正文类 canonical 样式都带**固定行距 28 磅**（`lineRule=exact`），而固定行距
        不随内容长高——套到图片段落上，图会被裁成一行高，只露出底部一条（用户实测：
        "图片隐于文字下方，只有底部一行的宽度露出来"）。图片段落也没有字体字号可言。
        只装了文本框的段落同理。"""
        body = (helpers.para("一、绪论", east_asia="宋体", size_hp=32, outline=0)
                + self._IMAGE_P + self._TEXTBOX_P
                + helpers.para("正文内容。", east_asia="宋体", size_hp=32))
        wd, _ = self._run(body)
        paras = list(self._doc(wd).iter(self.qn("w:p")))
        img = [p for p in paras if p.find(".//" + self.qn("w:drawing")) is not None][0]
        self.assertIsNone(img.find(self.qn("w:pPr") + "/" + self.qn("w:pStyle")),
                          "图片段落被指派了样式→固定行距会把图裁成一行")
        self.assertIsNone(img.find(self.qn("w:pPr") + "/" + self.qn("w:spacing")))
        tb = [p for p in paras if p.find(".//" + self.qn("w:pict")) is not None][0]
        self.assertIsNone(tb.find(self.qn("w:pPr") + "/" + self.qn("w:pStyle")))

    def test_hint_only_rfonts_is_not_a_leak(self):
        """`<w:rFonts w:hint="eastAsia"/>` 在中文 Word 里遍地都是，它**不覆盖任何字体**
        （只说明歧义字符按中文还是西文取字体），清理后留着它完全正常。坍缩不变量按
        "rFonts 元素在不在"判会把成百上千个合法段落报成泄漏（用户实测 295 段），
        必须按**属性**判。"""
        import zipfile
        p = ('<w:p><w:r><w:rPr><w:rFonts w:hint="eastAsia" w:eastAsia="宋体"/>'
             '<w:sz w:val="32"/></w:rPr><w:t xml:space="preserve">正文内容。</w:t></w:r></w:p>')
        src = os.path.join(self.tmp, "hint.docx")
        with zipfile.ZipFile(src, "w", zipfile.ZIP_DEFLATED) as z:
            z.writestr("[Content_Types].xml", _CT_LO)
            z.writestr("_rels/.rels", helpers.ROOT_RELS)
            z.writestr("word/_rels/document.xml.rels", _DRELS_LO)
            z.writestr("word/document.xml", helpers.document_xml(
                helpers.para("一、绪论", east_asia="宋体", size_hp=32, outline=0) + p))
            z.writestr("word/styles.xml", self._SPLIT_BOLD_STYLES)
            z.writestr("word/numbering.xml", '<w:numbering xmlns:w="%s"/>' % self.W)
        wd = self._run_src(src, "wbhint")
        validated = run(self._script("45_validate_output.py"), wd)
        self.assertTrue(validated["ok"], validated.get("errors"))
        # hint 保留下来了（我们没把它当字体覆盖去删）
        body_p = list(self._doc(wd).iter(self.qn("w:p")))[1]
        rf = body_p.find(".//" + self.qn("w:rFonts"))
        self.assertIsNotNone(rf)
        self.assertEqual(rf.get(self.qn("w:hint")), "eastAsia")
        self.assertIsNone(rf.get(self.qn("w:eastAsia")), "真正的字体覆盖应已清掉")

    def test_footer_before_header_reference_is_not_a_violation(self):
        # sectPr 里 header/footer 引用是可重复的 choice，任意交错都合法；
        # 写成先后关系会把大量真实文档误判成乱序。
        sys.path.insert(0, os.path.join(helpers.SCRIPTS, "lib"))
        from docxcommon import order_violations
        from lxml import etree
        sect = etree.fromstring(
            ('<w:sectPr xmlns:w="%s">'
             '<w:footerReference w:type="default"/><w:headerReference w:type="default"/>'
             '<w:pgSz w:w="11906" w:h="16838"/></w:sectPr>' % self.W).encode("utf-8"))
        self.assertEqual(order_violations(sect), [])

    def test_cloned_numbering_restarts_per_level(self):
        """标题要**逐级重新编号**（二级在其一级下从头数）。模板里常见的
        `<w:lvlRestart w:val="0"/>` 会关掉这个默认行为，导致二/三/四级变成全文连续
        编号（用户实测）。克隆时去掉它，恢复层级归零；原 abstractNum 不动。"""
        import zipfile
        from lxml import etree
        lvls = "".join(
            '<w:lvl w:ilvl="%d"><w:start w:val="1"/><w:numFmt w:val="decimal"/>'
            '<w:lvlRestart w:val="0"/><w:lvlText w:val="%%%d."/><w:lvlJc w:val="left"/>'
            '<w:pPr><w:ind w:left="420" w:hanging="420"/></w:pPr></w:lvl>' % (i, i + 1)
            for i in range(3))
        numbering = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                     '<w:numbering xmlns:w="%s"><w:abstractNum w:abstractNumId="0">%s'
                     '</w:abstractNum>'
                     '<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>'
                     '</w:numbering>' % (self.W, lvls))
        styles = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                  '<w:styles xmlns:w="%s">'
                  '<w:style w:type="paragraph" w:default="1" w:styleId="Normal">'
                  '<w:name w:val="Normal"/></w:style>'
                  '<w:style w:type="paragraph" w:styleId="Heading1">'
                  '<w:name w:val="heading 1"/><w:pPr><w:outlineLvl w:val="0"/></w:pPr>'
                  '</w:style>'
                  '<w:style w:type="paragraph" w:styleId="Heading2">'
                  '<w:name w:val="heading 2"/><w:pPr><w:outlineLvl w:val="1"/></w:pPr>'
                  '</w:style></w:styles>') % self.W
        doc = helpers.document_xml(
            '<w:p><w:pPr><w:pStyle w:val="Heading1"/><w:numPr><w:ilvl w:val="0"/>'
            '<w:numId w:val="1"/></w:numPr></w:pPr><w:r><w:rPr>'
            '<w:rFonts w:eastAsia="宋体"/><w:sz w:val="32"/></w:rPr>'
            '<w:t>绪论</w:t></w:r></w:p>'
            '<w:p><w:pPr><w:pStyle w:val="Heading2"/><w:numPr><w:ilvl w:val="1"/>'
            '<w:numId w:val="1"/></w:numPr></w:pPr><w:r><w:rPr>'
            '<w:rFonts w:eastAsia="宋体"/><w:sz w:val="32"/></w:rPr>'
            '<w:t>研究方法</w:t></w:r></w:p>')
        src = os.path.join(self.tmp, "restart.docx")
        with zipfile.ZipFile(src, "w", zipfile.ZIP_DEFLATED) as z:
            z.writestr("[Content_Types].xml", _CT_LO)
            z.writestr("_rels/.rels", helpers.ROOT_RELS)
            z.writestr("word/_rels/document.xml.rels", _DRELS_LO)
            z.writestr("word/document.xml", doc)
            z.writestr("word/styles.xml", styles)
            z.writestr("word/numbering.xml", numbering)
        wd = self._run_src(src, "wbrestart")
        num_root = etree.parse(os.path.join(wd, "out_pkg", "word",
                                            "numbering.xml")).getroot()
        by_id = {a.get(self.qn("w:abstractNumId")): a
                 for a in num_root.findall(self.qn("w:abstractNum"))}
        self.assertEqual(len(by_id["0"].findall(".//" + self.qn("w:lvlRestart"))), 3,
                         "原 abstractNum 被原地改了")
        clones = [a for aid, a in by_id.items()
                  if aid != "0" and a.find(self.qn("w:name")) is None]
        self.assertTrue(clones, "没有生成克隆")
        for c in clones:
            self.assertEqual(c.findall(".//" + self.qn("w:lvlRestart")), [],
                             "克隆里仍有 lvlRestart=0 → 编号会全局连续、不逐级归零")

    def test_injected_style_name_collision_is_avoided(self):
        """Word 要求样式名唯一，重名会让它提示"发现无法读取的内容"。文档里已存在同名
        样式时，改我们自己的名字，绝不动文档原有的样式。"""
        import zipfile
        styles = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                  '<w:styles xmlns:w="%s">'
                  '<w:style w:type="paragraph" w:default="1" w:styleId="Normal">'
                  '<w:name w:val="Normal"/></w:style>'
                  '<w:style w:type="paragraph" w:styleId="OldBody">'
                  '<w:name w:val="FGW正文"/></w:style></w:styles>') % self.W
        src = os.path.join(self.tmp, "dupname.docx")
        with zipfile.ZipFile(src, "w", zipfile.ZIP_DEFLATED) as z:
            z.writestr("[Content_Types].xml", _CT_LO)
            z.writestr("_rels/.rels", helpers.ROOT_RELS)
            z.writestr("word/_rels/document.xml.rels", _DRELS_LO)
            z.writestr("word/document.xml", helpers.document_xml(
                helpers.para("正文内容。", east_asia="宋体", size_hp=32)))
            z.writestr("word/styles.xml", styles)
            z.writestr("word/numbering.xml", '<w:numbering xmlns:w="%s"/>' % self.W)
        wd = self._run_src(src, "wbdup")
        names = {}
        for st in self._styles(wd).findall(self.qn("w:style")):
            nm = st.find(self.qn("w:name"))
            if nm is not None:
                names.setdefault(nm.get(self.qn("w:val")), []).append(
                    st.get(self.qn("w:styleId")))
        dup = {n: ids for n, ids in names.items() if len(ids) > 1}
        self.assertEqual(dup, {}, "样式重名：Word 会判定文档损坏")
        self.assertEqual(names["FGW正文"], ["OldBody"], "不该改动文档原有样式")

    def test_style_inherited_numbering_survives_assignment(self):
        """改指 canonical 样式会把**原样式携带的 numPr** 一起弄丢——编号"一、"消失
        等于改了原文。指派时必须把有效编号钉成段落的直接 numPr 保号。"""
        import zipfile
        styles = (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<w:styles xmlns:w="%s">'
            '<w:style w:type="paragraph" w:default="1" w:styleId="Normal">'
            '<w:name w:val="Normal"/></w:style>'
            '<w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/>'
            '<w:pPr><w:outlineLvl w:val="0"/>'
            # 编号挂在【样式】上，段落自身没有 numPr
            '<w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr>'
            '</w:pPr></w:style></w:styles>' % self.W)
        doc = helpers.document_xml(
            '<w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr>'
            '<w:r><w:rPr><w:rFonts w:eastAsia="宋体"/><w:sz w:val="32"/></w:rPr>'
            '<w:t xml:space="preserve">绪论</w:t></w:r></w:p>'
            '<w:p><w:r><w:rPr><w:rFonts w:eastAsia="宋体"/><w:sz w:val="32"/></w:rPr>'
            '<w:t xml:space="preserve">正文。</w:t></w:r></w:p>')
        src = os.path.join(self.tmp, "stylenum.docx")
        with zipfile.ZipFile(src, "w", zipfile.ZIP_DEFLATED) as z:
            z.writestr("[Content_Types].xml", _CT_LO)
            z.writestr("_rels/.rels", helpers.ROOT_RELS)
            z.writestr("word/_rels/document.xml.rels", _DRELS_LO)
            z.writestr("word/document.xml", doc)
            z.writestr("word/styles.xml", styles)
            z.writestr("word/numbering.xml", _NUM_LO)
        wd = run(self._script("05_new_workdir.py"), os.path.join(self.tmp, "wb3"))["workdir"]
        run(self._script("10_prepare_input.py"), src, wd)
        run(self._script("20_extract_structure.py"), wd)
        run(self._script("30_check_format.py"), wd)
        run(self._script("40_apply_fixes.py"), wd)
        p = list(self._doc(wd).iter(self.qn("w:p")))[0]
        self.assertEqual(p.find(self.qn("w:pPr") + "/" + self.qn("w:pStyle")).get(
            self.qn("w:val")), "FGWCanonH1")
        numid = p.find(self.qn("w:pPr") + "/" + self.qn("w:numPr") + "/" + self.qn("w:numId"))
        self.assertIsNotNone(numid, "指派样式后编号丢了（自动编号消失＝改了原文）")


if __name__ == "__main__":
    unittest.main()
