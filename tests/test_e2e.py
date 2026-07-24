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
    """LibreOffice .doc→.docx 把缩进表达成【绝对】twips 的 firstLine/hanging（挂
    在样式或编号层，不发 *Chars 变体）。仅写字符单位 firstLineChars 不能压过一个
    【继承来的绝对】hanging（Word 里 hanging 会赢），于是首行缩进被渲染成 -0.74cm
    的悬挂缩进。40 现在给每个字符单位缩进补一个绝对 firstLine/left 伴随值，使直接
    覆盖能压过继承的绝对缩进。此处锁死这一行为。"""

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

    def test_char_twips_two_chars_sanhao(self):
        # 2 chars (firstLineChars=200) at 三号(32 半点=16pt) = 640 twips.
        self.assertEqual(self.mod._char_twips(200, 32), 640)
        self.assertEqual(self.mod._char_twips(0, 32), 0)

    def test_first_line_gets_absolute_companion(self):
        pPr = self._ppr()
        self.mod._set_first_line_and_clear_left(pPr, 200, False, False, None, size_hp=32)
        ind = self._ind(pPr)
        self.assertEqual(ind.get(self.mod.qn("w:firstLineChars")), "200")
        # absolute companion present and positive -> beats an inherited hanging
        self.assertEqual(ind.get(self.mod.qn("w:firstLine")), "640")

    def test_direct_hanging_removed_when_setting_first_line(self):
        # a paragraph carrying an ABSOLUTE hanging (the LibreOffice shape) must
        # come out with a positive first-line indent and NO hanging.
        pPr = self._ppr('<w:ind w:hanging="420" w:left="420"/>')
        self.mod._set_first_line_and_clear_left(pPr, 200, True, False, None, size_hp=32)
        ind = self._ind(pPr)
        self.assertIsNone(ind.get(self.mod.qn("w:hanging")))
        self.assertEqual(ind.get(self.mod.qn("w:firstLine")), "640")
        self.assertEqual(ind.get(self.mod.qn("w:leftChars")), "0")

    def test_clear_no_indent_writes_absolute_zero(self):
        # title/caption "no indent": firstLineChars=0 must be paired with an
        # absolute firstLine=0 so an inherited absolute hanging is overridden.
        pPr = self._ppr('<w:ind w:hanging="420"/>')
        self.mod._set_first_line_and_clear_left(pPr, 0, True, True, None, size_hp=32)
        ind = self._ind(pPr)
        self.assertEqual(ind.get(self.mod.qn("w:firstLine")), "0")
        self.assertIsNone(ind.get(self.mod.qn("w:hanging")))

    def test_left_chars_gets_absolute_companion(self):
        # TOC per-level left indent (leftChars=200) also gets an absolute left.
        pPr = self._ppr()
        self.mod._set_first_line_and_clear_left(pPr, None, False, False, 200, size_hp=32)
        ind = self._ind(pPr)
        self.assertEqual(ind.get(self.mod.qn("w:leftChars")), "200")
        self.assertEqual(ind.get(self.mod.qn("w:left")), "640")


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
    """整链回归：一个从编号层继承【绝对 hanging】的自动编号标题，跑完 20→30→40
    后，其段落直接 w:ind 必须带上一个【正值绝对 firstLine】（压过继承的悬挂缩进）
    且不残留直接 hanging——即首行缩进而非 -0.74cm 悬挂缩进。"""

    def setUp(self):
        self.tmp = tempfile.mkdtemp(prefix="wrf_lo_")

    def tearDown(self):
        shutil.rmtree(self.tmp, ignore_errors=True)

    def _script(self, name):
        return os.path.join(helpers.SCRIPTS, name)

    def _build(self, path):
        import zipfile
        doc = (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>'
            '<w:p><w:pPr><w:pStyle w:val="Heading1"/>'
            '<w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr></w:pPr>'
            '<w:r><w:rPr><w:rFonts w:eastAsia="宋体"/><w:sz w:val="32"/></w:rPr>'
            '<w:t xml:space="preserve">概述</w:t></w:r></w:p>'
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

    def test_heading_gets_positive_absolute_first_line(self):
        from lxml import etree
        src = self._build(os.path.join(self.tmp, "lo.docx"))
        base = os.path.join(self.tmp, "wb")
        wd = run(self._script("05_new_workdir.py"), base)["workdir"]
        run(self._script("10_prepare_input.py"), src, wd)
        run(self._script("20_extract_structure.py"), wd)
        run(self._script("30_check_format.py"), wd)
        run(self._script("40_apply_fixes.py"), wd)

        w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

        def qn(t):
            return "{%s}%s" % (w, t.split(":")[1])

        root = etree.parse(os.path.join(wd, "out_pkg", "word", "document.xml")).getroot()
        heading_p = root.iter(qn("w:p")).__next__()  # first body paragraph = 概述 heading
        ind = heading_p.find(qn("w:pPr") + "/" + qn("w:ind"))
        self.assertIsNotNone(ind, "标题段落缺少 w:ind")
        # 首行缩进2字符仍写字符单位
        self.assertEqual(ind.get(qn("w:firstLineChars")), "200")
        # 关键：补了正值绝对 firstLine，压过编号层继承的绝对 hanging
        self.assertGreater(int(ind.get(qn("w:firstLine")) or "0"), 0)
        # 段落直接属性上不残留 hanging（否则 Word 里悬挂缩进会赢）
        self.assertIsNone(ind.get(qn("w:hanging")))


if __name__ == "__main__":
    unittest.main()
