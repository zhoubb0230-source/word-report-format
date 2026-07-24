# -*- coding: utf-8 -*-
"""End-to-end smoke test: run the real pipeline scripts (05→10→20→30→40→45) on a
minimal generated .docx and assert the output validates and preserves paragraph
count. Exercises the refactored shared helpers (tag_regions / write_shards /
in_textbox) through their actual entry points.

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

    def test_full_path_a(self):
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

        # shard_size must be persisted for 27 to reuse (P0-2 change)
        with open(os.path.join(wd, "structure.json"), encoding="utf-8") as f:
            structure = json.load(f)
        self.assertIn("shard_size", structure)

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

    def test_apply_review_reshards_with_persisted_size(self):
        # A tiny doc still round-trips through 27 with an empty overrides file,
        # confirming the shared write_shards + persisted shard_size work.
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
        self.assertTrue(os.path.isdir(os.path.join(wd, "shards")))


if __name__ == "__main__":
    unittest.main()
