# -*- coding: utf-8 -*-
"""Unit tests for the .doc converter backend facade (docconv / msword).

Locks the behaviour: soffice is preferred, Microsoft Word (Windows COM) is the
fallback, and can_convert_doc()/convert() reflect whichever backends exist.
These are pure-logic tests — no real LibreOffice / Word is invoked; the backend
availability probes are monkeypatched.
"""
import os
import sys
import unittest

sys.path.insert(0, os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))),
                                "scripts", "lib"))
import docconv
import msword
import soffice


class DocConvBackendTest(unittest.TestCase):
    def setUp(self):
        self._sof = soffice.soffice_available
        self._msw = msword.msword_available
        self._sof_conv = soffice.convert
        self._msw_conv = msword.convert

    def tearDown(self):
        soffice.soffice_available = self._sof
        msword.msword_available = self._msw
        soffice.convert = self._sof_conv
        msword.convert = self._msw_conv

    def _set(self, sof, msw):
        soffice.soffice_available = lambda: sof
        msword.msword_available = lambda: msw

    def test_can_convert_with_either_backend(self):
        self._set(True, False)
        self.assertTrue(docconv.can_convert_doc())
        self._set(False, True)   # e.g. Windows + Office, no LibreOffice
        self.assertTrue(docconv.can_convert_doc())
        self._set(True, True)
        self.assertTrue(docconv.can_convert_doc())

    def test_cannot_convert_when_no_backend(self):
        self._set(False, False)
        self.assertFalse(docconv.can_convert_doc())
        with self.assertRaises(RuntimeError):
            docconv.convert("x.doc", "docx", "/tmp/out")

    def test_prefers_soffice_when_both_present(self):
        self._set(True, True)
        called = {}
        soffice.convert = lambda *a, **k: called.setdefault("who", "soffice") or "/o.docx"
        msword.convert = lambda *a, **k: called.setdefault("who", "msword") or "/o.docx"
        docconv.convert("x.doc", "docx", "/tmp/out")
        self.assertEqual(called["who"], "soffice")

    def test_falls_back_to_msword(self):
        self._set(False, True)
        called = {}
        soffice.convert = lambda *a, **k: called.setdefault("who", "soffice") or "/o.docx"
        msword.convert = lambda *a, **k: called.setdefault("who", "msword") or "/o.docx"
        docconv.convert("x.doc", "docx", "/tmp/out")
        self.assertEqual(called["who"], "msword")

    def test_msword_unavailable_off_windows(self):
        # On this (non-Windows) test host, Word must never be reported available.
        if os.name != "nt":
            self.assertFalse(msword.msword_available())


if __name__ == "__main__":
    unittest.main()
