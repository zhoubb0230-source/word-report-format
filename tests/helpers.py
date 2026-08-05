# -*- coding: utf-8 -*-
"""Shared test helpers: put scripts/lib on sys.path, load a numbered script as a
module, and build a minimal valid .docx on disk. Stdlib + lxml only (matches the
project's zero-extra-dependency rule)."""
import importlib.util
import os
import sys
import zipfile

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SCRIPTS = os.path.join(ROOT, "scripts")
LIB = os.path.join(SCRIPTS, "lib")
SPEC_PATH = os.path.join(ROOT, "spec", "format_spec.json")

if LIB not in sys.path:
    sys.path.insert(0, LIB)


def load_script(filename):
    """Import a numbered pipeline script (e.g. '40_apply_fixes.py') as a module.
    Their names aren't valid identifiers, so a normal import can't reach them."""
    path = os.path.join(SCRIPTS, filename)
    modname = "script_" + filename.replace(".py", "").replace("-", "_")
    spec = importlib.util.spec_from_file_location(modname, path)
    mod = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(mod)
    return mod


CONTENT_TYPES = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
    '<Default Extension="xml" ContentType="application/xml"/>'
    '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
    '</Types>'
)
ROOT_RELS = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>'
    '</Relationships>'
)
DOC_RELS = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>'
)

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"


def para(text, east_asia="宋体", size_hp=32, jc=None, outline=None,
         style=None, first_line_chars=None):
    """One <w:p> XML string with a single run. Defaults are deliberately
    non-compliant (宋体 / no indent) so checks have something to fix."""
    ppr = []
    if style is not None:
        ppr.append('<w:pStyle w:val="%s"/>' % style)
    if outline is not None:
        ppr.append('<w:outlineLvl w:val="%d"/>' % outline)
    ind = []
    if first_line_chars is not None:
        ind.append('w:firstLineChars="%d"' % first_line_chars)
    if ind:
        ppr.append('<w:ind %s/>' % " ".join(ind))
    if jc is not None:
        ppr.append('<w:jc w:val="%s"/>' % jc)
    ppr_xml = ("<w:pPr>%s</w:pPr>" % "".join(ppr)) if ppr else ""
    rpr = '<w:rPr><w:rFonts w:eastAsia="%s"/><w:sz w:val="%d"/></w:rPr>' % (east_asia, size_hp)
    return '<w:p>%s<w:r>%s<w:t xml:space="preserve">%s</w:t></w:r></w:p>' % (ppr_xml, rpr, text)


def document_xml(body_inner, pg_mar=None):
    """Wrap paragraph XML in a full document.xml. pg_mar=(top,bottom,left,right,
    header,footer) adds a sectPr so page-margin checks have input."""
    sect = ""
    if pg_mar:
        t, b, l, r, h, f = pg_mar
        sect = ('<w:sectPr><w:pgMar w:top="%d" w:bottom="%d" w:left="%d" '
                'w:right="%d" w:header="%d" w:footer="%d"/></w:sectPr>'
                % (t, b, l, r, h, f))
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<w:document xmlns:w="%s"><w:body>%s%s</w:body></w:document>'
            % (W_NS, body_inner, sect))


def build_docx(path, body_inner, pg_mar=None):
    """Write a minimal but valid .docx to path. Includes the four parts the
    pipeline needs, including word/_rels/document.xml.rels (40's CommentWriter
    reads it)."""
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", CONTENT_TYPES)
        z.writestr("_rels/.rels", ROOT_RELS)
        z.writestr("word/_rels/document.xml.rels", DOC_RELS)
        z.writestr("word/document.xml", document_xml(body_inner, pg_mar))
    return path
