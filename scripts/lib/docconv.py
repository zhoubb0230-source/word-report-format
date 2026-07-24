# -*- coding: utf-8 -*-
"""
docconv.py — unified .doc <-> .docx converter facade.

Wraps the two backends so callers stay backend-agnostic:
  * soffice  — LibreOffice, works on any OS (sandbox/Linux/mac/Windows).
  * msword   — Microsoft Word via COM, Windows-only fallback (common on the
               user's own Windows machine where LibreOffice is not installed).

Preference order is soffice first (headless, no Office UI, works in the
sandbox), then msword. If neither is present, ``can_convert_doc()`` is False and
the pipeline must stop and notify the user rather than guess a format.
"""

import soffice
import msword


def backends():
    """Report which converters are available. Handy for env probing."""
    return {
        "soffice": soffice.soffice_available(),
        "msword": msword.msword_available(),
    }


def can_convert_doc():
    """True if ANY backend can convert .doc <-> .docx."""
    return soffice.soffice_available() or msword.msword_available()


def convert(input_path, out_format, out_dir):
    """
    Convert to ``out_format`` ('docx'/'doc') into ``out_dir`` using the first
    available backend. Returns the output path, or raises if none available.
    """
    if soffice.soffice_available():
        return soffice.convert(input_path, out_format, out_dir)
    if msword.msword_available():
        return msword.convert(input_path, out_format, out_dir)
    raise RuntimeError(
        "no .doc converter available (neither LibreOffice nor Microsoft Word)"
    )


if __name__ == "__main__":
    import json
    print(json.dumps({"backends": backends(), "can_convert_doc": can_convert_doc()},
                     ensure_ascii=False))
