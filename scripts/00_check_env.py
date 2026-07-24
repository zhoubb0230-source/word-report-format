# -*- coding: utf-8 -*-
"""
Stage 0 — probe the runtime environment. Prints a small JSON object.

The orchestrating agent reads this to decide:
  * whether .doc conversion is possible (LibreOffice/soffice OR, on Windows,
    Microsoft Word)
  * whether required Python libs exist

No arguments. Deterministic. Prints ONE JSON line to stdout.
"""
import json
import sys
import os

sys.path.insert(0, os.path.join(os.path.dirname(__file__), "lib"))

out = {"python": sys.version.split()[0], "lxml": False,
       "soffice": False, "msword": False, "ok": False}
try:
    import lxml  # noqa
    out["lxml"] = True
except Exception as e:
    out["lxml_error"] = str(e)

try:
    import soffice as _sof
    out["soffice"] = _sof.soffice_available()
except Exception as e:
    out["soffice_error"] = str(e)

try:
    import msword as _msw
    out["msword"] = _msw.msword_available()
except Exception as e:
    out["msword_error"] = str(e)

# .doc conversion works with either backend (soffice anywhere, or MS Word on Windows).
out["can_convert_doc"] = bool(out["soffice"] or out["msword"])
out["ok"] = out["lxml"]
print(json.dumps(out, ensure_ascii=False))
