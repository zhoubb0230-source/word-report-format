# -*- coding: utf-8 -*-
"""
Stage 0 — probe the runtime environment. Prints a small JSON object.

The orchestrating agent reads this to decide:
  * whether .doc conversion is possible (soffice present)
  * whether required Python libs exist

No arguments. Deterministic. Prints ONE JSON line to stdout.
"""
import json
import sys
import os

sys.path.insert(0, os.path.join(os.path.dirname(__file__), "lib"))

out = {"python": sys.version.split()[0], "lxml": False, "soffice": False, "ok": False}
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

out["can_convert_doc"] = bool(out["soffice"])
out["ok"] = out["lxml"]
print(json.dumps(out, ensure_ascii=False))
