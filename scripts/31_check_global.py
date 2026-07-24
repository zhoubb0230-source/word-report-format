# -*- coding: utf-8 -*-
"""
Stage 3b — GLOBAL checks that cannot be done per-shard: heading/figure/table
continuity, page setup, and document-level hints (cover fields, green cover,
caption content). Only needed on the SHARDED / sub-agent path.

    python 31_check_global.py <workdir>

Reads <workdir>/structure.json, writes <workdir>/fixes_parts/_global.json.
Prints a compact summary.
"""
import json
import os
import sys

HERE = os.path.dirname(__file__)
sys.path.insert(0, os.path.join(HERE, "lib"))
import checks


def main():
    if len(sys.argv) < 2:
        print(json.dumps({"status": "error",
                          "error": "usage: 31_check_global.py <workdir>"}))
        sys.exit(1)
    workdir = sys.argv[1]
    spec = checks.load_default_spec()
    with open(os.path.join(workdir, "structure.json"), encoding="utf-8") as f:
        structure = json.load(f)
    records = structure["records"]

    fixes = []
    ps = checks.check_page_setup(structure.get("page_setup"), spec)
    if ps:
        fixes.append(ps)
    fixes.extend(checks.continuity(records, spec))
    fixes.extend(checks.doc_hints(structure, spec))

    parts = os.path.join(workdir, "fixes_parts")
    os.makedirs(parts, exist_ok=True)
    with open(os.path.join(parts, "_global.json"), "w", encoding="utf-8") as f:
        json.dump(fixes, f, ensure_ascii=False)
    print(json.dumps({"status": "ok", "n_global_fixes": len(fixes)}, ensure_ascii=False))


if __name__ == "__main__":
    main()
