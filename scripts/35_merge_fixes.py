# -*- coding: utf-8 -*-
"""
Stage 3c — merge all <workdir>/fixes_parts/*.json into <workdir>/fixes.json.
Only needed on the SHARDED / sub-agent path. Guarantees a single valid JSON
array (avoids any hand-written JSON from a model).

    python 35_merge_fixes.py <workdir>
"""
import json
import os
import sys


def main():
    if len(sys.argv) < 2:
        print(json.dumps({"status": "error",
                          "error": "usage: 35_merge_fixes.py <workdir>"}))
        sys.exit(1)
    workdir = sys.argv[1]
    parts = os.path.join(workdir, "fixes_parts")
    fixes = []
    if os.path.isdir(parts):
        for name in sorted(os.listdir(parts)):
            if not name.endswith(".json"):
                continue
            with open(os.path.join(parts, name), encoding="utf-8") as f:
                data = json.load(f)
                if isinstance(data, list):
                    fixes.extend(data)
    # de-dup identical (para_index, op, rule_id) keeping first
    seen, deduped = set(), []
    for fx in fixes:
        key = (fx.get("para_index"), fx.get("op"), fx.get("rule_id"))
        if key in seen:
            continue
        seen.add(key)
        deduped.append(fx)

    with open(os.path.join(workdir, "fixes.json"), "w", encoding="utf-8") as f:
        json.dump(deduped, f, ensure_ascii=False, indent=1)

    by_op = {}
    for fx in deduped:
        by_op[fx["op"]] = by_op.get(fx["op"], 0) + 1
    print(json.dumps({"status": "ok", "n_fixes": len(deduped), "by_op": by_op},
                     ensure_ascii=False))


if __name__ == "__main__":
    main()
