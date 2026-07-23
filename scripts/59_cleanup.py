# -*- coding: utf-8 -*-
"""Stage 6 — remove a run's working directory once delivery is finished.

Usage:
    python 59_cleanup.py <workdir>

Run this ONLY after 50_finalize.py has copied the final document out to the
user's output directory (which lives OUTSIDE the workdir) — everything left in
<workdir> is disposable intermediate state (working.docx / structure.json /
shards/ / out_pkg/ / formatted.docx / ...).

Safety: refuses to delete anything that is not inside a ``.word_report_work``
base (see lib/workdir.is_inside_base), so a mistyped or unexpected path can never
turn this into an arbitrary rmtree. Prints ONE JSON line.
"""
import json
import os
import shutil
import sys

sys.path.insert(0, os.path.join(os.path.dirname(__file__), "lib"))
import workdir as wd


def main():
    if len(sys.argv) != 2:
        print(json.dumps({"status": "error",
                          "error": "usage: 59_cleanup.py <workdir>"}))
        sys.exit(1)
    path = sys.argv[1]

    if not os.path.exists(path):
        # Already gone — treat as success so cleanup is idempotent.
        print(json.dumps({"status": "ok", "removed": False, "reason": "not_found",
                          "workdir": os.path.abspath(path)}, ensure_ascii=False))
        return

    if not wd.is_inside_base(path):
        print(json.dumps({
            "status": "error", "reason": "refused_outside_base",
            "message": "拒绝删除：路径不在 %s 临时基目录内，为安全起见不予删除。" % wd.BASE_NAME,
            "workdir": os.path.abspath(path),
        }, ensure_ascii=False))
        sys.exit(2)

    try:
        shutil.rmtree(path)
    except OSError as e:
        print(json.dumps({"status": "error", "reason": "rmtree_failed",
                          "message": "删除工作目录失败: %s" % e,
                          "workdir": os.path.abspath(path)}, ensure_ascii=False))
        sys.exit(2)

    # Best-effort: drop the base dir too if this was the last run in it.
    base = wd.base_dir()
    try:
        if os.path.isdir(base) and not os.listdir(base):
            os.rmdir(base)
    except OSError:
        pass

    print(json.dumps({"status": "ok", "removed": True,
                      "workdir": os.path.abspath(path)}, ensure_ascii=False))


if __name__ == "__main__":
    main()
