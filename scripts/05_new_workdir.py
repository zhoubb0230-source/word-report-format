# -*- coding: utf-8 -*-
"""Stage 0.5 — create a fresh, uniquely-named working directory for this run.

Usage:
    python 05_new_workdir.py [base_dir]

Why this exists:
  * Keeps every intermediate artifact under ONE fixed temp base
    (<cwd>/.word_report_work by default) instead of scattering workdirs
    wherever the model happens to pick — solves "输出过程件散乱".
  * Uses tempfile.mkdtemp, which creates the directory ATOMICALLY with a name
    no other run holds, so several sessions running at once never share a
    workdir — solves "开多个 session 共用一个输出目录导致内容混乱".
  * Pairs with 59_cleanup.py, which removes the returned workdir once the final
    output has been finalized out of it — solves "处理完成后需要删除临时目录".

Pass an explicit base_dir to override the default (env WORD_REPORT_WORK_BASE
does the same). Prints ONE JSON line: {status, workdir, base}.
"""
import json
import os
import sys
import tempfile
from datetime import datetime

sys.path.insert(0, os.path.join(os.path.dirname(__file__), "lib"))
import workdir as wd


def main():
    if len(sys.argv) > 2:
        print(json.dumps({"status": "error",
                          "error": "usage: 05_new_workdir.py [base_dir]"}))
        sys.exit(1)
    base = os.path.abspath(sys.argv[1]) if len(sys.argv) == 2 else wd.base_dir()
    try:
        os.makedirs(base, exist_ok=True)
        prefix = "run_%s_" % datetime.now().strftime("%Y%m%d_%H%M%S")
        path = tempfile.mkdtemp(prefix=prefix, dir=base)
    except OSError as e:
        print(json.dumps({"status": "error", "reason": "mkdir_failed",
                          "message": "无法创建工作目录: %s" % e}, ensure_ascii=False))
        sys.exit(2)

    print(json.dumps({"status": "ok", "workdir": os.path.abspath(path),
                      "base": base}, ensure_ascii=False))


if __name__ == "__main__":
    main()
