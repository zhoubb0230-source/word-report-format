# -*- coding: utf-8 -*-
"""Finalize the formatted document with the required output name.

Usage:
    python 50_finalize.py <workdir> <output_dir>

Output name:  {original_stem}_格式化版本_{YYYYMMDD_HHMMSS}.{original_ext}

If the original was a .doc, the formatted .docx is converted back to .doc via
LibreOffice so the output type matches the input. If conversion is required but
unavailable, exits non-zero with a JSON message so the caller can notify the
user (never silently emits the wrong type).
"""
import json
import os
import shutil
import sys
from datetime import datetime

sys.path.insert(0, os.path.join(os.path.dirname(__file__), "lib"))
from soffice import convert, soffice_available


def main():
    if len(sys.argv) < 3:
        print(json.dumps({"status": "error",
                          "error": "usage: 50_finalize.py <workdir> <output_dir>"}))
        sys.exit(1)
    workdir, output_dir = sys.argv[1], sys.argv[2]
    os.makedirs(output_dir, exist_ok=True)

    meta = json.load(open(os.path.join(workdir, "meta.json"), encoding="utf-8"))
    formatted = os.path.join(workdir, "formatted.docx")
    if not os.path.exists(formatted):
        print(json.dumps({"status": "error", "error": "formatted.docx not found; run 40 first"},
                         ensure_ascii=False))
        sys.exit(1)

    stem = meta.get("original_stem", "document")
    ext = meta.get("original_ext", "docx")
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    out_name = "%s_格式化版本_%s.%s" % (stem, ts, ext)
    out_path = os.path.join(output_dir, out_name)

    if ext == "doc":
        if not soffice_available():
            print(json.dumps({
                "status": "error",
                "reason": "doc_conversion_unavailable",
                "message": "原文件为 .doc 格式，但当前环境无法进行 doc 转换，无法输出与原文件一致的 .doc 类型，请知悉。"
            }, ensure_ascii=False))
            sys.exit(2)
        produced = convert(formatted, "doc", output_dir)
        # convert() names by source stem (formatted.doc); rename to required name
        if os.path.abspath(produced) != os.path.abspath(out_path):
            if os.path.exists(out_path):
                os.remove(out_path)
            shutil.move(produced, out_path)
    else:
        shutil.copyfile(formatted, out_path)

    result = {"status": "ok", "output_path": out_path, "output_name": out_name,
              "original_ext": ext}
    json.dump(result, open(os.path.join(workdir, "finalize_report.json"), "w",
                           encoding="utf-8"), ensure_ascii=False, indent=1)
    print(json.dumps(result, ensure_ascii=False))


if __name__ == "__main__":
    main()
