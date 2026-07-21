# -*- coding: utf-8 -*-
"""
Stage 1 — prepare the input document into a working directory.

Usage:
    python 10_prepare_input.py <input.doc|.docx> <workdir>

Behaviour (deterministic; never modifies the original):
  * .docx input  -> copied to <workdir>/working.docx
  * .doc  input  -> if soffice available, converted to <workdir>/working.docx;
                    if NOT available, prints an error JSON with
                    status="error", reason="doc_conversion_unavailable" and
                    exits 2 so the agent can NOTIFY THE USER rather than guess.
  * writes <workdir>/meta.json recording original name, original extension
    (so the final output is emitted in the SAME format the user provided).

Prints ONE JSON line to stdout.
"""
import json
import os
import shutil
import sys

sys.path.insert(0, os.path.join(os.path.dirname(__file__), "lib"))
import soffice as sof


def fail(reason, message, code=2):
    print(json.dumps({"status": "error", "reason": reason, "message": message},
                     ensure_ascii=False))
    sys.exit(code)


def main():
    if len(sys.argv) != 3:
        fail("bad_args", "usage: 10_prepare_input.py <input> <workdir>", 1)
    src = sys.argv[1]
    workdir = sys.argv[2]
    if not os.path.isfile(src):
        fail("not_found", "input file not found: %s" % src, 1)
    os.makedirs(workdir, exist_ok=True)

    base = os.path.basename(src)
    stem, ext = os.path.splitext(base)
    ext = ext.lower().lstrip(".")
    working = os.path.join(workdir, "working.docx")

    if ext == "docx":
        shutil.copy(src, working)
        original_ext = "docx"
    elif ext == "doc":
        if not sof.soffice_available():
            fail("doc_conversion_unavailable",
                 "\u8f93\u5165\u4e3a .doc \u683c\u5f0f\uff0c\u4f46\u5f53\u524d\u73af\u5883\u672a\u5b89\u88c5 LibreOffice(soffice)\uff0c"
                 "\u65e0\u6cd5\u5c06 .doc \u8f6c\u6362\u4e3a\u53ef\u5904\u7406\u7684 .docx\u3002\u8bf7\u544a\u77e5\u7528\u6237\uff1a"
                 "\u5728\u5177\u5907 doc \u8f6c\u6362\u80fd\u529b\u7684\u73af\u5883\u4e2d\u91cd\u8bd5\uff0c\u6216\u7531\u7528\u6237\u5148\u5c06\u6587\u6863\u53e6\u5b58\u4e3a .docx \u540e\u63d0\u4f9b\u3002")
        try:
            converted = sof.convert(src, "docx", os.path.join(workdir, "_conv"))
        except Exception as e:
            fail("doc_conversion_failed",
                 "LibreOffice \u8f6c\u6362 .doc \u5931\u8d25\uff1a%s" % e, 3)
        shutil.copy(converted, working)
        original_ext = "doc"
    else:
        fail("unsupported_type",
             "\u4ec5\u652f\u6301 .doc / .docx\uff0c\u5f53\u524d\u4e3a: .%s" % ext, 1)

    meta = {
        "status": "ok",
        "original_path": os.path.abspath(src),
        "original_stem": stem,
        "original_ext": original_ext,
        "working_docx": os.path.abspath(working),
    }
    with open(os.path.join(workdir, "meta.json"), "w", encoding="utf-8") as f:
        json.dump(meta, f, ensure_ascii=False, indent=2)
    print(json.dumps(meta, ensure_ascii=False))


if __name__ == "__main__":
    main()
