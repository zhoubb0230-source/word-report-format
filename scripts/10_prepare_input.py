# -*- coding: utf-8 -*-
"""
Stage 1 — prepare the input document into a working directory.

Usage:
    python 10_prepare_input.py <input.doc|.docx> <workdir>

Behaviour (deterministic; never modifies the original):
  * .docx input  -> copied to <workdir>/working.docx
  * .doc  input  -> if a converter is available (LibreOffice, or Microsoft
                    Word on Windows), converted to <workdir>/working.docx;
                    if NONE available, prints an error JSON with
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
import docconv


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
        if not docconv.can_convert_doc():
            fail("doc_conversion_unavailable",
                 "输入为 .doc 格式，但当前环境既未安装 LibreOffice(soffice)、"
                 "也未检测到可用的 Microsoft Word（Windows），"
                 "无法将 .doc 转换为可处理的 .docx。请告知用户："
                 "在具备 doc 转换能力的环境中重试（Linux/Mac 装 LibreOffice，"
                 "或在装有 Office 的 Windows 上运行），或由用户先将文档另存为 .docx 后提供。")
        try:
            converted = docconv.convert(src, "docx", os.path.join(workdir, "_conv"))
        except Exception as e:
            fail("doc_conversion_failed",
                 "LibreOffice 转换 .doc 失败：%s" % e, 3)
        shutil.copy(converted, working)
        original_ext = "doc"
    else:
        fail("unsupported_type",
             "仅支持 .doc / .docx，当前为: .%s" % ext, 1)

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
