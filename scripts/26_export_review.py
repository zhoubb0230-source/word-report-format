# -*- coding: utf-8 -*-
"""
Stage 2b — export a compact, FULL-DOCUMENT heading/caption candidate list for
MODEL review.

Why this step exists: heading LEVEL and caption grouping are structural
judgments that a shape/style regex cannot always get right — e.g. a document
may use bare "1." for its actual level-1 sections (which mechanically looks
like the level-3 shape "digit + dot"), or a chapter-based caption scheme the
extraction stage guesses wrong. This is genuinely a full-document decision:
you cannot tell from one paragraph in isolation which level "1." means here,
only from seeing how ALL the headings in the document relate to each other.
That is why this step is NEVER sharded — read the whole list, in document
order, before deciding anything (see SKILL.md).

Usage:
    python 26_export_review.py <workdir>

Reads  <workdir>/structure.json
Writes <workdir>/review_candidates.json — two full-document-order lists:
    "headings": [{"i", "level" (current guess), "text", "style_id", "outline"}]
    "captions": [{"i", "kind", "num_raw", "text"}]
Prints a compact summary (counts only) plus the review file path.

After reading review_candidates.json, if every entry's classification
already looks right, no further action is needed — 31_check_global.py will
use the auto-detected levels as-is. Only write an overrides file (see
27_apply_review.py) for the entries that need correcting.
"""
import json
import os
import sys


def main():
    workdir = sys.argv[1]
    with open(os.path.join(workdir, "structure.json"), encoding="utf-8") as f:
        structure = json.load(f)
    records = structure["records"]

    headings = [
        {"i": r["i"], "level": r["level"], "text": r["text"],
         "style_id": r.get("style_id"), "outline": r["eff"].get("outline")}
        for r in records if r.get("region") == "body" and r.get("is_heading")
    ]
    captions = [
        {"i": r["i"], "kind": r["caption"]["kind"],
         "num_raw": r["caption"]["num_raw"], "text": r["text"]}
        for r in records if r.get("region") == "body" and r.get("caption")
    ]

    out_path = os.path.join(workdir, "review_candidates.json")
    with open(out_path, "w", encoding="utf-8") as f:
        json.dump({"headings": headings, "captions": captions}, f,
                  ensure_ascii=False, indent=1)

    print(json.dumps({
        "status": "ok",
        "n_headings": len(headings),
        "n_captions": len(captions),
        "review_candidates": out_path,
        "next_step": ("Read review_candidates.json in full (it is small — "
                       "headings/captions only, not the whole document), "
                       "judge each entry's level/kind from CONTEXT (numbering "
                       "sequence, nesting, surrounding headings), then if any "
                       "need correcting write an overrides.json and run "
                       "27_apply_review.py before checking format."),
    }, ensure_ascii=False))


if __name__ == "__main__":
    main()
