# -*- coding: utf-8 -*-
"""
Stage 2b — export a compact, FULL-DOCUMENT heading/caption candidate list for
MODEL review.

Why this step exists: heading LEVEL and caption grouping are structural
judgments that a shape/style regex cannot always get right — e.g. a document
may use bare "1." for its actual level-1 sections (which mechanically looks
like the level-3 shape "digit + dot"), or a chapter-based caption scheme the
extraction stage guesses wrong. Regex/style matching can also miss a heading
ENTIRELY — a heading can have no recognizable leading number AND no
heading-ish style name, which no amount of pattern tuning can catch from
text shape alone. This is genuinely a full-document decision: you cannot
tell from one paragraph in isolation which level "1." means here, only from
seeing how ALL the headings in the document relate to each other. That is
why this step is NEVER sharded — read the whole list, in document order,
before deciding anything (see SKILL.md).

Usage:
    python 26_export_review.py <workdir>

Reads  <workdir>/structure.json, spec/format_spec.json
Writes <workdir>/review_candidates.json — three full-document-order lists:
    "headings": [{"i", "level" (current guess), "text", "style_id", "outline"}]
        Already detected as a heading by outline/style/number-shape; the
        level shown is only the pre-review guess (see 20_extract_structure.py).
    "captions": [{"i", "kind", "num_raw", "text"}]
        Already detected as a numbered 图/表 caption.
    "possible_missed_headings": [{"i", "text", "font", "suggested_level"}]
        NOT detected as a heading by regex/style/outline at all, but its
        dominant font is one the spec reserves for headings only (黑体 for
        level 1, 楷体 for level 2 — body text is always 仿宋, so either font
        showing up on an unlabeled line is a strong non-regex signal that a
        heading's number went missing, or was never added). This list is
        exactly the "don't rely on regex alone" net: it's how a heading with
        no number and no heading-ish style still gets a chance to be found.
Prints a compact summary (counts only) plus the review file path.

After reading review_candidates.json, if every entry's classification
already looks right and possible_missed_headings has nothing worth
promoting, no further action is needed — 31_check_global.py will use the
auto-detected levels as-is. Only write an overrides file (see
27_apply_review.py) for the entries that need correcting or promoting.
"""
import json
import os
import sys

SPEC_PATH = os.path.join(os.path.dirname(__file__), "..", "spec", "format_spec.json")
MAX_CANDIDATE_LEN = 40  # a heading-length line; mirrors the "short" heuristic
                        # infer_heading_level() already uses for pattern-based guesses


def _font_matches(actual, spec_entry):
    if not actual or not spec_entry:
        return False
    for m in spec_entry.get("east_asia_match", [spec_entry.get("east_asia")]):
        if m and m in actual:
            return True
    return False


def _find_possible_missed_headings(records, spec):
    h1 = spec.get("headings", {}).get("1")
    h2 = spec.get("headings", {}).get("2")
    out = []
    for r in records:
        # Deliberately NOT restricted to region=="body": the cover/body
        # boundary is itself computed from which paragraph is the first
        # DETECTED heading, so a heading regex/style missed entirely can
        # land on the wrong side of that boundary and get excluded right
        # here if we required region=="body" -- exactly the chicken-and-egg
        # case this list exists to break. TOC entries are still skipped
        # (never candidates), and cover-page field labels that happen to use
        # a heading font are an acceptable source of noise here: the model
        # reviews this list, it doesn't get auto-applied.
        if r.get("region") == "toc" or r.get("is_blank"):
            continue
        if r.get("is_heading") or r.get("caption"):
            continue  # already accounted for in the other two lists
        if r.get("text_len", 0) > MAX_CANDIDATE_LEN:
            continue
        ea = r.get("eff", {}).get("east_asia")
        if _font_matches(ea, h1):
            suggested = 1
        elif _font_matches(ea, h2):
            suggested = 2
        else:
            continue
        out.append({"i": r["i"], "text": r["text"], "font": ea, "suggested_level": suggested})
    return out


def main():
    workdir = sys.argv[1]
    with open(os.path.join(workdir, "structure.json"), encoding="utf-8") as f:
        structure = json.load(f)
    records = structure["records"]
    with open(SPEC_PATH, encoding="utf-8") as f:
        spec = json.load(f)

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
    possible_missed = _find_possible_missed_headings(records, spec)

    out_path = os.path.join(workdir, "review_candidates.json")
    with open(out_path, "w", encoding="utf-8") as f:
        json.dump({"headings": headings, "captions": captions,
                  "possible_missed_headings": possible_missed}, f,
                  ensure_ascii=False, indent=1)

    print(json.dumps({
        "status": "ok",
        "n_headings": len(headings),
        "n_captions": len(captions),
        "n_possible_missed_headings": len(possible_missed),
        "review_candidates": out_path,
        "next_step": ("Read review_candidates.json in full (it is small — "
                       "headings/captions/possible_missed_headings only, not "
                       "the whole document). For 'headings'/'captions', judge "
                       "each entry's level/kind from CONTEXT (numbering "
                       "sequence, nesting, surrounding headings). For "
                       "'possible_missed_headings', decide whether each really "
                       "is a heading that lost its number (promote it) or just "
                       "emphasized body text (leave it alone). Write any "
                       "corrections to overrides.json and run "
                       "27_apply_review.py before checking format; if nothing "
                       "needs correcting, proceed straight to step 3."),
    }, ensure_ascii=False))


if __name__ == "__main__":
    main()
