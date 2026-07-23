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

Every heading/caption candidate carries a trust tier (see lib/headings.py):
  - CONFIRMED (level_source/caption source "outline" or "style"): a real
    w:outlineLvl or a heading/caption-ish paragraph style backs the
    classification. checks.py auto-renumbers these on mismatch without
    further confirmation — the model's only job for these is correcting the
    LEVEL/KIND when the shape guess is wrong (e.g. "1." guessed as level 3
    should really be level 1 in this document).
  - UNCONFIRMED ("pattern"): matched only because the text happens to look
    like a numbering shape, with no style/outline behind it. checks.py NEVER
    auto-edits these — only a review hint. They stay hint-only unless this
    review step explicitly confirms them via overrides.json (which stamps
    them "model_confirmed" and makes them auto-fixable from then on).
  - possible_missed_headings: not even pattern-matched — no number, no
    heading-ish style, nothing regex/style can catch — surfaced only via a
    font signal (see below). Also only ever promoted through an explicit
    override.

Usage:
    python 26_export_review.py <workdir>

Reads  <workdir>/structure.json, spec/format_spec.json
Writes <workdir>/review_candidates.json:
    "headings_confirmed":   [{"i", "level", "text", "style_id", "outline"}]
        outline/style-backed; model should correct the LEVEL if the shape
        guess is wrong, or set null if this was misclassified entirely
        (e.g. a caption that slipped through — should already be rare after
        the caption-style/shape guard, but double-check chapter/appendix
        titles etc.). Auto-renumbered once level is confirmed correct.
    "headings_unconfirmed": [{"i", "level" (guess), "text", "style_id"}]
        pattern-only; model should confirm/correct the level to promote it
        (auto-fixable from then on) or leave it alone (stays hint-only,
        never auto-edited).
    "captions_confirmed":   [{"i", "kind", "num_raw", "text"}]
        caption-style-backed (题注/图表标题/...); num_raw null means the
        style is confirmed but no digit was found (missing number).
    "captions_unconfirmed": [{"i", "kind", "num_raw", "text"}]
        matched only by the "图/表 + 数字" text shape, no caption style.
    "possible_missed_headings": [{"i", "text", "font", "suggested_level"}]
        NOT detected as a heading by regex/style/outline at all, but its
        dominant font is one the spec reserves for headings only (黑体 for
        level 1, 楷体 for level 2 — body text is always 仿宋, so either font
        showing up on an unlabeled line is a strong non-regex signal that a
        heading's number went missing, or was never added).
Prints a compact summary (counts only) plus the review file path.

After reading review_candidates.json: correct any wrong levels/kinds in the
*_confirmed lists (auto-applies once fixed). For *_unconfirmed and
possible_missed_headings, decide case by case whether each is really a
heading/caption that needs promoting via overrides.json, or should be left
alone (in which case it just stays a review hint in the output — never
silently modified). If nothing needs correcting or promoting, proceed
straight to step 3; 31_check_global.py will use the current classification
as-is.
"""
import json
import os
import sys

sys.path.insert(0, os.path.join(os.path.dirname(__file__), "lib"))
from checks import cover_role

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
            continue  # already accounted for in the other lists
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

    headings_confirmed, headings_unconfirmed = [], []
    for r in records:
        if r.get("region") != "body" or not r.get("is_heading"):
            continue
        entry = {"i": r["i"], "level": r["level"], "text": r["text"],
                 "style_id": r.get("style_id"), "outline": r["eff"].get("outline")}
        if r.get("level_source") == "pattern":
            headings_unconfirmed.append(entry)
        else:
            headings_confirmed.append(entry)

    captions_confirmed, captions_unconfirmed = [], []
    for r in records:
        if r.get("region") != "body" or not r.get("caption"):
            continue
        cap = r["caption"]
        entry = {"i": r["i"], "kind": cap["kind"], "num_raw": cap["num_raw"], "text": r["text"]}
        if cap.get("source") == "pattern":
            captions_unconfirmed.append(entry)
        else:
            captions_confirmed.append(entry)

    possible_missed = _find_possible_missed_headings(records, spec)

    # Cover paragraphs (whole cover, verbatim -- it is small): the model reads
    # these and assigns each a role in overrides.json["cover"] when the
    # heuristic guess is wrong. Covers have no fixed layout, so this is a
    # genuine full-page judgment (title may wrap; the 密级/文本编号 line may have
    # no key). guess_role is checks.cover_role's current classification.
    cover_paragraphs = []
    for r in records:
        if r.get("region") != "cover" or r.get("is_blank"):
            continue
        cover_paragraphs.append({
            "i": r["i"], "text": r["text"],
            "font": r["eff"].get("east_asia"), "size_hp": r["eff"].get("size_hp"),
            "jc": r["eff"].get("jc"), "guess_role": cover_role(r, spec),
        })

    out_path = os.path.join(workdir, "review_candidates.json")
    with open(out_path, "w", encoding="utf-8") as f:
        json.dump({
            "headings_confirmed": headings_confirmed,
            "headings_unconfirmed": headings_unconfirmed,
            "captions_confirmed": captions_confirmed,
            "captions_unconfirmed": captions_unconfirmed,
            "possible_missed_headings": possible_missed,
            "cover_paragraphs": cover_paragraphs,
        }, f, ensure_ascii=False, indent=1)

    print(json.dumps({
        "status": "ok",
        "n_headings_confirmed": len(headings_confirmed),
        "n_headings_unconfirmed": len(headings_unconfirmed),
        "n_captions_confirmed": len(captions_confirmed),
        "n_captions_unconfirmed": len(captions_unconfirmed),
        "n_possible_missed_headings": len(possible_missed),
        "n_cover_paragraphs": len(cover_paragraphs),
        "review_candidates": out_path,
        "next_step": ("Read review_candidates.json in full (it is small). "
                       "*_confirmed: correct level/kind from CONTEXT if the "
                       "shape guess is wrong (auto-applies once corrected). "
                       "*_unconfirmed and possible_missed_headings: decide "
                       "per entry whether to promote it via overrides.json "
                       "(never modified otherwise -- stays a review hint). "
                       "cover_paragraphs: check each guess_role (title/"
                       "classification/field); fix any wrong one in "
                       "overrides.json['cover'] ({para_index: role}); role "
                       "'other' exempts a line from cover formatting. "
                       "Then run 27_apply_review.py before checking format; "
                       "if nothing needs correcting, proceed straight to "
                       "step 3."),
    }, ensure_ascii=False))


if __name__ == "__main__":
    main()
