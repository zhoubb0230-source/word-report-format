# -*- coding: utf-8 -*-
"""
Stage 2c — apply the model's structural review decisions onto structure.json.

This is the ONLY place a model's judgment calls enter the pipeline, and it is
deliberately narrow: the model may only assign a heading's LEVEL (1-4, or
null to say "not actually a heading") and a caption's KIND ("figure"/"table",
or null to say "not actually a caption"). It does not hand-write fix objects,
font values, or numbering tokens — those stay 100% deterministic, computed
downstream by checks.py from spec/format_spec.json exactly as before. This
keeps the "no model-invented formatting rules" boundary intact while letting
the model correct the two things a regex genuinely cannot always resolve:
which level a given numbering shape means IN THIS document, and whether a
paragraph regex/style missed entirely is actually a heading (a heading can
have no recognizable number AND no distinctive style — pure regex/style
matching cannot see it; 26_export_review.py's possible_missed_headings list,
built from font signals instead, is one way to find candidates worth
promoting here, but any real para_index in the document may be targeted).

Every heading/caption this script touches is stamped level_source /
caption["source"] = "model_confirmed". That matters downstream: checks.py's
continuity() only ever auto-renumbers confirmed entries (outline/style/
model_confirmed) — a "pattern"-only entry (matched a numbering shape with no
style/outline backing) is NEVER auto-edited, only hint-flagged, until this
script's explicit "model_confirmed" stamp promotes it.

Usage:
    python 27_apply_review.py <workdir> <overrides.json>

overrides.json (write this yourself — schema is intentionally tiny):
{
  "headings": {"<para_index>": 1|2|3|4|null, ...},
  "captions": {"<para_index>": "figure"|"table"|null, ...},
  "cover":    {"<para_index>": "title"|"classification"|"field"|"other"|null, ...},
  "cover_present": ["密级", "项目名称", "考核年份", ...]
}
"cover_present" (optional) is the model's semantic verdict of which required
cover items are actually FILLED IN (not just which labels appear) — the reliable
source for title-derived 项目名称/考核年份 and for spotting a still-placeholder
title. When present it drives the "missing cover info" hint; when omitted, a
deterministic baseline is used.
The "cover" section assigns each cover paragraph its FORMATTING role (checks.py
then applies the spec font/size for that role deterministically): "title" ->
方正小标宋/20磅/居中, "classification" (密级·文本编号 line) -> 仿宋/16磅,
"field" (项目名称/承担单位/…) -> 方正黑体/15磅, "other" -> leave untouched. Use
it whenever the heuristic guess_role in review_candidates.json is wrong -- e.g.
a 密级/文本编号 line with no key ("机密  202501323023") the keyword heuristic
missed. null clears the override (fall back to the heuristic).
Only list paragraphs that need correcting; omitted ones keep their
auto-detected classification from 20_extract_structure.py. Any para_index
that names a real paragraph in the document is accepted — this both
corrects already-detected candidates AND promotes a paragraph regex/style
never flagged at all (e.g. a heading with no number and no heading-ish
style) into a heading, or demotes a false positive to null.

Effect: rewrites <workdir>/structure.json in place (a pre-review copy is
kept as structure.pre_review.json) and recomputes the cover/toc/body region
boundary (promoting/demoting the document's first heading can move where
that boundary falls — everything before it is otherwise silently excluded
from all body-region checks), so that the format check (30_check_format.py)
sees the corrected level — which also determines which font spec applies
(黑体 for level 1, 楷体 for level 2, 仿宋 for level 3/4), so correcting a
level must refresh formatting, not just renumbering.

Safe to run with an empty overrides.json ({} or {"headings":{},"captions":{}})
if review found nothing to correct — it's a no-op in that case.
"""
import json
import os
import shutil
import sys

sys.path.insert(0, os.path.join(os.path.dirname(__file__), "lib"))
from headings import RE_CAPTION, parse_leading_label
from structure import tag_regions


_COVER_ROLES = ("title", "classification", "field", "other")


def _validate(overrides, records_by_i):
    if not isinstance(overrides, dict):
        raise ValueError("overrides.json must be a JSON object")
    headings = overrides.get("headings") or {}
    captions = overrides.get("captions") or {}
    cover = overrides.get("cover") or {}
    if not isinstance(headings, dict) or not isinstance(captions, dict) \
            or not isinstance(cover, dict):
        raise ValueError("'headings', 'captions' and 'cover' must be objects")
    for k, v in headings.items():
        i = int(k)  # raises ValueError if not int-like
        if i not in records_by_i:
            raise ValueError("headings override references unknown para_index %s" % k)
        if v is not None and v not in (1, 2, 3, 4):
            raise ValueError("headings override level must be 1-4 or null, got %r for %s" % (v, k))
    for k, v in captions.items():
        i = int(k)
        if i not in records_by_i:
            raise ValueError("captions override references unknown para_index %s" % k)
        if v is not None and v not in ("figure", "table"):
            raise ValueError("captions override kind must be 'figure'/'table'/null, got %r for %s" % (v, k))
    for k, v in cover.items():
        i = int(k)
        if i not in records_by_i:
            raise ValueError("cover override references unknown para_index %s" % k)
        if v is not None and v not in _COVER_ROLES:
            raise ValueError("cover override role must be one of %s or null, got %r for %s"
                             % (_COVER_ROLES, v, k))
    return headings, captions, cover


def _full_text(r):
    # structure.json truncates text to 60 chars; the leading numbering
    # label is always within the first handful of characters, so the
    # truncation never loses it.
    return r.get("text") or ""


def main():
    if len(sys.argv) < 3:
        print(json.dumps({"status": "error",
                          "error": "usage: 27_apply_review.py <workdir> <overrides.json>"}))
        sys.exit(1)
    workdir = sys.argv[1]
    overrides_path = sys.argv[2]

    structure_path = os.path.join(workdir, "structure.json")
    with open(structure_path, encoding="utf-8") as f:
        structure = json.load(f)
    records = structure["records"]
    records_by_i = {r["i"]: r for r in records}

    with open(overrides_path, encoding="utf-8") as f:
        overrides = json.load(f)
    headings_ov, captions_ov, cover_ov = _validate(overrides, records_by_i)

    for k, v in headings_ov.items():
        r = records_by_i[int(k)]
        if v is None:
            r["is_heading"] = False
            r["level"] = None
            r["level_source"] = None
            r["num_raw"] = None
        else:
            r["is_heading"] = True
            r["level"] = v
            # The model just explicitly reviewed this paragraph with full
            # document context — that is strictly more reliable than a bare
            # style/outline match, so it is auto-fixable from here on
            # (continuity() only ever hint-flags "pattern"-sourced entries).
            r["level_source"] = "model_confirmed"
            r["num_raw"] = parse_leading_label(_full_text(r))

    for k, v in captions_ov.items():
        r = records_by_i[int(k)]
        if v is None:
            r["caption"] = None
        else:
            mc = RE_CAPTION.match(_full_text(r))
            rest = mc.group(3).strip() if mc else ""
            r["caption"] = {
                "kind": v,
                "num_raw": mc.group(2) if mc else None,
                "has_content": bool(rest),
                "source": "model_confirmed",
            }

    for k, v in cover_ov.items():
        r = records_by_i[int(k)]
        if v is None:
            # clear the override -> fall back to the heuristic cover_role()
            r["cover_role"] = None
        else:
            r["cover_role"] = v
            # keep is_title consistent so the title-block flag and the cover role
            # never disagree (checks.cover_role treats an explicit role as
            # authoritative, but is_title is also read elsewhere).
            r["is_title"] = (v == "title")

    # Optional AI top-level semantic completeness verdict: the required cover
    # items the model confirms are actually filled in (esp. title-derived
    # 项目名称/考核年份, which no substring match can verify). Authoritative for
    # the "missing cover info" hint when present.
    cover_present = overrides.get("cover_present")
    if cover_present is not None:
        if not isinstance(cover_present, list) or not all(isinstance(x, str) for x in cover_present):
            raise ValueError("cover_present must be a list of field-name strings")
        structure["cover_present"] = cover_present

    # Promoting a previously-undetected paragraph to a heading (or demoting the
    # old "first heading") can move where the cover/body boundary falls, so the
    # region tags must be recomputed with the same logic 20_extract_structure.py
    # used — otherwise paragraphs between a newly promoted early heading and the
    # old first-detected one stay mis-tagged "cover", invisible to every
    # body-region check (font, indent, continuity).
    tag_regions(records)

    bak = os.path.join(workdir, "structure.pre_review.json")
    if not os.path.exists(bak):
        shutil.copyfile(structure_path, bak)
    with open(structure_path, "w", encoding="utf-8") as f:
        json.dump(structure, f, ensure_ascii=False)

    print(json.dumps({
        "status": "ok",
        "n_heading_overrides": len(headings_ov),
        "n_caption_overrides": len(captions_ov),
        "n_cover_overrides": len(cover_ov),
    }, ensure_ascii=False))


if __name__ == "__main__":
    main()
