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

Usage:
    python 27_apply_review.py <workdir> <overrides.json>

overrides.json (write this yourself — schema is intentionally tiny):
{
  "headings": {"<para_index>": 1|2|3|4|null, ...},
  "captions": {"<para_index>": "figure"|"table"|null, ...}
}
Only list paragraphs that need correcting; omitted ones keep their
auto-detected classification from 20_extract_structure.py. Any para_index
that names a real paragraph in the document is accepted — this both
corrects already-detected candidates AND promotes a paragraph regex/style
never flagged at all (e.g. a heading with no number and no heading-ish
style) into a heading, or demotes a false positive to null.

Effect: rewrites <workdir>/structure.json in place (a pre-review copy is
kept as structure.pre_review.json), recomputes the cover/toc/body region
boundary (promoting/demoting the document's first heading can move where
that boundary falls — everything before it is otherwise silently excluded
from all body-region checks), AND regenerates <workdir>/shards/*, so that
whichever check path runs next (full or sharded) sees the corrected level —
which also determines which font spec applies (黑体 for level 1, 楷体 for
level 2, 仿宋 for level 3/4), so correcting a level must refresh formatting,
not just renumbering.

Safe to run with an empty overrides.json ({} or {"headings":{},"captions":{}})
if review found nothing to correct — it's a no-op in that case.
"""
import json
import os
import shutil
import sys

sys.path.insert(0, os.path.join(os.path.dirname(__file__), "lib"))
from headings import RE_CAPTION, parse_leading_label


def _validate(overrides, records_by_i):
    if not isinstance(overrides, dict):
        raise ValueError("overrides.json must be a JSON object")
    headings = overrides.get("headings") or {}
    captions = overrides.get("captions") or {}
    if not isinstance(headings, dict) or not isinstance(captions, dict):
        raise ValueError("'headings' and 'captions' must be objects")
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
    return headings, captions


def _full_text(r):
    # structure.json truncates text to 60 chars; the leading numbering
    # label is always within the first handful of characters, so the
    # truncation never loses it.
    return r.get("text") or ""


def _retag_regions(records):
    # Mirrors 20_extract_structure.py's region-tagging pass exactly. Must be
    # redone here: promoting a previously-undetected paragraph to a heading
    # (or demoting the old "first heading") can move where the cover/body
    # boundary actually falls. Without this, paragraphs between a newly
    # promoted early heading and the old first-detected one would stay
    # mis-tagged "cover" — invisible to every body-region check (font,
    # indent, continuity) even though they are now clearly body content.
    first_toc_idx = next((r["i"] for r in records if r["is_toc"]), None)
    first_heading_idx = next((r["i"] for r in records if r["is_heading"]), None)
    cover_end = first_toc_idx if first_toc_idx is not None else first_heading_idx
    for r in records:
        if r["is_toc"]:
            r["region"] = "toc"
        elif cover_end is not None and r["i"] < cover_end:
            r["region"] = "cover"
        else:
            r["region"] = "body"


def _rewrite_shards(workdir, records):
    shard_dir = os.path.join(workdir, "shards")
    if not os.path.isdir(shard_dir):
        return
    existing = sorted(f for f in os.listdir(shard_dir) if f.startswith("shard_"))
    shard_size = 400
    if existing:
        with open(os.path.join(shard_dir, existing[0]), encoding="utf-8") as f:
            shard_size = max(len(json.load(f).get("records", [])), 1)
    for f in existing:
        os.remove(os.path.join(shard_dir, f))
    for si, start in enumerate(range(0, len(records), shard_size)):
        chunk = records[start:start + shard_size]
        name = "shard_%03d.json" % si
        with open(os.path.join(shard_dir, name), "w", encoding="utf-8") as f:
            json.dump({"shard_id": si, "range": [chunk[0]["i"], chunk[-1]["i"]],
                       "records": chunk}, f, ensure_ascii=False)


def main():
    workdir = sys.argv[1]
    overrides_path = sys.argv[2]

    structure_path = os.path.join(workdir, "structure.json")
    with open(structure_path, encoding="utf-8") as f:
        structure = json.load(f)
    records = structure["records"]
    records_by_i = {r["i"]: r for r in records}

    with open(overrides_path, encoding="utf-8") as f:
        overrides = json.load(f)
    headings_ov, captions_ov = _validate(overrides, records_by_i)

    for k, v in headings_ov.items():
        r = records_by_i[int(k)]
        if v is None:
            r["is_heading"] = False
            r["level"] = None
            r["num_raw"] = None
        else:
            r["is_heading"] = True
            r["level"] = v
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
            }

    _retag_regions(records)

    bak = os.path.join(workdir, "structure.pre_review.json")
    if not os.path.exists(bak):
        shutil.copyfile(structure_path, bak)
    with open(structure_path, "w", encoding="utf-8") as f:
        json.dump(structure, f, ensure_ascii=False)

    _rewrite_shards(workdir, records)

    print(json.dumps({
        "status": "ok",
        "n_heading_overrides": len(headings_ov),
        "n_caption_overrides": len(captions_ov),
    }, ensure_ascii=False))


if __name__ == "__main__":
    main()
