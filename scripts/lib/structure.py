# -*- coding: utf-8 -*-
"""
structure.py — shared structure-pipeline helpers used by BOTH the extraction
stage (20_extract_structure.py) and the model-review-apply stage
(27_apply_review.py).

These two operations MUST agree byte-for-byte, but were previously copy-pasted
into each script (27's own comments read "Mirrors 20…exactly" / "Must be redone
here"), which is exactly the kind of duplication that drifts. They live here now
so there is a single source of truth:

  * tag_regions(records)      — split records into cover / toc / body by the
    first-TOC / first-heading boundary. Promoting or demoting the document's
    first heading moves that boundary, so 27 must re-run the identical logic
    after applying the model's overrides.
  * write_shards(shard_dir, records, shard_size) — (re)generate the paragraph
    shards sub-agents consume, atomically replacing any previous shards.

Deterministic. No model calls.
"""
import json
import os


def tag_regions(records):
    """Assign each record a "region" of "cover" / "toc" / "body".

    cover = every non-TOC paragraph BEFORE the first TOC (or, if there is no
    TOC, before the first detected heading); toc = any TOC-styled/field
    paragraph; body = everything else. Mutates records in place.
    """
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


def write_shards(shard_dir, records, shard_size):
    """(Re)write <shard_dir>/shard_NNN.json in slices of shard_size records.

    Removes any pre-existing shard_*.json first so a re-shard never leaves a
    stale tail shard behind. Returns the list of shard file names written.
    """
    os.makedirs(shard_dir, exist_ok=True)
    for f in os.listdir(shard_dir):
        if f.startswith("shard_") and f.endswith(".json"):
            os.remove(os.path.join(shard_dir, f))
    names = []
    for si, start in enumerate(range(0, len(records), shard_size)):
        chunk = records[start:start + shard_size]
        name = "shard_%03d.json" % si
        with open(os.path.join(shard_dir, name), "w", encoding="utf-8") as f:
            json.dump({"shard_id": si, "range": [chunk[0]["i"], chunk[-1]["i"]],
                       "records": chunk}, f, ensure_ascii=False)
        names.append(name)
    return names
