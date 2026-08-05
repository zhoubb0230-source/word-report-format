# -*- coding: utf-8 -*-
"""
structure.py — shared structure-pipeline helper used by BOTH the extraction
stage (20_extract_structure.py) and the model-review-apply stage
(27_apply_review.py).

This region-tagging MUST agree byte-for-byte between the two stages, but was
previously copy-pasted into each script (27's own comments read "Mirrors
20…exactly" / "Must be redone here"), which is exactly the kind of duplication
that drifts. It lives here now so there is a single source of truth:

  * tag_regions(records)      — split records into cover / toc / body by the
    first-TOC / first-heading boundary. Promoting or demoting the document's
    first heading moves that boundary, so 27 must re-run the identical logic
    after applying the model's overrides.

Deterministic. No model calls.
"""


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
