# -*- coding: utf-8 -*-
"""Stage 4b — structural sanity check of the produced formatted.docx.

Usage:
    python 45_validate_output.py <workdir>

Runs AFTER 40_apply_fixes.py and BEFORE 50_finalize.py. Confirms the edited
document is not corrupted before it is delivered to the user — the failure mode
this guards against is real Microsoft Word opening the file and prompting
"内容有问题，需要修复". Checks (all lightweight, pure lxml — no python-docx):

  1. zip integrity        —— ZipFile.testzip() reports no bad CRC.
  2. required parts       —— [Content_Types].xml + word/document.xml present.
  3. XML well-formedness  —— every .xml / .rels part parses.
  4. paragraph parity     —— body-paragraph count in formatted.docx equals the
     pre-edit working.docx. Format / renumber / comment ops never add or remove
     a <w:p>, so a mismatch means the structure was damaged.
  5. numbering-clamp collapse invariant (方案C 甲法) —— every auto-numbered
     paragraph that received an indent fix must NOT still inherit a hanging from
     the NUMBERING layer in the output (that would mean the clone-clamp didn't
     take). Verified with the unified cascade + provenance, so a real leak fails
     hard; a style-inherited hanging (known-unfixed, Word-gated) is a non-fatal
     note only.
  5b. 元素顺序 —— OOXML 的复杂类型基本都是 xsd:sequence，子元素乱序时文件仍然
     良构（第 3 项查不出），但真实 Word 拒绝打开并提示"发现无法读取的内容"。
  5c. 批注一致性 + 编号身份链接 —— 悬空/重复的批注 id、不配对的批注区间，以及两条
     abstractNum 抢同一个 styleLink / 同一个级别 pStyle；同样是"良构但 Word 拒绝
     打开"的一类。
  6. 全坍缩不变量 (方案C 阶段3) —— 被指派 canonical 样式的段落，该样式承载的每个
     属性都必须由【样式层】供给；仍由 direct/numbering 供给且值≠canonical 即泄漏，
     硬失败。这是没有语料、没有 Word 时的确定性安全网（供给层是纯 XML 事实）。

Writes <workdir>/validate_report.json and prints a single-line JSON summary.
Exits non-zero (2) if any check fails, so the caller stops before finalizing.
"""
import json
import os
import sys
import zipfile

sys.path.insert(0, os.path.join(os.path.dirname(__file__), "lib"))

from lxml import etree
from docxcommon import (qn, iter_body_paragraphs, StyleResolver,
                        load_numbering_levels, get_style_id, get_pPr, get_mark_rpr,
                        read_ppr, char_styles_overriding, order_violations, in_textbox,
                        comment_problems, numbering_link_problems)
from cascade import resolve_ppr_with_provenance
import canonstyles
from checks import load_default_spec

REQUIRED_PARTS = ("[Content_Types].xml", "word/document.xml")

# canonical 样式承载的每一类属性对应哪些 pPr 键（用于全坍缩不变量逐键比对）。
_GOVERNED_KEYS = {
    "ind": ("first_line_chars", "first_line", "left_chars", "left",
            "start_chars", "start", "right_chars", "right",
            "end_chars", "end", "hanging_chars", "hanging"),
    "line": ("line", "line_rule"),
    "space_before_after": ("space_before", "space_after",
                           "space_before_lines", "space_after_lines"),
    "jc": ("jc",),
    "outline": ("outline",),
}


def _count_body_paragraphs(zf, name="word/document.xml"):
    """Parse document.xml straight from the zip and count body paragraphs with
    the SAME iterator extraction/apply use, so the numbers are comparable."""
    parser = etree.XMLParser(remove_blank_text=False, huge_tree=True)
    with zf.open(name) as f:
        root = etree.parse(f, parser).getroot()
    return sum(1 for _ in iter_body_paragraphs(root))


def _read_root(zf, name):
    parser = etree.XMLParser(remove_blank_text=False, huge_tree=True)
    with zf.open(name) as f:
        return etree.parse(f, parser).getroot()


def _check_numbering_clamp(zf, indent_fix_indices):
    """方案C 甲法坍缩不变量：在 OUTPUT 上重解析每个"拿到缩进修复的自动编号段"，
    断言其有效 hanging **不再由【编号层】供给**——否则说明克隆钳没生效（numbering.xml
    没去 hanging，或段落 numPr 没改指克隆）。用 cascade provenance 精确定位供给层：

      * hanging 仍由 `numbering` 供给 → LEAK（致命：钳失败，退2）。
      * hanging 由 `style`/`direct` 供给 → note（非致命）：样式层继承 hanging 属**已知
        未修**（待全角色 canonical 样式注入，Word-gated），不该在此硬失败误杀交付。

    只查"拿到缩进修复"的段落（避免对未触碰的合法列表段误报）。返回 (leaks, notes)。"""
    names = set(zf.namelist())
    if "word/document.xml" not in names:
        return [], []
    doc_root = _read_root(zf, "word/document.xml")
    styles_root = _read_root(zf, "word/styles.xml") if "word/styles.xml" in names else None
    numbering_root = (_read_root(zf, "word/numbering.xml")
                      if "word/numbering.xml" in names else None)
    resolver = StyleResolver(styles_root)
    levels = load_numbering_levels(numbering_root)
    para_by_idx = {i: p for i, p in iter_body_paragraphs(doc_root)}

    leaks, notes = [], []
    for idx in sorted(set(indent_fix_indices)):
        p = para_by_idx.get(idx)
        if p is None:
            continue
        eff, owner = resolve_ppr_with_provenance(
            resolver, get_style_id(p), get_pPr(p), get_mark_rpr(p), levels)
        nid = eff.get("num_id")
        if not nid or nid == "0":
            continue  # 非自动编号：不属甲法钳的职责范围
        for key in ("hanging", "hanging_chars"):
            v = eff.get(key)
            if v in (None, 0, "0"):
                continue
            rec = {"para_index": idx, "key": key, "value": v, "owner": owner.get(key)}
            (leaks if owner.get(key) == "numbering" else notes).append(rec)
    return leaks, notes


def _check_element_order(zf):
    """OOXML **元素顺序**自检——这是"Word 提示发现无法读取的内容"的头号成因。

    WordprocessingML 的复杂类型几乎都是 `xsd:sequence`：把 `<w:suff>` 写在
    `<w:numFmt>` 前面、把 `<w:numPr>` 插到 `<w:tabs>` 后面，文件仍是**良构 XML**、
    zip 也完好、上面第 3 项检查照样通过——但真实 Word 会拒绝打开。良构 ≠ 合法，
    所以这一项单独查。返回 [(part, 容器, 乱序子元素, 它前面的子元素)]。

    只查 `docxcommon.ELEMENT_ORDER` 登记的容器、跳过未登记的子元素：顺序表不全时
    只会漏报，不会把 Word 本来打得开的文档误判成坏文件。"""
    out = []
    for part in ("word/document.xml", "word/styles.xml", "word/numbering.xml",
                 "word/settings.xml"):
        if part not in set(zf.namelist()):
            continue
        for parent, child, prev in order_violations(_read_root(zf, part)):
            out.append({"part": part, "container": parent,
                        "element": child, "after": prev})
    return out


def _check_comments_and_numbering(zf):
    """批注一致性 + 编号身份链接唯一性——都是"良构但 Word 拒绝打开"的成因，
    与元素顺序同类，`45` 里一并硬查。规则实现在 `docxcommon`，与 diagnose_docx 共用。"""
    names = set(zf.namelist())
    out = []
    if "word/document.xml" in names:
        comments = (_read_root(zf, "word/comments.xml")
                    if "word/comments.xml" in names else None)
        out.extend(comment_problems(_read_root(zf, "word/document.xml"), comments))
    if "word/numbering.xml" in names:
        out.extend(numbering_link_problems(_read_root(zf, "word/numbering.xml")))
    return out


def _canonical_expectations(spec):
    """{styleId: (governs, canonical pPr dict)} —— 期望值直接从 canonical 样式定义
    解析（`canonstyles` 是唯一真源），不在这里重写一遍 spec 换算。"""
    out = {}
    for d in canonstyles.canonical_style_defs(spec):
        el = etree.fromstring(
            ('<w:wrap xmlns:w="%s">%s</w:wrap>'
             % ("http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                d["xml"])).encode("utf-8"))[0]
        out[d["id"]] = (d["governs"], read_ppr(el.find(qn("w:pPr"))))
    return out


# run 级别：样式承载 fonts/size/bold 时，run 上不许再有同名直接覆盖或抢戏的字符样式。
# **字体看的是 rFonts 的属性、不是元素本身**：中文 Word 到处都是 `<w:rFonts w:hint="eastAsia"/>`，
# `hint` 只说明歧义字符按中文还是西文取字体，**不覆盖任何字体**，留着它完全正常。按"元素在不
# 在"判会把成百上千个合法段落报成泄漏（用户实测 295 段），把交付整个挡下来。
_FONT_ATTRS = ("w:ascii", "w:asciiTheme", "w:hAnsi", "w:hAnsiTheme",
               "w:eastAsia", "w:eastAsiaTheme", "w:cs", "w:cstheme")
_RUN_GOVERNED_TAGS = {"size": ("w:sz", "w:szCs"), "bold": ("w:b", "w:bCs")}


def _run_level_leaks(idx, p, style_id, governs, overriding_char_styles):
    """canonical 样式承载的**字体/字号/加粗**必须由样式供给，run 上不得再有直接覆盖。

    这一条是"同一个三级标题前半部分加粗、后半部分不加粗"的机器化断言：那种半粗半细
    正是几个 run 各带一份直接 rPr（或一个设了字体/加粗的字符样式）造成的，段落级 pPr
    检查看不见它。

    三处刻意的放行，都是为了不误杀合法文档：
      * 只看**有可见文字**的 run——批注引用、域字符那些 run 不承载正文格式；
      * **文本框（`w:txbxContent`）里的 run 跳过**——它们属于另一个逻辑段落，apply 也
        从不改它们（`_iter_runs` 同样跳过），在这里报泄漏就是自相矛盾；
      * 字符样式只算 `overriding_char_styles` 里那些（真设了字体/字号/加粗的），超链接
        色、批注引用之类不参与。"""
    leaks = []
    for r in p.iter(qn("w:r")):
        if in_textbox(r, stop_at=p):
            continue
        if not any((t.text or "").strip() for t in r.findall(qn("w:t"))):
            continue
        rpr = r.find(qn("w:rPr"))
        if rpr is None:
            continue
        rs = rpr.find(qn("w:rStyle"))
        if rs is not None and rs.get(qn("w:val")) in overriding_char_styles:
            leaks.append({"para_index": idx, "style": style_id, "key": "rStyle",
                          "value": rs.get(qn("w:val")), "owner": "char_style",
                          "canonical": None})
        if governs.get("fonts"):
            rf = rpr.find(qn("w:rFonts"))
            got = [a for a in _FONT_ATTRS if rf is not None and rf.get(qn(a)) is not None]
            if got:
                leaks.append({"para_index": idx, "style": style_id, "key": "rFonts",
                              "value": ",".join(a.split(":")[1] for a in got),
                              "owner": "direct", "canonical": None})
        for flag, tags in _RUN_GOVERNED_TAGS.items():
            if not governs.get(flag):
                continue
            for tag in tags:
                if rpr.find(qn(tag)) is not None:
                    leaks.append({"para_index": idx, "style": style_id,
                                  "key": tag.split(":")[1], "value": "direct",
                                  "owner": "direct", "canonical": None})
    return leaks


def _check_canonical_collapse(zf, spec):
    """方案C 阶段3 **全坍缩不变量**：凡是被指派了 canonical 样式的段落，该样式承载的
    每一个属性都必须由**样式层**供给。

    这是"无语料安全网"——沙箱里没有 Word 可渲染（handoff §1），但"值由哪一层供给"
    是纯 XML 事实，可以确定性断言。apply 之后在**产物**上重解析四层 cascade
    （docDefaults→样式→编号层→直接）并取 provenance：某个 canonical 属性若仍由
    `direct` 或 `numbering` 供给，说明"清直接覆盖"或"甲法克隆钳"漏了，Word 里就会
    看到规范外的缩进/行距——当场硬失败，别交付。

    唯一放行的情况：供给值**恰好等于 canonical 值**。甲法克隆钳会在编号层显式写
    `left=0`（中和级别缩进正是它的手段），这与 canonical 的 left=0 一致，不算泄漏。

    返回 leaks 列表；空列表＝全部坍缩到样式层。"""
    names = set(zf.namelist())
    if "word/document.xml" not in names or not spec:
        return []
    expectations = _canonical_expectations(spec)
    doc_root = _read_root(zf, "word/document.xml")
    styles_root = _read_root(zf, "word/styles.xml") if "word/styles.xml" in names else None
    numbering_root = (_read_root(zf, "word/numbering.xml")
                      if "word/numbering.xml" in names else None)
    resolver = StyleResolver(styles_root)
    levels = load_numbering_levels(numbering_root)
    overriding_char_styles = char_styles_overriding(styles_root)

    leaks = []
    for idx, p in iter_body_paragraphs(doc_root):
        sid = get_style_id(p)
        exp = expectations.get(sid)
        if exp is None:
            continue          # 未指派 canonical 样式的段落不在本不变量的射程内
        governs, canon = exp
        leaks.extend(_run_level_leaks(idx, p, sid, governs, overriding_char_styles))
        eff, owner = resolve_ppr_with_provenance(
            resolver, sid, get_pPr(p), get_mark_rpr(p), levels)
        for flag, keys in _GOVERNED_KEYS.items():
            if not governs.get(flag):
                continue
            for key in keys:
                layer = owner.get(key)
                if layer not in ("direct", "numbering"):
                    continue
                if eff.get(key) == canon.get(key):
                    continue  # 值与 canonical 一致（如克隆钳写的 left=0）
                leaks.append({"para_index": idx, "style": sid, "key": key,
                              "value": eff.get(key), "owner": layer,
                              "canonical": canon.get(key)})
    return leaks


def validate(formatted_path, reference_path=None, indent_fix_indices=()):
    errors = []
    info = {}

    if not os.path.exists(formatted_path):
        return {"status": "error", "ok": False,
                "errors": ["formatted.docx not found: %s" % formatted_path]}

    # 1. zip integrity + open
    try:
        zf = zipfile.ZipFile(formatted_path)
    except zipfile.BadZipFile as e:
        return {"status": "error", "ok": False,
                "errors": ["not a valid zip/docx: %s" % e]}
    with zf:
        bad = zf.testzip()
        if bad is not None:
            errors.append("corrupted zip entry (bad CRC): %s" % bad)

        names = set(zf.namelist())

        # 2. required parts present
        for req in REQUIRED_PARTS:
            if req not in names:
                errors.append("missing required part: %s" % req)

        # 3. XML well-formedness of every xml/rels part
        parser = etree.XMLParser(remove_blank_text=False, huge_tree=True)
        malformed = []
        for name in names:
            if not (name.endswith(".xml") or name.endswith(".rels")):
                continue
            try:
                with zf.open(name) as f:
                    etree.parse(f, parser)
            except etree.XMLSyntaxError as e:
                malformed.append({"part": name, "error": str(e)})
        if malformed:
            errors.append("malformed XML in %d part(s)" % len(malformed))
            info["malformed"] = malformed

        # 4. paragraph parity vs the pre-edit reference
        if "word/document.xml" in names:
            try:
                n_out = _count_body_paragraphs(zf)
                info["paragraphs_out"] = n_out
            except etree.XMLSyntaxError as e:
                errors.append("cannot parse word/document.xml: %s" % e)
                n_out = None
            if reference_path and os.path.exists(reference_path) and n_out is not None:
                try:
                    with zipfile.ZipFile(reference_path) as rzf:
                        n_ref = _count_body_paragraphs(rzf)
                    info["paragraphs_reference"] = n_ref
                    if n_out != n_ref:
                        errors.append(
                            "paragraph count changed: formatted=%d, original=%d "
                            "(format/renumber/comment ops must not add or remove "
                            "paragraphs — structure likely damaged)" % (n_out, n_ref))
                except (zipfile.BadZipFile, KeyError, etree.XMLSyntaxError) as e:
                    info["paragraph_parity_skipped"] = "reference unreadable: %s" % e

        # 5. 方案C 甲法坍缩不变量：拿到缩进修复的自动编号段，有效 hanging 不得仍由
        #    编号层供给（那意味克隆钳没生效）。样式层继承 hanging 是已知未修，作非致命
        #    note。检查失败不阻断交付；只有"编号层泄漏"才致命。
        if indent_fix_indices:
            try:
                leaks, notes = _check_numbering_clamp(zf, indent_fix_indices)
                if leaks:
                    errors.append(
                        "编号层未钳干净：%d 段有效 hanging 仍来自编号层（克隆钳失败）"
                        % len(leaks))
                    info["numbering_leaks"] = leaks
                if notes:
                    info["style_hanging_notes"] = notes
            except (etree.XMLSyntaxError, KeyError) as e:
                info["clamp_invariant_skipped"] = str(e)

        # 5b. 元素顺序（OOXML sequence）：良构但乱序的 XML 会让 Word 提示"发现无法
        #     读取的内容"。这是第 3 项"良构"检查覆盖不到的一类损坏，单列一项。
        try:
            order_bad = _check_element_order(zf)
            if order_bad:
                errors.append("元素顺序违反 OOXML sequence：%d 处（Word 会提示"
                              "内容无法读取）" % len(order_bad))
                info["element_order_violations"] = order_bad[:50]
        except (etree.XMLSyntaxError, KeyError) as e:
            info["element_order_skipped"] = str(e)

        # 5c. 批注一致性（悬空引用/重复 id/区间不配对）与编号身份链接唯一性。
        #     同 5b 一类：XML 良构、zip 完好，Word 照样报"发现无法读取的内容"。
        try:
            struct_bad = _check_comments_and_numbering(zf)
            if struct_bad:
                errors.append("批注/编号结构损坏：%d 处（Word 会提示内容无法读取）"
                              % len(struct_bad))
                info["structural_problems"] = struct_bad[:50]
        except (etree.XMLSyntaxError, KeyError) as e:
            info["structural_check_skipped"] = str(e)

        # 6. 全坍缩不变量（方案C 阶段3）：指派了 canonical 样式的段落，其样式承载的
        #    属性必须由样式层供给；仍由 direct/numbering 供给且值不等于 canonical
        #    ＝没钳干净，Word 里会泄漏出规范外的格式 → 硬失败。
        try:
            spec = load_default_spec()
        except (OSError, ValueError):
            spec = None
        if spec:
            try:
                canon_leaks = _check_canonical_collapse(zf, spec)
                if canon_leaks:
                    errors.append(
                        "canonical 属性未坍缩到样式层：%d 处仍由直接/编号层供给"
                        % len(canon_leaks))
                    info["canonical_leaks"] = canon_leaks[:50]
            except (etree.XMLSyntaxError, KeyError) as e:
                info["canonical_invariant_skipped"] = str(e)

    ok = not errors
    return {"status": "ok" if ok else "error", "ok": ok,
            "formatted_docx": formatted_path, "errors": errors, **info}


def main():
    if len(sys.argv) < 2:
        print(json.dumps({"status": "error",
                          "error": "usage: 45_validate_output.py <workdir>"}))
        sys.exit(1)
    workdir = sys.argv[1]
    formatted = os.path.join(workdir, "formatted.docx")

    reference = None
    meta_path = os.path.join(workdir, "meta.json")
    if os.path.exists(meta_path):
        try:
            meta = json.load(open(meta_path, encoding="utf-8"))
            reference = meta.get("working_docx")
        except (OSError, ValueError):
            reference = None

    # Indices of paragraphs that received an indent fix (same trigger set 40 uses
    # for the numbering clamp) — the collapse invariant only inspects these.
    indent_fix_indices = []
    fixes_path = os.path.join(workdir, "fixes.json")
    if os.path.exists(fixes_path):
        try:
            for fx in json.load(open(fixes_path, encoding="utf-8")):
                if fx.get("op") == "format" and (
                        fx.get("set_first_line_chars") is not None
                        or fx.get("clear_left_indent")
                        or fx.get("clear_right_indent")
                        or fx.get("set_left_chars") is not None):
                    idx = fx.get("para_index")
                    if idx is not None:
                        indent_fix_indices.append(idx)
        except (OSError, ValueError):
            pass

    report = validate(formatted, reference, indent_fix_indices)
    json.dump(report, open(os.path.join(workdir, "validate_report.json"), "w",
                           encoding="utf-8"), ensure_ascii=False, indent=1)
    print(json.dumps(report, ensure_ascii=False))
    sys.exit(0 if report.get("ok") else 2)


if __name__ == "__main__":
    main()
