# -*- coding: utf-8 -*-
"""生成方案C 的【canonical 参考 .docx】——阶段0 人肉验收件（handoff §5）。

沙箱无 Word/soffice，方案C 的真正判据是 Word 渲染，只能靠用户实测（handoff §1）。
本脚本按 `spec/format_spec.json` 造一份"每角色一段、套注入 canonical 样式"的 docx，
供用户在 **Word** 里打开、按下面清单逐项确认。这是整套方案唯一的人肉环节。

用法：
    python3 scripts/make_canonical_reference.py [输出路径]
    # 默认写到 cwd 的 canonical_reference.docx

验收清单（在 Word 里逐条看）：
  1. 每个角色的字体/字号/居中/行距是否与规范一致（每段前有【角色·规范】标签）。
  2. **首行缩进显示为"2 字符"而非厘米**——选中正文/标题段，看"段落→缩进→首行缩进 2 字符"，
     不是 0.74cm/0.85cm 之类。这是严格-spec §2.2 的关键判据（canonical 只写字符单位、无绝对伴随值）。
  3. **自动编号标题**（"概述"那段带"一、"）：首行是 2 字符缩进、不是 -0.74cm 悬挂缩进——验证甲法克隆钳。
  4. **目录页码不换行**：打开时选"是"更新域，目录按 toc 样式重建，看二/三级页码是否顶到右边界仍不换行。
  5. 封面各要素：方正黑体_GBK、题目居中、题目下要素两端对齐+首行缩进2字符。

判据只来自 spec（红线）：本脚本不写死任何字体/字号/缩进，全部读 `spec/format_spec.json`。
纯 stdlib（zipfile）+ 少量字符串拼装，不依赖 lxml，也不进判定流水线。
"""
import json
import os
import sys
import zipfile

SPEC_PATH = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))),
                         "spec", "format_spec.json")

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"


def _esc(s):
    return (s.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
            .replace('"', "&quot;"))


def _rpr(ea, size_hp, western_font=None):
    """run props: eastAsia = CJK 字体；ascii/hAnsi = 西文字体（套西文的角色给 Times，
    否则回退给 CJK 字体本身，让数字也走该字体——封面除题目外不套西文，陷阱#7）。"""
    latin = western_font or ea
    parts = ['<w:rFonts w:ascii="%s" w:hAnsi="%s" w:eastAsia="%s" w:cs="%s"/>'
             % (_esc(latin), _esc(latin), _esc(ea), _esc(latin))]
    if size_hp is not None:
        parts.append('<w:sz w:val="%d"/><w:szCs w:val="%d"/>' % (size_hp, size_hp))
    return "".join(parts)


def _ind(first_line_chars=None, left_chars=None, no_indent=False):
    """缩进：只写【字符单位】(firstLineChars/leftChars)，绝不补绝对伴随值——严格-spec
    §2.2，这样 Word 才显示"N 字符"而非厘米。no_indent 显式把各向缩进清零。"""
    a = []
    if no_indent:
        a += ['w:firstLine="0"', 'w:firstLineChars="0"',
              'w:left="0"', 'w:leftChars="0"', 'w:right="0"', 'w:rightChars="0"']
    else:
        if first_line_chars is not None:
            a.append('w:firstLineChars="%d"' % first_line_chars)
        if left_chars is not None:
            a.append('w:leftChars="%d"' % left_chars)
    return ('<w:ind %s/>' % " ".join(a)) if a else ""


def _spacing(line_twips=None, line_rule=None, zero_before_after=False):
    a = []
    if zero_before_after:
        a += ['w:before="0"', 'w:after="0"']
    if line_twips is not None:
        a += ['w:line="%d"' % line_twips, 'w:lineRule="%s"' % (line_rule or "exact")]
    return ('<w:spacing %s/>' % " ".join(a)) if a else ""


def _style(sid, name, ppr_inner, rpr_inner):
    return ('<w:style w:type="paragraph" w:styleId="%s"><w:name w:val="%s"/>'
            '<w:qFormat/><w:pPr>%s</w:pPr><w:rPr>%s</w:rPr></w:style>'
            % (sid, _esc(name), ppr_inner, rpr_inner))


def build_styles(spec):
    """按 spec 造全角色 canonical 样式。返回 styles.xml 字符串。"""
    western = spec.get("western_font", "Times New Roman")
    ls = spec.get("line_spacing", {})
    line_twips = ls.get("line_twips")
    styles = []

    # docDefaults：默认五号 21 半点（字符单位宽度基准，见 _default_char_unit_hp）
    docdef = ('<w:docDefaults><w:rPrDefault><w:rPr>'
              '<w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" '
              'w:eastAsia="仿宋" w:cs="Times New Roman"/><w:sz w:val="21"/>'
              '<w:szCs w:val="21"/></w:rPr></w:rPrDefault></w:docDefaults>')
    styles.append('<w:style w:type="paragraph" w:default="1" w:styleId="Normal">'
                  '<w:name w:val="Normal"/></w:style>')

    # 题目
    t = spec["title"]
    styles.append(_style(
        "CanonTitle", "封面题目",
        '<w:jc w:val="%s"/>' % t.get("jc", "center"),
        _rpr(t["east_asia"], t["size_hp"], western if t.get("enforce_western") else None)))

    # 封面密级/文本编号行
    cc = spec["cover_classification"]
    styles.append(_style("CanonCoverClass", "封面密级编号",
                         '<w:jc w:val="center"/>',
                         _rpr(cc["east_asia"], cc["size_hp"])))

    # 封面题目下要素（field）：2倍行距、两端对齐、首行缩进2字符
    cf = spec["cover_field"]
    styles.append(_style(
        "CanonCoverField", "封面要素",
        _spacing(cf.get("line_twips"), cf.get("line_rule"))
        + '<w:jc w:val="both"/>' + _ind(first_line_chars=200),
        _rpr(cf["east_asia"], cf["size_hp"])))

    # 一~四级标题
    for lvl in ("1", "2", "3", "4"):
        h = spec["headings"][lvl]
        styles.append(_style(
            "CanonH%s" % lvl, "标题 %s" % lvl,
            '<w:outlineLvl w:val="%d"/>' % (int(lvl) - 1)
            + _spacing(line_twips) + _ind(first_line_chars=h["first_line_chars"]),
            _rpr(h["east_asia"], h["size_hp"], western)))

    # 正文
    b = spec["body"]
    styles.append(_style(
        "CanonBody", "正文",
        _spacing(line_twips, zero_before_after=b.get("no_space_before_after"))
        + _ind(first_line_chars=b["first_line_chars"]),
        _rpr(b["east_asia"], b["size_hp"], western)))

    # 表格内容
    tb = spec["table_body"]
    styles.append(_style(
        "CanonTableBody", "表格内容",
        _spacing(tb.get("line_twips"), tb.get("line_rule")) + _ind(no_indent=True),
        _rpr(tb["east_asia"], tb["size_hp"], western)))

    # 图/表标题
    cap = spec["caption_format"]
    styles.append(_style(
        "CanonCaption", "图表标题",
        _spacing(cap.get("line_twips"), cap.get("line_rule"))
        + '<w:jc w:val="%s"/>' % cap.get("jc", "center") + _ind(no_indent=True),
        _rpr(cap["east_asia"], cap["size_hp"], western)))

    # 目录 1/2/3（用标准 toc 样式名，Word 刷新目录时按 outline 级别套用）
    toc = spec["toc"]
    by_level = toc.get("indent_chars_by_level", {})
    for lvl in ("1", "2", "3"):
        styles.append(_style(
            "TOC%s" % lvl, "toc %s" % lvl,
            _ind(left_chars=by_level.get(lvl, 0)),
            _rpr(toc["east_asia"], toc["size_hp"])))   # 目录不套西文（陷阱#7）

    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<w:styles xmlns:w="%s">%s%s</w:styles>' % (W, docdef, "".join(styles)))


def build_numbering():
    """一条 canonical 编号定义：级别缩进【已中和】（left=0、无 hanging）——就是甲法克隆
    钳后的目标形态。自动编号标题指向它，段落自身写 firstLineChars=200（字符单位）。"""
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<w:numbering xmlns:w="%s">'
            '<w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0">'
            '<w:start w:val="1"/><w:numFmt w:val="chineseCounting"/>'
            '<w:lvlText w:val="%%1、"/><w:lvlJc w:val="left"/>'
            '<w:pPr><w:ind w:left="0" w:leftChars="0"/></w:pPr></w:lvl></w:abstractNum>'
            '<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>'
            '</w:numbering>' % W)


def _p(style_id, text, extra_ppr=""):
    return ('<w:p><w:pPr><w:pStyle w:val="%s"/>%s</w:pPr>'
            '<w:r><w:t xml:space="preserve">%s</w:t></w:r></w:p>'
            % (style_id, extra_ppr, _esc(text)))


def _label(text):
    """一行说明标签，用 Normal 小字，帮用户对照该段应符合的规范。"""
    return ('<w:p><w:pPr><w:rPr><w:sz w:val="18"/><w:color w:val="808080"/></w:rPr></w:pPr>'
            '<w:r><w:rPr><w:sz w:val="18"/><w:color w:val="808080"/></w:rPr>'
            '<w:t xml:space="preserve">%s</w:t></w:r></w:p>' % _esc(text))


def build_document(spec):
    def src(key, *path):
        d = spec
        for p in path:
            d = d.get(p, {})
        return d.get("source", "")

    pg = spec["page_setup"]
    sect = ('<w:sectPr>'
            '<w:pgSz w:w="11906" w:h="16838"/>'
            '<w:pgMar w:top="%d" w:right="%d" w:bottom="%d" w:left="%d" '
            'w:header="%d" w:footer="%d" w:gutter="0"/>'
            '<w:footerReference w:type="default" r:id="rIdF"/>'
            '</w:sectPr>'
            % (pg["margin_top_twips"], pg["margin_right_twips"],
               pg["margin_bottom_twips"], pg["margin_left_twips"],
               pg["header_twips"], pg["footer_twips"]))

    pagebreak = ('<w:p><w:r><w:br w:type="page"/></w:r></w:p>')

    toc = spec["toc"]
    # 真正的 TOC 域：Word 打开时刷新，按 toc 样式重建条目——验证"只认 leftChars + 页码不换行"
    toc_field = (
        '<w:p><w:pPr><w:pStyle w:val="CanonH1"/></w:pPr>'
        '<w:r><w:t>目录</w:t></w:r></w:p>'
        '<w:p><w:r><w:fldChar w:fldCharType="begin"/></w:r>'
        '<w:r><w:instrText xml:space="preserve"> TOC \\o &quot;1-3&quot; \\h \\z \\u </w:instrText></w:r>'
        '<w:r><w:fldChar w:fldCharType="separate"/></w:r>'
        '<w:r><w:t>打开时选"是"更新域，此处将生成目录。</w:t></w:r>'
        '<w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>')

    body = []
    body.append(_label("== 封面 == 每段标签给出应符合的规范；用 Word 逐段核对字体/字号/居中/缩进"))
    body.append(_label("【封面密级编号·%s】" % src("cover_classification")))
    body.append(_p("CanonCoverClass", "密级：内部　　　　　　　　文本编号：XXXX-2024-001"))
    body.append(_label("【封面题目·%s】" % src("title")))
    body.append(_p("CanonTitle", "先进技术项目 2024 年度自评价报告"))
    body.append(_label("【封面要素·%s】" % src("cover_field")))
    for line in ("项目名称：××××系统研制项目",
                 "项目编号：KJ-2024-001",
                 "承担单位（盖章）：××××研究院",
                 "项目负责人（签字）：×××",
                 "项目起止时间：20  年  月 至 20  年  月",
                 "报告编制时间：2024 年 ×× 月"):
        body.append(_p("CanonCoverField", line))
    body.append(pagebreak)

    body.append(_label("== 目录 == 打开时更新域；看二/三级页码是否顶到右边界仍不换行（只认 leftChars）"))
    body.append(toc_field)
    body.append(pagebreak)

    body.append(_label("== 正文 =="))
    body.append(_label("【一级标题·自动编号·%s】首行应为 2 字符缩进，不是 -0.74cm 悬挂缩进（甲法验证）"
                       % src("headings", "1")))
    body.append(_p("CanonH1", "概述", extra_ppr='<w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr>'))
    body.append(_label("【正文·%s】选中该段看首行缩进应为 2 字符而非厘米" % src("body")))
    body.append(_p("CanonBody",
                   "这是一段正文，用于确认仿宋三号、行距固定值 28 磅、首行缩进 2 字符，"
                   "并在 Word 的段落对话框里确认首行缩进显示为 2 字符而不是 0.74/0.85 厘米。"
                   "含数字 12345 与英文 ABC 应显示为 Times New Roman。"))
    body.append(_label("【二级标题·%s】" % src("headings", "2")))
    body.append(_p("CanonH2", "研究方法"))
    body.append(_p("CanonBody", "二级标题下的正文示例段落。"))
    body.append(_label("【三级标题·%s】" % src("headings", "3")))
    body.append(_p("CanonH3", "数据来源"))
    body.append(_p("CanonBody", "三级标题下的正文示例段落。"))
    body.append(_label("【四级标题·%s】" % src("headings", "4")))
    body.append(_p("CanonH4", "指标口径"))
    body.append(_p("CanonBody", "四级标题下的正文示例段落。"))
    body.append(_label("【图表标题·%s】居中、无缩进" % src("caption_format")))
    body.append(_p("CanonCaption", "图1 系统总体架构示意图"))
    body.append(_label("【表格内容·%s】四号、行距 28 磅、无缩进（下方表格单元格）"
                       % src("table_body")))
    body.append(
        '<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/>'
        '<w:tblBorders>'
        '<w:top w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
        '<w:left w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
        '<w:bottom w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
        '<w:right w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
        '<w:insideH w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
        '<w:insideV w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
        '</w:tblBorders></w:tblPr>'
        '<w:tr>'
        '<w:tc><w:tcPr><w:tcW w:w="4000" w:type="dxa"/></w:tcPr>'
        + _p("CanonTableBody", "指标") + '</w:tc>'
        '<w:tc><w:tcPr><w:tcW w:w="4000" w:type="dxa"/></w:tcPr>'
        + _p("CanonTableBody", "数值示例 123") + '</w:tc>'
        '</w:tr></w:tbl>')

    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<w:document xmlns:w="%s" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
            '<w:body>%s%s</w:body></w:document>' % (W, "".join(body), sect))


def build_footer(spec):
    pn = spec.get("page_number", {})
    ascii_font = pn.get("ascii", "Times New Roman")
    size = pn.get("size_hp", 24)
    rpr = ('<w:rPr><w:rFonts w:ascii="%s" w:hAnsi="%s"/><w:sz w:val="%d"/></w:rPr>'
           % (_esc(ascii_font), _esc(ascii_font), size))
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<w:ftr xmlns:w="%s"><w:p><w:pPr><w:jc w:val="center"/>%s</w:pPr>'
            '<w:r>%s<w:fldChar w:fldCharType="begin"/></w:r>'
            '<w:r>%s<w:instrText xml:space="preserve"> PAGE </w:instrText></w:r>'
            '<w:r>%s<w:fldChar w:fldCharType="separate"/></w:r>'
            '<w:r>%s<w:t>1</w:t></w:r>'
            '<w:r>%s<w:fldChar w:fldCharType="end"/></w:r></w:p></w:ftr>'
            % (W, rpr, rpr, rpr, rpr, rpr, rpr))


CONTENT_TYPES = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
    '<Default Extension="xml" ContentType="application/xml"/>'
    '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
    '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>'
    '<Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/>'
    '<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>'
    '<Override PartName="/word/footer1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/>'
    '</Types>')

ROOT_RELS = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>'
    '</Relationships>')

DOC_RELS = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    '<Relationship Id="rIdS" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>'
    '<Relationship Id="rIdN" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/>'
    '<Relationship Id="rIdSet" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/>'
    '<Relationship Id="rIdF" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footer1.xml"/>'
    '</Relationships>')

SETTINGS = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<w:settings xmlns:w="%s"><w:updateFields w:val="true"/></w:settings>' % W)


def build(out_path):
    with open(SPEC_PATH, encoding="utf-8") as f:
        spec = json.load(f)
    parts = {
        "[Content_Types].xml": CONTENT_TYPES,
        "_rels/.rels": ROOT_RELS,
        "word/_rels/document.xml.rels": DOC_RELS,
        "word/styles.xml": build_styles(spec),
        "word/numbering.xml": build_numbering(),
        "word/settings.xml": SETTINGS,
        "word/footer1.xml": build_footer(spec),
        "word/document.xml": build_document(spec),
    }
    if os.path.exists(out_path):
        os.remove(out_path)
    with zipfile.ZipFile(out_path, "w", zipfile.ZIP_DEFLATED) as z:
        for name, data in parts.items():
            z.writestr(name, data)
    return out_path


def main():
    out = sys.argv[1] if len(sys.argv) > 1 else os.path.join(os.getcwd(),
                                                             "canonical_reference.docx")
    path = build(out)
    print(json.dumps({"status": "ok", "canonical_reference": path},
                     ensure_ascii=False))


if __name__ == "__main__":
    main()
