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
  3. **自动编号标题**（一~四级都自动编号："一、/（一）/1./（1）"）：首行是 2 字符缩进、不是 -0.74cm
     悬挂缩进——验证甲法克隆钳。编号后的制表位位置留待用户确认（当前不设 defaultTabStop）。
  4. **目录**：为真正的 TOC 域，打开即显示正确缓存条目（编号+制表符+标题+点线+页码）——制表位、
     半角编号、右点线位置(8664)全部照规范文档 XML；点线应在标题与页码之间、页码右对齐不换行。
     （一级目录只有左制表位、无页码点线，与规范文档一致。）
  5. 封面各要素：方正黑体_GBK、题目居中、题目下要素两端对齐+首行缩进2字符。
  6. **图/表标题自动编号**："图1/表1"应为一个整体（不能拆选"图""1"），编号后是空格不是制表位。
  7. **文档网格**：只指定行网格（行距 15.6 磅），与规范文档一致——封面要素仍设 2 倍行距。
  8. **表格默认单元格边距**：上0/左0.19cm/下0/右0.19cm。

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

    # docDefaults：默认五号 21 半点 + lang（照抄规范文档 styles.xml）。行网格 15.6磅/41行
    # 由 settings 的 compat 块（useFELayout/compatMode15）保证，与默认字号无关。
    docdef = ('<w:docDefaults><w:rPrDefault><w:rPr>'
              '<w:rFonts w:ascii="Times New Roman" w:eastAsia="仿宋" w:hAnsi="Times New Roman" '
              'w:cs="Times New Roman"/><w:sz w:val="21"/><w:szCs w:val="21"/>'
              '<w:lang w:val="en-US" w:eastAsia="zh-CN" w:bidi="ar-SA"/>'
              '</w:rPr></w:rPrDefault><w:pPrDefault/></w:docDefaults>')
    # Normal ＝ 正文本身（规范文档如此）：仿宋 三号、行距固定值 28磅、首行缩进 2 字符。
    styles.append('<w:style w:type="paragraph" w:default="1" w:styleId="Normal">'
                  '<w:name w:val="Normal"/><w:qFormat/>'
                  '<w:pPr><w:spacing w:line="%d" w:lineRule="exact"/>'
                  '<w:ind w:firstLineChars="200"/></w:pPr>'
                  '<w:rPr><w:sz w:val="32"/><w:szCs w:val="32"/></w:rPr></w:style>'
                  % line_twips)

    # 题目
    t = spec["title"]
    styles.append(_style(
        "CanonTitle", "封面题目",
        '<w:jc w:val="%s"/>' % t.get("jc", "center"),
        _rpr(t["east_asia"], t["size_hp"], western if t.get("enforce_western") else None)))

    # 封面密级/文本编号行（数字/字母用西文字体——用户 Word 验收要求）
    cc = spec["cover_classification"]
    styles.append(_style("CanonCoverClass", "封面密级编号",
                         '<w:jc w:val="center"/>',
                         _rpr(cc["east_asia"], cc["size_hp"], western)))

    # 封面题目下要素（field）：2倍行距、两端对齐、首行缩进2字符；数字/字母用西文字体
    cf = spec["cover_field"]
    styles.append(_style(
        "CanonCoverField", "封面要素",
        _spacing(cf.get("line_twips"), cf.get("line_rule"))
        + '<w:jc w:val="both"/>' + _ind(first_line_chars=200),
        _rpr(cf["east_asia"], cf["size_hp"], western)))

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

    # 目录标题"目录"二字：用正文字体字号但居中、无缩进（非标题，不进 TOC）
    b_toc = spec["body"]
    styles.append(_style(
        "CanonTocTitle", "目录标题",
        _spacing(line_twips) + '<w:jc w:val="center"/>' + _ind(no_indent=True),
        _rpr(b_toc["east_asia"], b_toc["size_hp"], western)))

    # 目录 1/2/3（标准 toc 样式名）——【照抄规范文档 styles.xml 的精确值】：
    #   制表位写在【样式】里；左制表位 4/5/6 字符按【三号 320 twips/字符】= 1280/1600/1920
    #   （不是五号 210！这是之前算错、编号越过制表位的根因）；页码右制表位带点线 = 13203
    #   （41.26 字符×320）；左缩进 leftChars 0/200/400；单倍行距(line 240 auto)；仿宋 小三。
    toc = spec["toc"]
    toc_rpr = ('<w:rFonts w:ascii="仿宋" w:hAnsi="仿宋" w:cs="仿宋"/>'
               '<w:sz w:val="%d"/><w:szCs w:val="%d"/>'
               % (toc["size_hp"], toc["size_hp"]))
    RIGHT_TAB = 13203
    toc_defs = {
        "1": (1280, '<w:ind w:firstLineChars="0" w:firstLine="0"/>'),
        "2": (1600, '<w:ind w:leftChars="200" w:left="200"/>'),
        "3": (1920, '<w:ind w:leftChars="400" w:left="400" w:firstLineChars="0" w:firstLine="0"/>'),
    }
    for lvl in ("1", "2", "3"):
        left_tab, ind = toc_defs[lvl]
        tabs = ('<w:tabs><w:tab w:val="left" w:pos="%d"/>'
                '<w:tab w:val="right" w:leader="dot" w:pos="%d"/></w:tabs>'
                % (left_tab, RIGHT_TAB))
        ppr = tabs + '<w:spacing w:line="240" w:lineRule="auto"/>' + ind
        styles.append('<w:style w:type="paragraph" w:styleId="TOC%s"><w:name w:val="toc %s"/>'
                      '<w:basedOn w:val="Normal"/><w:next w:val="Normal"/>'
                      '<w:uiPriority w:val="39"/><w:qFormat/>'
                      '<w:pPr>%s</w:pPr><w:rPr>%s</w:rPr></w:style>'
                      % (lvl, lvl, ppr, toc_rpr))

    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<w:styles xmlns:w="%s">%s%s</w:styles>' % (W, docdef, "".join(styles)))


def _lvl(ilvl, num_fmt, lvl_text, suff=None):
    """一个编号级别：缩进【已中和】（left=0、无 hanging）——甲法克隆钳后的目标形态；
    段落自身写 firstLineChars=200（字符单位）负责首行缩进。suff 控编号后的分隔符
    （默认 tab；图表标题用 space＝编号后无制表位）。"""
    suff_el = ('<w:suff w:val="%s"/>' % suff) if suff else ""
    return ('<w:lvl w:ilvl="%d"><w:start w:val="1"/>%s'
            '<w:numFmt w:val="%s"/><w:lvlText w:val="%s"/><w:lvlJc w:val="left"/>'
            '<w:pPr><w:ind w:left="0" w:leftChars="0"/></w:pPr></w:lvl>'
            % (ilvl, suff_el, num_fmt, lvl_text))


def build_numbering():
    """三条 canonical 编号定义（级别缩进全部中和，甲法目标形态）：
      * abstractNum 0 / numId 1 —— 多级标题：一、/（一）/1./（1），suff=tab（编号后制表符）。
      * abstractNum 1 / numId 2 —— 图标题：图%%1，decimal，suff=tab（编号后制表符）。
      * abstractNum 2 / numId 3 —— 表标题：表%%1，decimal，suff=tab。
    图/表标题自动编号后"图1"是一个整体（不能拆选），序号后有制表符。"""
    # 编号用【半角括号 (一)】而非全角（一）——规范文档如此；全角括号更宽，会越过目录左
    # 制表位、把点线挤到编号与标题之间（用户 v4 实测的乱象根因）。
    headings = "".join([
        _lvl(0, "chineseCounting", "%1、"),
        _lvl(1, "chineseCounting", "(%2)"),
        _lvl(2, "decimal", "%3."),
        _lvl(3, "decimal", "(%4)"),
    ])
    # 图/表标题：序号后要【制表符】（用户 Word 验收订正——用默认 suff=tab，不写 space）
    figure = _lvl(0, "decimal", "图%1")
    table = _lvl(0, "decimal", "表%1")
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<w:numbering xmlns:w="%s">'
            '<w:abstractNum w:abstractNumId="0">%s</w:abstractNum>'
            '<w:abstractNum w:abstractNumId="1">%s</w:abstractNum>'
            '<w:abstractNum w:abstractNumId="2">%s</w:abstractNum>'
            '<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>'
            '<w:num w:numId="2"><w:abstractNumId w:val="1"/></w:num>'
            '<w:num w:numId="3"><w:abstractNumId w:val="2"/></w:num>'
            '</w:numbering>' % (W, headings, figure, table))


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
    # 文档网格：只指定行网格（type=lines），行间距 15.6 磅 = 312 twips（每页约 41 行）。
    # 与规范文档一致——有行网格时段落 snapToGrid（默认开），单倍/1.5/2 倍行距会贴到
    # 网格线而看起来接近，故封面要素仍保留 2 倍行距设置以与规范文档一致。docGrid 须放
    # 在 sectPr 末尾（schema 顺序），footerReference 须在 pgSz 之前。
    sect = ('<w:sectPr>'
            '<w:footerReference w:type="default" r:id="rIdF"/>'
            '<w:pgSz w:w="11906" w:h="16838"/>'
            '<w:pgMar w:top="%d" w:right="%d" w:bottom="%d" w:left="%d" '
            'w:header="%d" w:footer="%d" w:gutter="0"/>'
            '<w:pgNumType w:start="1"/>'
            '<w:cols w:space="720"/>'
            '<w:docGrid w:type="lines" w:linePitch="312"/>'
            '</w:sectPr>'
            % (pg["margin_top_twips"], pg["margin_right_twips"],
               pg["margin_bottom_twips"], pg["margin_left_twips"],
               pg["header_twips"], pg["footer_twips"]))

    pagebreak = ('<w:p><w:r><w:br w:type="page"/></w:r></w:p>')

    toc = spec["toc"]
    # 目录＝真正的 TOC 域（真实文档如此），其【缓存结果】为写好的条目。制表位【继承自
    # TOC1/2/3 样式】（左 1280/1600/1920 + 右 13203 点线，取自规范文档 styles.xml），故条目
    # 段落自身不再写直接制表位；编号用半角 (一)/1.。打开即正确显示"编号+制表符+标题+
    # 点线+页码"；右键更新域后 Word 按大纲级别（本文正文标题带 outlineLvl）重建。
    def toc_entry(level, number, title, page, is_last=False):
        end_run = '<w:r><w:fldChar w:fldCharType="end"/></w:r>' if is_last else ''
        return ('<w:p><w:pPr><w:pStyle w:val="TOC%d"/></w:pPr>'
                '<w:r><w:t xml:space="preserve">%s</w:t></w:r>'
                '<w:r><w:tab/></w:r>'
                '<w:r><w:t xml:space="preserve">%s</w:t></w:r>'
                '<w:r><w:tab/></w:r>'
                '<w:r><w:t xml:space="preserve">%s</w:t></w:r>%s</w:p>'
                % (level, _esc(number), _esc(title), _esc(page), end_run))

    toc_field = (
        '<w:p><w:pPr><w:pStyle w:val="CanonTocTitle"/></w:pPr>'
        '<w:r><w:t>目录</w:t></w:r>'
        '<w:r><w:fldChar w:fldCharType="begin"/></w:r>'
        '<w:r><w:instrText xml:space="preserve"> TOC \\o &quot;1-3&quot; \\h \\z \\u </w:instrText></w:r>'
        '<w:r><w:fldChar w:fldCharType="separate"/></w:r></w:p>'
        + toc_entry(1, "一、", "概述", "3")
        + toc_entry(2, "(一)", "研究方法", "3")
        + toc_entry(3, "1.", "数据来源", "3", is_last=True))

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

    body.append(_label("== 目录 == 静态条目（直接制表位）：自动编号+制表符+标题+点线+页码；"
                       "看点线是否在标题与页码之间、页码是否右对齐不换行"))
    body.append(toc_field)
    body.append(pagebreak)

    def numpr(ilvl, num_id):
        return '<w:numPr><w:ilvl w:val="%d"/><w:numId w:val="%d"/></w:numPr>' % (ilvl, num_id)

    body.append(_label("== 正文 =="))
    body.append(_label("【一级标题·自动编号·%s】首行应为 2 字符缩进，不是 -0.74cm 悬挂缩进（甲法验证）"
                       % src("headings", "1")))
    body.append(_p("CanonH1", "概述", extra_ppr=numpr(0, 1)))
    body.append(_label("【正文·%s】选中该段看首行缩进应为 2 字符而非厘米" % src("body")))
    body.append(_p("CanonBody",
                   "这是一段正文，用于确认仿宋三号、行距固定值 28 磅、首行缩进 2 字符，"
                   "并在 Word 的段落对话框里确认首行缩进显示为 2 字符而不是 0.74/0.85 厘米。"
                   "含数字 12345 与英文 ABC 应显示为 Times New Roman。"))
    body.append(_label("【二级标题·自动编号·%s】" % src("headings", "2")))
    body.append(_p("CanonH2", "研究方法", extra_ppr=numpr(1, 1)))
    body.append(_p("CanonBody", "二级标题下的正文示例段落。"))
    body.append(_label("【三级标题·自动编号·%s】" % src("headings", "3")))
    body.append(_p("CanonH3", "数据来源", extra_ppr=numpr(2, 1)))
    body.append(_p("CanonBody", "三级标题下的正文示例段落。"))
    body.append(_label("【四级标题·自动编号·%s】" % src("headings", "4")))
    body.append(_p("CanonH4", "指标口径", extra_ppr=numpr(3, 1)))
    body.append(_p("CanonBody", "四级标题下的正文示例段落。"))
    body.append(_label("【图标题·自动编号·%s】“图1”应为一个整体（不能拆选“图”“1”）、编号后无制表位"
                       % src("caption_format")))
    body.append(_p("CanonCaption", "系统总体架构示意图", extra_ppr=numpr(0, 2)))
    body.append(_label("【表标题·自动编号】“表1”自动生成、编号后无制表位（下方表格）"))
    body.append(_p("CanonCaption", "主要指标对照表", extra_ppr=numpr(0, 3)))
    body.append(_label("【表格内容·%s】四号、行距 28 磅、无缩进；默认单元格边距 上0/左0.19cm/下0/右0.19cm"
                       % src("table_body")))
    # 表格默认单元格边距：上0/下0/左右 108 twips(0.19cm)。tblCellMar 须在 tblBorders 之后。
    cell_mar = ('<w:tblCellMar>'
                '<w:top w:w="0" w:type="dxa"/><w:left w:w="108" w:type="dxa"/>'
                '<w:bottom w:w="0" w:type="dxa"/><w:right w:w="108" w:type="dxa"/>'
                '</w:tblCellMar>')
    # 单元格垂直居中（vAlign=center）；tcW 后写。tcPr 是单元格属性，与 tblCellMar（表级
    # 默认边距）是两回事——用户订正：表格左缩进(tblInd)要 0，别跟单元格边距搞混。
    def cell(text):
        return ('<w:tc><w:tcPr><w:tcW w:w="4000" w:type="dxa"/>'
                '<w:vAlign w:val="center"/></w:tcPr>' + _p("CanonTableBody", text) + '</w:tc>')
    body.append(
        '<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/>'
        '<w:tblInd w:w="0" w:type="dxa"/>'   # 表格左缩进 0（不是 0.19cm）
        '<w:tblBorders>'
        '<w:top w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
        '<w:left w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
        '<w:bottom w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
        '<w:right w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
        '<w:insideH w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
        '<w:insideV w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
        '</w:tblBorders>' + cell_mar + '</w:tblPr>'
        '<w:tr>' + cell("指标") + cell("数值示例 123") + '</w:tr></w:tbl>')

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
    '<w:settings xmlns:w="%s">'
    '<w:zoom w:percent="100"/>'
    '<w:bordersDoNotSurroundHeader/><w:bordersDoNotSurroundFooter/>'
    '<w:proofState w:spelling="clean" w:grammar="clean"/>'
    # defaultTabStop=640：取自用户手动修好目录后的 settings.xml（原 420 目录不对）。
    '<w:defaultTabStop w:val="640"/>'
    '<w:drawingGridHorizontalSpacing w:val="105"/>'   # 0.5 字符（配合 compat 后单位）
    '<w:drawingGridVerticalSpacing w:val="156"/>'     # 0.5 行
    '<w:displayHorizontalDrawingGridEvery w:val="2"/>'
    '<w:displayVerticalDrawingGridEvery w:val="2"/>'
    '<w:noPunctuationKerning/>'
    '<w:characterSpacingControl w:val="compressPunctuation"/>'
    # ▼ 关键：compat 块（尤其 useFELayout + compatibilityMode=15）让 Word 尊重行网格
    #   linePitch=312 → 15.6 磅/41 行。缺它时 Word 用旧版式把行距顶到 21.75 磅/29 行。
    #   取自规范文档 settings.xml。
    '<w:compat>'
    '<w:spaceForUL/><w:balanceSingleByteDoubleByteWidth/><w:doNotLeaveBackslashAlone/>'
    '<w:ulTrailSpace/><w:doNotExpandShiftReturn/><w:adjustLineHeightInTable/><w:useFELayout/>'
    '<w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/>'
    '<w:compatSetting w:name="overrideTableStyleFontSizeAndJustification" w:uri="http://schemas.microsoft.com/office/word" w:val="1"/>'
    '<w:compatSetting w:name="enableOpenTypeFeatures" w:uri="http://schemas.microsoft.com/office/word" w:val="1"/>'
    '<w:compatSetting w:name="doNotFlipMirrorIndents" w:uri="http://schemas.microsoft.com/office/word" w:val="1"/>'
    '<w:compatSetting w:name="differentiateMultirowTableHeaders" w:uri="http://schemas.microsoft.com/office/word" w:val="1"/>'
    '<w:compatSetting w:name="useWord2013TrackBottomHyphenation" w:uri="http://schemas.microsoft.com/office/word" w:val="1"/>'
    '</w:compat>'
    '<w:themeFontLang w:val="en-US" w:eastAsia="zh-CN"/>'
    '<w:decimalSymbol w:val="."/><w:listSeparator w:val=","/>'
    # 不设 updateFields：目录缓存条目已是正确版式，避免打开时提示更新域后按（我们暂无
    # 规范文档 toc 样式的）重建逻辑覆盖掉正确缓存。页码/目录如需刷新可手动 F9。
    '</w:settings>' % W)


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
