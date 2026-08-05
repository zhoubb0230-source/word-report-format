# -*- coding: utf-8 -*-
"""生成方案C 的【canonical 参考 .docx】——阶段0 人肉验收件（handoff §5）。

沙箱无 Word/soffice，方案C 的真正判据是 Word 渲染，只能靠用户实测（handoff §1）。
本脚本按 `spec/format_spec.json` 造一份"每角色一段、套注入 canonical 样式"的 docx，
供用户在 **Word** 里打开、按下面清单逐项确认。这是整套方案唯一的人肉环节。

用法：
    python3 scripts/make_canonical_reference.py [输出路径]
    # 默认写到 cwd 的 canonical_reference.docx

验收清单（在 Word 里逐条看）：
  1. 每个角色的字体/字号/居中/行距是否与规范一致（每段前有【角色·规范】标签）；一~四级标题**加粗**。
  2. **首行缩进显示为"2 字符"而非厘米**——选中正文/标题段，看"段落→缩进→首行缩进 2 字符"，
     不是 0.74cm/0.85cm 之类。这是严格-spec §2.2 的关键判据（canonical 只写字符单位、无绝对伴随值）。
  3. **自动编号标题**（一~四级都自动编号："一、/(一)/1./（1）"）：首行是 2 字符缩进、不是 -0.74cm
     悬挂缩进——验证甲法克隆钳。编号后的制表位位置留待用户确认（当前不设 defaultTabStop）。
     **序号字体**：一级黑体、二级的半角括号是 Times New Roman（括号里的中文数字是楷体）、
     四级整个"（1）"是仿宋——点在序号上看 Word 的字体框。
  4. **目录**：为真正的 TOC 域，打开即显示正确缓存条目（编号+制表符+标题+点线+页码）——制表位、
     半角编号、右点线位置(8664)全部照规范文档 XML；点线应在标题与页码之间、页码右对齐不换行。
     （一级目录只有左制表位、无页码点线，与规范文档一致。）
  5. 封面各要素：方正黑体_GBK、题目居中、题目下要素两端对齐+首行缩进2字符。
  6. **图/表标题自动编号**："图1/表1"应为一个整体（不能拆选"图""1"），编号后是**制表符**；
     文字里不再有静态的"图1/表1"，编号全部由 Word 生成（插删图表后自动重排）。
  7. **文档网格**：只指定行网格（行距 15.6 磅），与规范文档一致——封面要素仍设 2 倍行距。
  8. **表格默认单元格边距**：上0/左0.19cm/下0/右0.19cm。

判据只来自 spec（红线）：本脚本不写死任何字体/字号/缩进，全部读 `spec/format_spec.json`。
纯 stdlib（zipfile）+ 少量字符串拼装，不依赖 lxml，也不进判定流水线。
"""
import json
import os
import sys
import zipfile

sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), "lib"))

import canonstyles
from canonstyles import STYLE_ID_BY_ROLE, _esc

SPEC_PATH = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))),
                         "spec", "format_spec.json")

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

# 文档字符单位字号（Normal/docDefaults 的五号）——目录制表位按它把"字符"换算成
# twips（陷阱 #12：字符单位随 Normal 走）。
CHAR_UNIT_HP = 21


def build_styles(spec):
    """按 spec 造全角色 canonical 样式，返回 styles.xml 字符串。

    样式本身**不在这里定义**——权威定义在 `scripts/lib/canonstyles.py`，与
    `40_apply_fixes.py` 注入进真实文档的是同一份字符串（避免"参考件验收通过、
    流水线注入的却是另一套"的漂移）。本函数只负责把它们拼成一个 styles.xml。"""
    parts = [canonstyles.doc_defaults_xml(spec), canonstyles.normal_style_xml(spec)]
    parts += [d["xml"] for d in canonstyles.canonical_style_defs(spec)]
    # 目录 1/2/3（标准 toc 样式名）：仿宋 小三、leftChars 0/200/400、左制表位
    # 4/5/6 字符 + 右点线制表位 41.26 字符（Normal=五号 → 840/1050/1260 + 8665）。
    parts += [canonstyles.toc_style_xml(spec, lvl, CHAR_UNIT_HP) for lvl in ("1", "2", "3")]
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<w:styles xmlns:w="%s">%s</w:styles>' % (W, "".join(parts)))


def build_numbering(spec):
    """三条 canonical 编号定义（级别缩进全部中和，甲法目标形态）：
      * abstractNum 0 / numId 1 —— 多级标题：一、/(一)/1./（1），suff=tab（编号后制表符），
        每级带自己的序号字体（一级黑体、二级半角括号走 Times、四级仿宋）。
      * abstractNum 1 / numId 2 —— 图标题：图%%1，decimal，suff=tab。
      * abstractNum 2 / numId 3 —— 表标题：表%%1，decimal，suff=tab。

    三条**全部直接取自 `canonstyles`**（标题 `heading_numbering_def`、图表
    `caption_numbering_defs`）——`40_apply_fixes.py` 往真实文档里注入的是同一份字符串，
    参考件与流水线不会各写一套。标题那条曾经在这里手写过一遍，结果两边的括号形状
    （半角 vs 全角）就漂了，用户当轮报"二级标题的括号不是英文括号"；漂移是这个项目最
    容易犯的错，别再把它抄回来。自动编号后"图1"是一个整体（不能拆选），编号后是制表符。"""
    headings = canonstyles.heading_numbering_def(spec)["lvl_xml"]
    caps = {d["kind"]: d["lvl_xml"] for d in canonstyles.caption_numbering_defs(spec)}
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<w:numbering xmlns:w="%s">'
            '<w:abstractNum w:abstractNumId="0">%s</w:abstractNum>'
            '<w:abstractNum w:abstractNumId="1">%s</w:abstractNum>'
            '<w:abstractNum w:abstractNumId="2">%s</w:abstractNum>'
            '<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>'
            '<w:num w:numId="2"><w:abstractNumId w:val="1"/></w:num>'
            '<w:num w:numId="3"><w:abstractNumId w:val="2"/></w:num>'
            '</w:numbering>'
            % (W, headings, caps.get("figure", ""), caps.get("table", "")))


def _p(style_id, text, extra_ppr=""):
    return ('<w:p><w:pPr><w:pStyle w:val="%s"/>%s</w:pPr>'
            '<w:r><w:t xml:space="preserve">%s</w:t></w:r></w:p>'
            % (style_id, extra_ppr, _esc(text)))


def _label(text):
    """一行说明标签，用 Normal 小字，帮用户对照该段应符合的规范。"""
    # rPr 的子元素顺序照 CT_RPr 的 sequence：color(在 spacing 前) 排在 sz 之前。
    # 顺序写反文件仍良构，但 Word 会提示"发现无法读取的内容"。
    rpr = '<w:rPr><w:color w:val="808080"/><w:sz w:val="18"/></w:rPr>'
    return ('<w:p><w:pPr>%s</w:pPr><w:r>%s'
            '<w:t xml:space="preserve">%s</w:t></w:r></w:p>' % (rpr, rpr, _esc(text)))


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
    grid = canonstyles.doc_grid_attrs(spec)
    docgrid = ('<w:docGrid w:type="%s" w:linePitch="%s"/>'
               % (grid["type"], grid["linePitch"])) if grid else ""
    sect = ('<w:sectPr>'
            '<w:footerReference w:type="default" r:id="rIdF"/>'
            '<w:pgSz w:w="11906" w:h="16838"/>'
            '<w:pgMar w:top="%d" w:right="%d" w:bottom="%d" w:left="%d" '
            'w:header="%d" w:footer="%d" w:gutter="0"/>'
            '<w:pgNumType w:start="1"/>'
            '<w:cols w:space="720"/>'
            '%s'
            '</w:sectPr>'
            % (pg["margin_top_twips"], pg["margin_right_twips"],
               pg["margin_bottom_twips"], pg["margin_left_twips"],
               pg["header_twips"], pg["footer_twips"], docgrid))

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
        '<w:p><w:pPr><w:pStyle w:val="%s"/></w:pPr>' % STYLE_ID_BY_ROLE["toc_title"]
        + '<w:r><w:t>目录</w:t></w:r>'
        '<w:r><w:fldChar w:fldCharType="begin"/></w:r>'
        '<w:r><w:instrText xml:space="preserve"> TOC \\o &quot;1-3&quot; \\h \\z \\u </w:instrText></w:r>'
        '<w:r><w:fldChar w:fldCharType="separate"/></w:r></w:p>'
        + toc_entry(1, "一、", "概述", "3")
        + toc_entry(2, "(一)", "研究方法", "3")
        + toc_entry(3, "1.", "数据来源", "3", is_last=True))

    body = []
    body.append(_label("== 封面 == 每段标签给出应符合的规范；用 Word 逐段核对字体/字号/居中/缩进"))
    body.append(_label("【封面密级编号·%s】" % src("cover_classification")))
    # 居中写成【段落直接属性】而非样式属性：规范对密级/文本编号行的**对齐方式没有
    # 规定**（checks 也只校验字体字号），故 canonical 样式不承载 jc——否则注入真实
    # 文档时会顺手改掉模板自己的版式。这里居中只是参考件的排版示范。
    body.append(_p(STYLE_ID_BY_ROLE["cover_classification"],
                   "密级：内部　　　　　　　　文本编号：XXXX-2024-001",
                   extra_ppr='<w:jc w:val="center"/>'))
    body.append(_label("【封面题目·%s】" % src("title")))
    body.append(_p(STYLE_ID_BY_ROLE["title"], "先进技术项目 2024 年度自评价报告"))
    body.append(_label("【封面要素·%s】" % src("cover_field")))
    for line in ("项目名称：××××系统研制项目",
                 "项目编号：KJ-2024-001",
                 "承担单位（盖章）：××××研究院",
                 "项目负责人（签字）：×××",
                 "项目起止时间：20  年  月 至 20  年  月",
                 "报告编制时间：2024 年 ×× 月"):
        body.append(_p(STYLE_ID_BY_ROLE["cover_field"], line))
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
    body.append(_p(STYLE_ID_BY_ROLE["heading1"], "概述", extra_ppr=numpr(0, 1)))
    body.append(_label("【正文·%s】选中该段看首行缩进应为 2 字符而非厘米" % src("body")))
    body.append(_p(STYLE_ID_BY_ROLE["body"],
                   "这是一段正文，用于确认仿宋三号、行距固定值 28 磅、首行缩进 2 字符，"
                   "并在 Word 的段落对话框里确认首行缩进显示为 2 字符而不是 0.74/0.85 厘米。"
                   "含数字 12345 与英文 ABC 应显示为 Times New Roman。"))
    body.append(_label("【二级标题·自动编号·%s】" % src("headings", "2")))
    body.append(_p(STYLE_ID_BY_ROLE["heading2"], "研究方法", extra_ppr=numpr(1, 1)))
    body.append(_p(STYLE_ID_BY_ROLE["body"], "二级标题下的正文示例段落。"))
    body.append(_label("【三级标题·自动编号·%s】" % src("headings", "3")))
    body.append(_p(STYLE_ID_BY_ROLE["heading3"], "数据来源", extra_ppr=numpr(2, 1)))
    body.append(_p(STYLE_ID_BY_ROLE["body"], "三级标题下的正文示例段落。"))
    body.append(_label("【四级标题·自动编号·%s】" % src("headings", "4")))
    body.append(_p(STYLE_ID_BY_ROLE["heading4"], "指标口径", extra_ppr=numpr(3, 1)))
    body.append(_p(STYLE_ID_BY_ROLE["body"], "四级标题下的正文示例段落。"))
    body.append(_label("【图标题·自动编号·%s】“图1”应为一个整体（不能拆选“图”“1”）、编号后无制表位"
                       % src("caption_format")))
    body.append(_p(STYLE_ID_BY_ROLE["caption_figure"], "系统总体架构示意图", extra_ppr=numpr(0, 2)))
    body.append(_label("【表标题·自动编号】“表1”自动生成、编号后无制表位（下方表格）"))
    body.append(_p(STYLE_ID_BY_ROLE["caption_table"], "主要指标对照表", extra_ppr=numpr(0, 3)))
    body.append(_label("【表格内容·%s】四号、行距 28 磅、无缩进；默认单元格边距 上0/左0.19cm/下0/右0.19cm"
                       % src("table_body")))
    # 表级默认单元格边距（spec.table_defaults）。tblCellMar 须在 tblBorders 之后。
    cell_mar = canonstyles.cell_margins_xml(spec)
    # 单元格垂直居中（vAlign）；tcW 后写。tcPr 是单元格属性，与 tblCellMar（表级
    # 默认边距）是两回事——用户订正：表格左缩进(tblInd)要 0，别跟单元格边距搞混。
    def cell(text):
        return ('<w:tc><w:tcPr><w:tcW w:w="4000" w:type="dxa"/>'
                + canonstyles.cell_valign_xml(spec) + '</w:tcPr>'
                + _p(STYLE_ID_BY_ROLE["table_body"], text) + '</w:tc>')
    body.append(
        '<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/>'
        + canonstyles.table_ind_xml(spec) +
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

def build_settings(spec):
    """settings.xml。承载【文档网格】的那几个子元素（defaultTabStop / 绘图网格 /
    compat 块 / themeFontLang）由 `canonstyles.settings_children_xml` 从 spec 生成
    ——40 往真实文档里写的是同一份，避免参考件与流水线漂移。

    不设 updateFields：目录缓存条目已是正确版式，避免打开时提示更新域后覆盖掉正确
    缓存。页码/目录如需刷新可手动 F9。"""
    grid = "".join(xml for _tag, xml in canonstyles.settings_children_xml(spec))
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<w:settings xmlns:w="%s">'
            '<w:zoom w:percent="100"/>'
            '<w:bordersDoNotSurroundHeader/><w:bordersDoNotSurroundFooter/>'
            '<w:proofState w:spelling="clean" w:grammar="clean"/>'
            '%s'
            '<w:decimalSymbol w:val="."/><w:listSeparator w:val=","/>'
            '</w:settings>' % (W, grid))


def build(out_path):
    with open(SPEC_PATH, encoding="utf-8") as f:
        spec = json.load(f)
    parts = {
        "[Content_Types].xml": CONTENT_TYPES,
        "_rels/.rels": ROOT_RELS,
        "word/_rels/document.xml.rels": DOC_RELS,
        "word/styles.xml": build_styles(spec),
        "word/numbering.xml": build_numbering(spec),
        "word/settings.xml": build_settings(spec),
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
