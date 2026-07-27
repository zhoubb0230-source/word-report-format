# -*- coding: utf-8 -*-
"""canonstyles.py —— 方案C 阶段2 的 **canonical 样式唯一渲染源**。

方案C 的核心动作是把"规范格式值"从段落直接属性搬进**注入的命名样式**里
（`HANDOFF_方案C实施.md` §2/§4）。这些样式的精确参数**已由用户在 Word 里逐项验收**
（阶段0 参考件，见 `已知陷阱与设计决策.md` #12），所以它们只能有**一份**定义：

  * `scripts/make_canonical_reference.py` —— 生成人肉验收用的参考 .docx；
  * `scripts/40_apply_fixes.py`           —— 往真实文档里注入同一批样式并指派。

历史教训（CLAUDE.md「区域划分是共享逻辑」）：同一份规则各写一遍必然漂移。故本模块
把样式**渲染成 XML 字符串**——参考件直接拼进 styles.xml，40 用 lxml 解析同一字符串后
插入/替换。两条路径共用同一份字节，结构上不可能漂移。

**纯 stdlib**（不 import lxml）：参考件生成器不依赖 lxml，本模块必须同样零依赖。
判定值全部读 `spec/format_spec.json`（红线：判定值只来自 spec），本模块不写死任何
字体/字号/缩进数值——只有"哪个角色由哪些属性构成"的结构性知识。
"""

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

# 角色 -> (styleId, 样式名)。角色名与 `checks.paragraph_role()` 的返回值一一对应，
# 样式指派靠它对表。
#
# ⚠️ 正文样式名必须是 "FGW正文" 而**不能**是 "正文"（陷阱 #12）：Word 把【文档网格
# 字体】链接到名为"正文"(=Normal) 的样式，正文内容样式若也叫"正文"会抢占该链接、
# 把网格字体拽成三号，行网格 15.6磅/41行 就退回 21.75磅/29行。
ROLE_STYLES = (
    ("title",                "FGWCanonTitle",      "封面题目"),
    ("cover_classification", "FGWCanonCoverClass", "封面密级编号"),
    ("cover_field",          "FGWCanonCoverField", "封面要素"),
    ("heading1",             "FGWCanonH1",         "标题 1"),
    ("heading2",             "FGWCanonH2",         "标题 2"),
    ("heading3",             "FGWCanonH3",         "标题 3"),
    ("heading4",             "FGWCanonH4",         "标题 4"),
    ("body",                 "FGWCanonBody",       "FGW正文"),
    ("caption",              "FGWCanonCaption",    "图表标题"),
    ("table_body",           "FGWCanonTableBody",  "表格内容"),
    ("toc_title",            "FGWCanonTocTitle",   "目录标题"),
)

STYLE_ID_BY_ROLE = {role: sid for role, sid, _n in ROLE_STYLES}
CANONICAL_STYLE_IDS = frozenset(sid for _r, sid, _n in ROLE_STYLES)

# Word/LibreOffice 里 Normal 样式的 styleId（文档网格字体的链接对象）。
NORMAL_STYLE_ID = "Normal"


# ---------------------------------------------------------------------------
# 低层拼装
# ---------------------------------------------------------------------------
def _esc(s):
    return (str(s).replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
            .replace('"', "&quot;"))


def char_twips(chars, size_hp):
    """字符单位缩进（百分之一字符，200=2字符）→ 该字号下的绝对宽度 twips。

    一个中日韩字符宽 1 em，故 twips = size_hp/2 磅 × 20 twips/磅 = size_hp×10。
    ``size_hp`` 必须是**文档字符单位字号**（docDefaults/Normal 的字号，通常五号
    21），不是段落自身字号——Word 就是按这个尺子量"N 字符"的（陷阱 #10/#12）。
    只用于必须写绝对值的地方（制表位没有字符单位形式），canonical 缩进本身一律
    只写字符单位（严格-spec §2.2）。"""
    return int(round(chars / 100.0 * size_hp * 10))


def _rpr(ea, size_hp, western_font=None):
    """run 属性：eastAsia = 中文字体；ascii/hAnsi/cs = 西文字体。套西文的角色给
    Times，否则回退成中文字体本身（数字也走该字体）。"""
    latin = western_font or ea
    parts = ['<w:rFonts w:ascii="%s" w:hAnsi="%s" w:eastAsia="%s" w:cs="%s"/>'
             % (_esc(latin), _esc(latin), _esc(ea), _esc(latin))]
    if size_hp is not None:
        parts.append('<w:sz w:val="%d"/><w:szCs w:val="%d"/>' % (size_hp, size_hp))
    return "".join(parts)


def _ind(first_line_chars=None, left_chars=None, no_indent=False):
    """缩进：实际缩进量**只写字符单位**（firstLineChars/leftChars），绝不补非零的
    绝对伴随值——严格-spec §2.2，这样 Word 才显示"N 字符"而非厘米（陷阱 #10）。

    同时把**不要的方向显式写 0**（含绝对形式）。零值没有单位歧义，写 0 不违反
    §2.2，却能挡住从 basedOn 父样式继承来的绝对左/右缩进——注入进真实文档时，
    父样式（往往是文档自己的 Normal）可能带缩进，不显式归零就会漏进来。"""
    a = []
    if no_indent:
        a += ['w:firstLine="0"', 'w:firstLineChars="0"',
              'w:left="0"', 'w:leftChars="0"', 'w:right="0"', 'w:rightChars="0"']
    else:
        if first_line_chars is not None:
            a.append('w:firstLineChars="%d"' % first_line_chars)
            a += ['w:left="0"', 'w:leftChars="0"', 'w:right="0"', 'w:rightChars="0"']
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


def _style_xml(sid, name, ppr_inner, rpr_inner, based_on=None):
    based = '<w:basedOn w:val="%s"/>' % _esc(based_on) if based_on else ''
    return ('<w:style w:type="paragraph" w:styleId="%s"><w:name w:val="%s"/>%s'
            '<w:qFormat/><w:pPr>%s</w:pPr><w:rPr>%s</w:rPr></w:style>'
            % (_esc(sid), _esc(name), based, ppr_inner, rpr_inner))


# ---------------------------------------------------------------------------
# 角色 -> canonical 样式
# ---------------------------------------------------------------------------
def _governs(fonts=False, western=False, size=False, line=False,
             space_before_after=False, jc=False, ind=False, outline=False):
    """该样式**承载**（因而段落上要清掉的）属性集合。

    这是"清直接覆盖"与 45 步全坍缩不变量的共同依据：样式承载什么，段落就不许再
    用直接属性表达同一属性；样式**没**承载的属性（如密级行的对齐方式——规范对它
    没有规定）保持原样不动，避免样式指派顺手抹掉模板版式。"""
    return {"fonts": fonts, "western": western, "size": size, "line": line,
            "space_before_after": space_before_after, "jc": jc,
            "ind": ind, "outline": outline}


def canonical_style_defs(spec):
    """按 spec 造全角色 canonical 样式定义。

    返回 [{"role","id","name","xml","governs"}]，顺序稳定（ROLE_STYLES 的顺序）。
    ``xml`` 是一个完整 `<w:style>` 元素的字符串，参考件与 40 共用同一份。"""
    western = spec.get("western_font", "Times New Roman")
    ls = spec.get("line_spacing", {})
    line_twips = ls.get("line_twips")
    line_rule = ls.get("line_rule", "exact")
    out = []

    def add(role, ppr, rpr, governs, based_on=None):
        sid = STYLE_ID_BY_ROLE[role]
        name = next(n for r, _s, n in ROLE_STYLES if r == role)
        out.append({"role": role, "id": sid, "name": name, "governs": governs,
                    "xml": _style_xml(sid, name, ppr, rpr, based_on)})

    def latin(entry):
        """该角色是否套西文（Times）。封面各角色由 spec 的 enforce_western 决定
        （阶段0 Word 验收：封面数字/西文走 Times），正文类角色一律套。"""
        return western if entry.get("enforce_western") else None

    # 封面题目：方正小标宋_GBK 小一、居中、无缩进
    t = spec["title"]
    add("title",
        '<w:jc w:val="%s"/>' % t.get("jc", "center") + _ind(no_indent=True),
        _rpr(t["east_asia"], t["size_hp"], latin(t)),
        _governs(fonts=True, western=bool(latin(t)), size=True, jc=True, ind=True))

    # 封面密级/文本编号行：只管字体字号——规范对它的对齐/缩进没有规定，样式便
    # 不承载，模板原有版式（居中、缩进）保持不动。
    cc = spec["cover_classification"]
    add("cover_classification", "", _rpr(cc["east_asia"], cc["size_hp"], latin(cc)),
        _governs(fonts=True, western=bool(latin(cc)), size=True))

    # 封面题目下要素：2倍行距、两端对齐、首行缩进2字符
    cf = spec["cover_field"]
    add("cover_field",
        _spacing(cf.get("line_twips"), cf.get("line_rule"))
        + '<w:jc w:val="both"/>' + _ind(first_line_chars=200),
        _rpr(cf["east_asia"], cf["size_hp"], latin(cf)),
        _governs(fonts=True, western=bool(latin(cf)), size=True,
                 line=bool(cf.get("line_twips")), jc=True, ind=True))

    # 一~四级标题：黑体/楷体/仿宋/仿宋 三号、行距固定28磅、首行缩进2字符、大纲级别
    for lvl in ("1", "2", "3", "4"):
        h = spec["headings"][lvl]
        add("heading" + lvl,
            '<w:outlineLvl w:val="%d"/>' % (int(lvl) - 1)
            + _spacing(line_twips, line_rule) + _ind(first_line_chars=h["first_line_chars"]),
            _rpr(h["east_asia"], h["size_hp"], western),
            _governs(fonts=True, western=True, size=True, line=line_twips is not None,
                     ind=True, outline=True))

    # 正文
    b = spec["body"]
    add("body",
        _spacing(line_twips, line_rule, zero_before_after=b.get("no_space_before_after"))
        + _ind(first_line_chars=b["first_line_chars"]),
        _rpr(b["east_asia"], b["size_hp"], western),
        _governs(fonts=True, western=True, size=True, line=line_twips is not None,
                 space_before_after=bool(b.get("no_space_before_after")), ind=True))

    # 图/表标题
    cap = spec["caption_format"]
    add("caption",
        _spacing(cap.get("line_twips"), cap.get("line_rule"))
        + '<w:jc w:val="%s"/>' % cap.get("jc", "center") + _ind(no_indent=True),
        _rpr(cap["east_asia"], cap["size_hp"], western),
        _governs(fonts=True, western=True, size=True,
                 line=bool(cap.get("line_twips")), jc=True, ind=bool(cap.get("no_indent"))))

    # 表格内容
    tb = spec["table_body"]
    add("table_body",
        _spacing(tb.get("line_twips"), tb.get("line_rule")) + _ind(no_indent=True),
        _rpr(tb["east_asia"], tb["size_hp"], western),
        _governs(fonts=True, western=True, size=True,
                 line=bool(tb.get("line_twips")), ind=bool(tb.get("no_indent"))))

    # "目录"二字：用正文字体字号但居中、无缩进（它不是标题、不进 TOC）。
    # 参考件用得到；流水线里这行按目录区处理（见 checks.paragraph_role）。
    add("toc_title",
        _spacing(line_twips, line_rule) + '<w:jc w:val="center"/>' + _ind(no_indent=True),
        _rpr(b["east_asia"], b["size_hp"], western),
        _governs(fonts=True, western=True, size=True, line=line_twips is not None,
                 jc=True, ind=True))
    return out


def governs_by_style_id(spec):
    """{styleId: governs} —— 45 步全坍缩不变量按输出文档里的 pStyle 反查用。"""
    return {d["id"]: d["governs"] for d in canonical_style_defs(spec)}


# ---------------------------------------------------------------------------
# docDefaults / Normal —— 文档网格字体的载体（陷阱 #12）
# ---------------------------------------------------------------------------
def doc_defaults_xml(spec):
    """docDefaults：默认西文 Times / 中文正文字体 / **五号(21)** + lang。

    默认字号必须是五号：它同时是 *Chars 缩进的字符单位尺子，也是行网格能否落到
    15.6磅/41行 的前提（陷阱 #12）。"""
    grid = spec.get("document_grid", {})
    size = grid.get("doc_defaults_size_hp", 21)
    ea = spec.get("body", {}).get("east_asia", "仿宋")
    western = spec.get("western_font", "Times New Roman")
    lang = grid.get("theme_font_lang_east_asia", "zh-CN")
    return ('<w:docDefaults><w:rPrDefault><w:rPr>'
            '<w:rFonts w:ascii="%s" w:eastAsia="%s" w:hAnsi="%s" w:cs="%s"/>'
            '<w:sz w:val="%d"/><w:szCs w:val="%d"/>'
            '<w:lang w:val="en-US" w:eastAsia="%s" w:bidi="ar-SA"/>'
            '</w:rPr></w:rPrDefault><w:pPrDefault/></w:docDefaults>'
            % (_esc(western), _esc(ea), _esc(western), _esc(western),
               size, size, _esc(lang)))


def normal_style_xml(spec):
    """Normal(正文) 样式 ＝ **五号**：只驱动文档网格字体，正文内容走 FGW正文(三号)。

    用户实测（陷阱 #12）：Word 的【文档网格字体】就是 Normal 的字号；Normal=三号
    时行高 21.75磅 > 15.6磅，网格被顶高成每页 29 行。故这里必须是五号，正文内容
    另由 FGWCanonBody 承载三号。"""
    grid = spec.get("document_grid", {})
    size = grid.get("normal_size_hp", 21)
    ea = spec.get("body", {}).get("east_asia", "仿宋")
    western = spec.get("western_font", "Times New Roman")
    return ('<w:style w:type="paragraph" w:default="1" w:styleId="%s">'
            '<w:name w:val="Normal"/><w:qFormat/>'
            '<w:rPr><w:rFonts w:ascii="%s" w:eastAsia="%s" w:hAnsi="%s" w:cs="%s"/>'
            '<w:sz w:val="%d"/><w:szCs w:val="%d"/></w:rPr></w:style>'
            % (NORMAL_STYLE_ID, _esc(western), _esc(ea), _esc(western),
               _esc(western), size, size))


# ---------------------------------------------------------------------------
# 目录样式（TOC1..N）
# ---------------------------------------------------------------------------
def toc_level_props(toc_spec, level, char_unit_hp=21):
    """某一级目录样式的 canonical pPr 取值：

        {"left_chars": 左缩进(字符×100),
         "tabs": [(val, pos_twips, leader_or_None), ...]}

    制表位**必须写绝对 twips**（制表位没有字符单位形式），按 spec 的字符数 ×
    文档字符单位字号换算——Normal=五号(21) → 左 840/1050/1260、右点线 8665；
    Normal=三号(32) 则 1280/1600/1920 + 13203（陷阱 #12：字符单位随 Normal 走）。
    """
    lv = str(level) if level is not None else None
    left_chars = (toc_spec.get("indent_chars_by_level") or {}).get(lv, 0) if lv else 0
    tabs = []
    left_tab_chars = (toc_spec.get("tab_left_chars_by_level") or {}).get(lv) if lv else None
    if left_tab_chars is not None:
        tabs.append(("left", char_twips(left_tab_chars, char_unit_hp), None))
    right_chars = toc_spec.get("tab_right_chars")
    if right_chars is not None:
        tabs.append(("right", char_twips(right_chars, char_unit_hp),
                     toc_spec.get("tab_leader") or "dot"))
    return {"left_chars": left_chars, "tabs": tabs}


def toc_style_xml(spec, level, char_unit_hp=21):
    """参考件用：整条 TOC{level} 样式（basedOn Normal，仿宋 小三 + 缩进 + 制表位）。
    40 不用这个——真实文档里 TOC 域引用的是文档**自己的**目录样式，只能就地 patch
    （见 40 的 `_patch_toc_styles`），两边共用 `toc_level_props` 的取值。"""
    toc = spec["toc"]
    props = toc_level_props(toc, level, char_unit_hp)
    tabs_xml = "".join(
        '<w:tab w:val="%s" %sw:pos="%d"/>'
        % (val, ('w:leader="%s" ' % leader) if leader else "", pos)
        for val, pos, leader in props["tabs"])
    tabs_xml = ('<w:tabs>%s</w:tabs>' % tabs_xml) if tabs_xml else ""
    lc = props["left_chars"]
    ind = ('<w:ind w:leftChars="%d" w:left="%d"/>' % (lc, lc)) if lc else ""
    rpr = ('<w:rFonts w:ascii="%s" w:hAnsi="%s" w:cs="%s"/>'
           '<w:sz w:val="%d"/><w:szCs w:val="%d"/>'
           % (_esc(toc["east_asia"]), _esc(toc["east_asia"]), _esc(toc["east_asia"]),
              toc["size_hp"], toc["size_hp"]))
    return ('<w:style w:type="paragraph" w:styleId="TOC%s"><w:name w:val="toc %s"/>'
            '<w:basedOn w:val="%s"/><w:next w:val="%s"/>'
            '<w:uiPriority w:val="39"/><w:qFormat/>'
            '<w:pPr>%s</w:pPr><w:rPr>%s</w:rPr></w:style>'
            % (level, level, NORMAL_STYLE_ID, NORMAL_STYLE_ID, tabs_xml + ind, rpr))


# ---------------------------------------------------------------------------
# 文档级：settings 的网格/兼容块 + sectPr 的 docGrid
# ---------------------------------------------------------------------------
def settings_children_xml(spec):
    """settings.xml 里承载【文档网格】的子元素，返回 [(localname, xml)]。

    行网格 15.6磅/41行 不是只靠 sectPr 的 docGrid——**compat 块（尤其 useFELayout
    + compatibilityMode=15）才是开关**，缺它 Word 按旧版式把行距顶到 21.75磅/29行
    （陷阱 #12，用户实测）。defaultTabStop=420(2字符) 决定标题自动编号后制表符落点。
    """
    grid = spec.get("document_grid") or {}
    if not grid:
        return []
    out = []
    tab = grid.get("default_tab_stop_twips")
    if tab is not None:
        out.append(("defaultTabStop", '<w:defaultTabStop w:val="%d"/>' % tab))
    for key, tag in (("drawing_grid_h_twips", "drawingGridHorizontalSpacing"),
                     ("drawing_grid_v_twips", "drawingGridVerticalSpacing"),
                     ("display_grid_every_h", "displayHorizontalDrawingGridEvery"),
                     ("display_grid_every_v", "displayVerticalDrawingGridEvery")):
        if grid.get(key) is not None:
            out.append((tag, '<w:%s w:val="%d"/>' % (tag, grid[key])))
    if grid.get("no_punctuation_kerning"):
        out.append(("noPunctuationKerning", '<w:noPunctuationKerning/>'))
    if grid.get("character_spacing_control"):
        out.append(("characterSpacingControl",
                    '<w:characterSpacingControl w:val="%s"/>'
                    % _esc(grid["character_spacing_control"])))
    flags = "".join('<w:%s/>' % f for f in (grid.get("compat_flags") or []))
    settings = "".join(
        '<w:compatSetting w:name="%s" w:uri="http://schemas.microsoft.com/office/word" '
        'w:val="%s"/>' % (_esc(k), _esc(v))
        for k, v in (grid.get("compat_settings") or {}).items())
    if flags or settings:
        out.append(("compat", '<w:compat>%s%s</w:compat>' % (flags, settings)))
    if grid.get("theme_font_lang_east_asia"):
        out.append(("themeFontLang", '<w:themeFontLang w:val="en-US" w:eastAsia="%s"/>'
                    % _esc(grid["theme_font_lang_east_asia"])))
    return out


def table_ind_xml(spec):
    """表格左缩进（tblInd）。**与单元格边距 tblCellMar 是两回事**，别混（陷阱 #12）。"""
    td = spec.get("table_defaults") or {}
    if td.get("tbl_ind_twips") is None:
        return ""
    return '<w:tblInd w:w="%d" w:type="dxa"/>' % td["tbl_ind_twips"]


def cell_margins_xml(spec):
    """表级默认单元格边距 tblCellMar（上0/左0.19cm/下0/右0.19cm）。"""
    td = spec.get("table_defaults") or {}
    cm = td.get("cell_margin_twips") or {}
    if not cm:
        return ""
    return ('<w:tblCellMar>%s</w:tblCellMar>'
            % "".join('<w:%s w:w="%d" w:type="dxa"/>' % (side, cm.get(side, 0))
                      for side in ("top", "left", "bottom", "right")))


def cell_valign_xml(spec):
    """单元格内容垂直对齐（vAlign）。"""
    td = spec.get("table_defaults") or {}
    va = td.get("cell_valign")
    return ('<w:vAlign w:val="%s"/>' % _esc(va)) if va else ""


def doc_grid_attrs(spec):
    """sectPr/docGrid 的属性 {type, linePitch}；spec 没配网格时返回 {}。"""
    grid = spec.get("document_grid") or {}
    if grid.get("line_pitch_twips") is None:
        return {}
    return {"type": grid.get("grid_type", "lines"),
            "linePitch": str(grid["line_pitch_twips"])}
