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
# 图/表标题**分成两个样式**（名字里各只含"图"或"表"）：图表标题一旦改用 Word 自动编号，
# 文字里的静态"图N/表N"前缀就被删掉了，此后**只能靠样式名判断这是图标题还是表标题**
# （`headings.caption_kind_from_style` 对同时含图与表的"图表标题"返回 None＝认不出）。
# 合成一个样式会让二次运行认不出种类、把图表标题降级成正文。
#
# ⚠️ 样式**名**统一带 "FGW" 前缀：Word 要求样式名在文档内**唯一**，重名会让它在打开时
# 报"发现无法读取的内容"。而"标题 1""图标题""表标题""目录标题"这些名字在真实模板里
# 很可能已经存在（LibreOffice 转换出来的 docx 尤其爱写本地化名），直接叫这些名字必然
# 撞车。前缀既保证唯一，也让用户在样式列表里一眼认出哪些是本工具注入的。
# 前缀不影响识别：`heading_level_from_style_name("FGW标题1")` 仍得 1，
# `caption_kind_from_style` 仍能从"FGW图标题/FGW表标题"里认出图与表。
ROLE_STYLES = (
    ("title",                "FGWCanonTitle",      "FGW封面题目"),
    ("cover_classification", "FGWCanonCoverClass", "FGW封面密级编号"),
    ("cover_field",          "FGWCanonCoverField", "FGW封面要素"),
    ("heading1",             "FGWCanonH1",         "FGW标题1"),
    ("heading2",             "FGWCanonH2",         "FGW标题2"),
    ("heading3",             "FGWCanonH3",         "FGW标题3"),
    ("heading4",             "FGWCanonH4",         "FGW标题4"),
    ("body",                 "FGWCanonBody",       "FGW正文"),
    ("caption_figure",       "FGWCanonCaptionFig", "FGW图标题"),
    ("caption_table",        "FGWCanonCaptionTbl", "FGW表标题"),
    ("table_body",           "FGWCanonTableBody",  "FGW表格内容"),
    ("toc_title",            "FGWCanonTocTitle",   "FGW目录标题"),
)

# 角色名 -> 图表种类，供编号入列时选对序列。
CAPTION_ROLE_BY_KIND = {"figure": "caption_figure", "table": "caption_table"}

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


def _rpr(ea, size_hp, western_font=None, bold=False):
    """run 属性：eastAsia = 中文字体；ascii/hAnsi/cs = 西文字体。套西文的角色给
    Times，否则回退成中文字体本身（数字也走该字体）。``bold`` 写 `w:b`+`w:bCs`
    （中日韩与西文各一套开关，只写 `w:b` 时西文部分可能不加粗）。"""
    latin = western_font or ea
    parts = ['<w:rFonts w:ascii="%s" w:hAnsi="%s" w:eastAsia="%s" w:cs="%s"/>'
             % (_esc(latin), _esc(latin), _esc(ea), _esc(latin))]
    if bold:
        parts.append('<w:b/><w:bCs/>')
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


def _ppr(spacing="", ind="", jc="", outline=""):
    """按 **CT_PPrBase 的 sequence** 拼 pPr：… spacing → ind → jc → … → outlineLvl。

    别按"想到哪写到哪"的顺序拼——WordprocessingML 是 xsd:sequence，顺序错了文件仍是
    良构 XML，但真实 Word 会拒绝打开（"发现无法读取的内容"）。这个函数就是为了让调用
    方无法写错顺序：参数名对应元素，拼装顺序在这里定死一次。"""
    return spacing + ind + jc + outline


def _style_xml(sid, name, ppr_inner, rpr_inner, based_on=None):
    based = '<w:basedOn w:val="%s"/>' % _esc(based_on) if based_on else ''
    return ('<w:style w:type="paragraph" w:styleId="%s"><w:name w:val="%s"/>%s'
            '<w:qFormat/><w:pPr>%s</w:pPr><w:rPr>%s</w:rPr></w:style>'
            % (_esc(sid), _esc(name), based, ppr_inner, rpr_inner))


# ---------------------------------------------------------------------------
# 角色 -> canonical 样式
# ---------------------------------------------------------------------------
def _governs(fonts=False, western=False, size=False, bold=False, line=False,
             space_before_after=False, jc=False, ind=False, outline=False):
    """该样式**承载**（因而段落上要清掉的）属性集合。

    这是"清直接覆盖"与 45 步全坍缩不变量的共同依据：样式承载什么，段落就不许再
    用直接属性表达同一属性；样式**没**承载的属性（如密级行的对齐方式——规范对它
    没有规定）保持原样不动，避免样式指派顺手抹掉模板版式。"""
    return {"fonts": fonts, "western": western, "size": size, "bold": bold,
            "line": line, "space_before_after": space_before_after, "jc": jc,
            "ind": ind, "outline": outline}


def canonical_style_defs(spec, caption_num_ids=None):
    """按 spec 造全角色 canonical 样式定义。

    返回 [{"role","id","name","xml","governs"}]，顺序稳定（ROLE_STYLES 的顺序）。
    ``xml`` 是一个完整 `<w:style>` 元素的字符串，参考件与 40 共用同一份。

    ``caption_num_ids`` = {"figure": numId, "table": numId}：把图/表标题的自动编号
    **写进样式本身**（`w:numPr`），而不是只挂在段落上。这样用户在 Word 里插入一张新图、
    给标题**套上"图标题"样式**就直接生成编号——只挂段落的话新段落不会有编号（用户实测
    反馈）。numId 是每个文档各自分配的，所以由调用方在注入编号定义后传进来。"""
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
        _ppr(ind=_ind(no_indent=True), jc='<w:jc w:val="%s"/>' % t.get("jc", "center")),
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
        _ppr(spacing=_spacing(cf.get("line_twips"), cf.get("line_rule")),
             ind=_ind(first_line_chars=200), jc='<w:jc w:val="both"/>'),
        _rpr(cf["east_asia"], cf["size_hp"], latin(cf)),
        _governs(fonts=True, western=bool(latin(cf)), size=True,
                 line=bool(cf.get("line_twips")), jc=True, ind=True))

    # 一~四级标题：黑体/楷体/仿宋/仿宋 三号加粗、行距固定28磅、首行缩进2字符、大纲级别
    for lvl in ("1", "2", "3", "4"):
        h = spec["headings"][lvl]
        add("heading" + lvl,
            _ppr(spacing=_spacing(line_twips, line_rule),
                 ind=_ind(first_line_chars=h["first_line_chars"]),
                 outline='<w:outlineLvl w:val="%d"/>' % (int(lvl) - 1)),
            _rpr(h["east_asia"], h["size_hp"], western, bold=bool(h.get("bold"))),
            _governs(fonts=True, western=True, size=True, bold=bool(h.get("bold")),
                     line=line_twips is not None, ind=True, outline=True))

    # 正文
    b = spec["body"]
    add("body",
        _ppr(spacing=_spacing(line_twips, line_rule,
                              zero_before_after=b.get("no_space_before_after")),
             ind=_ind(first_line_chars=b["first_line_chars"])),
        _rpr(b["east_asia"], b["size_hp"], western),
        _governs(fonts=True, western=True, size=True, line=line_twips is not None,
                 space_before_after=bool(b.get("no_space_before_after")), ind=True))

    # 图标题 / 表标题（同一套格式，两个样式名——种类要靠样式名回读，见 ROLE_STYLES）
    cap = spec["caption_format"]
    nums = caption_num_ids or {}
    for role, kind in (("caption_figure", "figure"), ("caption_table", "table")):
        # 自动编号写进样式：套上样式即生成编号（numPr 在 CT_PPrBase 里排在 spacing 前）
        num_id = nums.get(kind)
        numpr = ('<w:numPr><w:ilvl w:val="0"/><w:numId w:val="%s"/></w:numPr>'
                 % _esc(num_id)) if num_id is not None else ""
        add(role,
            numpr + _ppr(spacing=_spacing(cap.get("line_twips"), cap.get("line_rule")),
                         ind=_ind(no_indent=True),
                         jc='<w:jc w:val="%s"/>' % cap.get("jc", "center")),
            _rpr(cap["east_asia"], cap["size_hp"], western),
            _governs(fonts=True, western=True, size=True,
                     line=bool(cap.get("line_twips")), jc=True,
                     ind=bool(cap.get("no_indent"))))

    # 表格内容
    tb = spec["table_body"]
    add("table_body",
        _ppr(spacing=_spacing(tb.get("line_twips"), tb.get("line_rule")),
             ind=_ind(no_indent=True)),
        _rpr(tb["east_asia"], tb["size_hp"], western),
        _governs(fonts=True, western=True, size=True,
                 line=bool(tb.get("line_twips")), ind=bool(tb.get("no_indent"))))

    # "目录"二字：用正文字体字号但居中、无缩进（它不是标题、不进 TOC）。
    # 参考件用得到；流水线里这行按目录区处理（见 checks.paragraph_role）。
    add("toc_title",
        _ppr(spacing=_spacing(line_twips, line_rule), ind=_ind(no_indent=True),
             jc='<w:jc w:val="center"/>'),
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


# ---------------------------------------------------------------------------
# 图/表标题的自动编号
# ---------------------------------------------------------------------------
# 注入的编号定义靠 `<w:name>` 打标签认领，二次运行才能找回同一条定义而不是越注入越多。
CAPTION_NUM_MARKER = {"figure": "FGWCaptionFigure", "table": "FGWCaptionTable"}


def numbering_level_xml(num_fmt, lvl_text, suff=None, ilvl=0,
                        east_asia=None, western=None):
    """一个编号级别，**缩进已中和**（left=0、无 hanging）。

    中和是必须的：编号层在 cascade 里压过样式层，级别自带的 hanging 会把 canonical
    样式的缩进盖掉（陷阱 #11）。`suff` 控编号与文字之间的分隔符——图/表标题要
    `tab`（编号后一个制表符），Word 的默认值也是 tab，这里显式写出来免得依赖默认。

    ``east_asia`` / ``western`` 给这一级的**序号本身**指定字体（写成级别的 `w:rPr`，
    等价于 Word「定义多级列表 → 字体」里设的那一套）。不写就没有 rPr，序号跟随段落
    标记 → 段落样式；写了则序号严格按这里渲染，**与标题文字的字体解耦**——规范里
    "一级序号黑体、二级半角括号 Times New Roman、四级序号仿宋"就靠它表达。
    **只写 rFonts**：字号/加粗仍旧继承标题样式，免得序号与标题文字一粗一细（#13/#19）。

    **子元素顺序必须照 CT_Lvl 的 sequence**：start → numFmt → lvlRestart → pStyle →
    isLgl → **suff** → lvlText → … → pPr → **rPr**。把 `suff` 写到 `numFmt` 前面文件
    依然良构、45 的 XML 检查也过，但**真实 Word 直接拒绝打开**（"发现无法读取的内容"）
    ——踩过一次。`rPr` 同理，必须排在 `pPr` **之后**。"""
    suff_el = ('<w:suff w:val="%s"/>' % _esc(suff)) if suff else ""
    rpr_el = ""
    if east_asia or western:
        # 只写给了值的那一半：没给中文字体就别拿西文字体去填 eastAsia（缺省属性会
        # 正常继承段落标记/样式，硬填反而把中文序号拽成西文字体）。
        attrs = []
        if western:
            attrs.append('w:ascii="%s" w:hAnsi="%s" w:cs="%s"'
                         % (_esc(western), _esc(western), _esc(western)))
        if east_asia:
            attrs.append('w:eastAsia="%s"' % _esc(east_asia))
        rpr_el = '<w:rPr><w:rFonts %s/></w:rPr>' % " ".join(attrs)
    return ('<w:lvl w:ilvl="%d"><w:start w:val="1"/><w:numFmt w:val="%s"/>%s'
            '<w:lvlText w:val="%s"/><w:lvlJc w:val="left"/>'
            '<w:pPr><w:ind w:left="0" w:leftChars="0"/></w:pPr>%s</w:lvl>'
            % (ilvl, _esc(num_fmt), suff_el, _esc(lvl_text), rpr_el))


def caption_numbering_defs(spec):
    """图/表标题各一条 canonical 自动编号定义。

    返回 [{"kind","marker","multi_level_type","lvl_text","lvl_xml"}]（figure 在前）。
    **静态编号不可取**：插入/删除一张图之后所有后续编号都要重排，而 Word 自动编号自己
    维护序列；用户阶段2 验收明确指出"图/表标题采用了静态编号，这一点不正确"。

    `spec.captions` 为真源：前缀取 `figure_prefixes[0]`/`table_prefixes[0]`，编号格式
    取 `num_fmt`，编号后分隔符取 `suffix`。"""
    caps = spec.get("captions") or {}
    if not caps.get("auto_number"):
        return []
    num_fmt = caps.get("num_fmt", "decimal")
    suff = caps.get("suffix", "tab")
    out = []
    for kind, key in (("figure", "figure_prefixes"), ("table", "table_prefixes")):
        prefixes = caps.get(key) or []
        if not prefixes:
            continue
        lvl_text = "%s%%1" % prefixes[0]
        out.append({"kind": kind, "marker": CAPTION_NUM_MARKER[kind],
                    "multi_level_type": "singleLevel", "lvl_text": lvl_text,
                    "lvl_xml": numbering_level_xml(num_fmt, lvl_text, suff)})
    return out


# ---------------------------------------------------------------------------
# 一~四级标题的多级自动编号
# ---------------------------------------------------------------------------
HEADING_NUM_MARKER = "FGWHeadingNumbering"

# 每级 (numFmt, lvlText 模板) 的**兜底**值——正常取 `spec.heading_numbering.levels`
# （判定值只来自 spec），spec 缺这一块时才用这里。**必须与 `checks._heading_token` 的
# 静态 token 同形**——一、/(一)/1./（1）。文档里两种编号方式（手写在文字里的、Word
# 自动生成的）常常并存，形状不一致就成了"同一篇文档两套编号规则"。
# 二级用**半角括号** `(一)`：用户第六轮验收明确要求"二级标题前后的括号是英文括号"，
# 阶段0 的参考件本来也是半角（全角括号更宽，在目录里会越过左制表位）。
HEADING_LEVEL_SHAPES = (
    ("chineseCounting", "%1、"),
    ("chineseCounting", "(%2)"),
    ("decimal", "%3."),
    ("decimal", "（%4）"),
)


def heading_level_shapes(spec=None):
    """四级标题编号的 (numFmt, lvlText)，优先取 spec.heading_numbering.levels。"""
    levels = ((spec or {}).get("heading_numbering") or {}).get("levels") or {}
    out = []
    for i, fallback in enumerate(HEADING_LEVEL_SHAPES):
        entry = levels.get(str(i + 1)) or {}
        out.append((entry.get("num_fmt") or fallback[0],
                    entry.get("lvl_text") or fallback[1]))
    return tuple(out)


def heading_level_fonts(spec=None):
    """四级标题**序号自身**的 (中文字体, 西文字体)，取自 `heading_numbering.levels`。

    序号的字体不能一律跟着标题样式走（那是 2026-08 第六轮验收报的三个问题）：

      * 一级 `一、` 是中文，要与标题文字同为**黑体**——标题样式本身是黑体，但西文位
        走的是 Times，序号里万一出现西文就会不一致，故两边都钉成黑体；
      * 二级 `(一)` 的**半角括号是西文字符**，要求走 **Times New Roman**，而括号里的
        中文数字仍是标题的楷体——一个 rFonts 正好分别管 ascii/hAnsi 与 eastAsia；
      * 四级 `（1）` 的数字与全角括号**整体用仿宋**（不走 Times）。

    spec 没给某一级时的兜底：中文＝该级标题的 `east_asia`（"序号与标题内容同字体"的
    默认语义），西文＝`spec.western_font`。"""
    spec = spec or {}
    levels = (spec.get("heading_numbering") or {}).get("levels") or {}
    headings = spec.get("headings") or {}
    western_default = spec.get("western_font") or "Times New Roman"
    out = []
    for lvl in ("1", "2", "3", "4"):
        entry = levels.get(lvl) or {}
        ea = entry.get("east_asia") or (headings.get(lvl) or {}).get("east_asia")
        out.append((ea, entry.get("western") or western_default))
    return tuple(out)


def heading_numbering_def(spec=None):
    """canonical 的**四级标题编号**定义（一条 multilevel abstractNum）。

    为什么要自己注入一条、而不是沿用模板里那条并就地修补（2026-07 推翻上一轮做法）：

      * 规范要求标题**逐级重新编号**（二级在其一级下从头数）。Word 的逐级归零只在
        **同一个列表实例（同一个 numId）内**才成立；而真实模板里一~四级标题经常挂在
        **各自独立**的编号定义上（LibreOffice/WPS 转换尤其爱这么写），此时二级永远看
        不到一级出现、只能一路数下去——就是用户实测的"二/三/四级变成全局编号"。这种
        情况下无论怎么改 `lvlRestart` 都救不回来，唯一的解法是把四级**并进同一条多级
        列表**。
      * 级别的 `lvlText` 也是规范值（同 `_heading_token`），沿用模板的话自动编号会渲染
        成"1.1"这类模板自带形状，与手写序号被改成的"（一）"打架。

    级别里**不写 `lvlRestart`**：省略即 Word 默认的"上一级出现时归零"，正是规范要的
    逐级重新编号。缩进已中和（left=0、无 hanging），首行缩进由 canonical 标题样式承载。

    **编号与标题之间是制表符**（`suff=tab`，取自 `spec.heading_numbering.suffix`）：这是
    阶段0 参考件里经用户 Word 验收的排版（陷阱 #12「标题编号后制表符＝defaultTabStop」），
    落点由 settings 的 `defaultTabStop`(2字符) 决定。**别改成 `nothing`/`space`**——
    2026-07 曾按"与手写序号同形"的直觉改成不加分隔符，用户当轮就报"标题序号后的制表符
    没有了"。手写序号那条路径（`_heading_insert_prefix`）是纯文本、没有制表位可用，
    与这里不是一回事，不要互相看齐。

    **每级带自己的 `rPr`（序号字体）**：见 `heading_level_fonts`——序号跟着段落标记走
    的话，一级序号会是 Times 的西文位、二级的半角括号也是标题的楷体，正是用户第六轮
    报的问题。只钉字体，字号/加粗仍继承标题样式。"""
    suff = ((spec or {}).get("heading_numbering") or {}).get("suffix") or "tab"
    fonts = heading_level_fonts(spec)
    lvls = "".join(
        numbering_level_xml(fmt, lvl_text, suff=suff, ilvl=i,
                            east_asia=fonts[i][0], western=fonts[i][1])
        for i, (fmt, lvl_text) in enumerate(heading_level_shapes(spec)))
    return {"kind": "heading", "marker": HEADING_NUM_MARKER,
            "multi_level_type": "multilevel", "lvl_xml": lvls}


def injected_numbering_defs(spec):
    """本工具要注入 numbering.xml 的**全部** canonical 编号定义：图、表、标题。

    统一由 40 的 `_ensure_injected_numbering` 按 `marker` 幂等注入/认领。
    `spec.heading_numbering.auto_number` 关掉时不注入标题编号定义——此时自动编号的标题
    保持原样（`_heading_num_ref` 的"spec 没给定义"分支）。"""
    defs = caption_numbering_defs(spec)
    if ((spec.get("heading_numbering") or {}).get("auto_number")
            or "heading_numbering" not in spec):
        defs = defs + [heading_numbering_def(spec)]
    return defs


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
