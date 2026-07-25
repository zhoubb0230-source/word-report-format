# -*- coding: utf-8 -*-
"""
cascade.py — 方案C 的"统一 cascade + 归属"归一层。

把散落在流水线各处（`20_extract_structure.py` 兜编号层 ind、`40_apply_fixes.py`
判断该不该动某段）对"有效属性从哪层来"的推断，收进**一个**地方。核心是：一段
paragraph 的有效格式属性由四层按优先级叠出——

    docDefaults  <  段落样式链  <  编号层  <  直接属性

（见 HANDOFF_方案C实施.md §2.3）。这里既给出**折叠后的有效值**，也给出**每个属性
由哪一层供给**（provenance）。后者是方案C"量化还剩几层没钳干净"的量尺：apply 把
canonical 值写进样式层后，凡是有效值仍由 `numbering`/`direct` 供给的属性，就是还没
钳干净、会在 Word 里泄漏的残留。

本模块**纯读、无副作用、不改任何现有脚本的行为**（阶段1"零行为改变"）：它只是把
`StyleResolver.resolve_cascade` 包一层便捷 API，并额外产出 provenance。仅依赖 stdlib
+ lxml + 同目录的 docxcommon。
"""

from docxcommon import (
    qn, read_ppr, get_pPr, get_style_id, get_mark_rpr,
    INDENT_KEYS,
)

# 叠加顺序 = 优先级从低到高。后一层的非 None 值覆盖前一层。
LAYER_ORDER = ("docDefaults", "style", "numbering", "direct")


def _overlay(base, extra, owner, layer):
    """把 extra 的非 None 值叠到 base 上，并记录该键的供给层为 layer。"""
    for k, v in extra.items():
        if v is not None:
            base[k] = v
            owner[k] = layer


def effective_paragraph_props(p, resolver, num_levels=None):
    """便捷入口：给一个 <w:p>，返回 resolve_cascade 折出的 (ppr, rpr)。

    等价于 20_extract_structure.py 里 `resolver.resolve(...)` + 手动兜编号层
    ind 的组合，但用的是**正确优先级**（编号层压过样式、直接压过编号层），且
    覆盖全部 pPr 键而不仅是 INDENT_KEYS。仅供检测/坍缩不变量按需调用，**不**在
    本阶段改写 20 的既有逻辑。"""
    sid = get_style_id(p)
    ppr_el = get_pPr(p)
    mark_rpr = get_mark_rpr(p)
    return resolver.resolve_cascade(sid, ppr_el, mark_rpr, num_levels)


def resolve_ppr_with_provenance(resolver, style_id, direct_ppr_el,
                                mark_rpr_el=None, num_levels=None):
    """折叠四层 pPr，返回 (eff, owner)：

      * eff   —— 折叠后的有效 pPr dict（与 resolve_cascade 的 ppr 一致）。
      * owner —— {属性名: 供给该有效值的层名}，层名取自 LAYER_ORDER。
                  未被任何层设置的键不出现在 owner 里。

    provenance 是方案C 的诊断量尺：apply 后重跑本函数，任一 canonical 承载的
    属性若 owner 仍是 'numbering' 或 'direct'，即该层未被钳干净、会在 Word 泄漏。
    此处不做判定、不比 spec，只如实报告"谁供的值"。"""
    eff, owner = {}, {}

    # 1) docDefaults
    _overlay(eff, dict(resolver.doc_ppr), owner, "docDefaults")

    # 2) 段落样式链（basedOn 由根向下，深者覆盖浅者），压成一层 'style'
    style_ppr = {}
    sid = style_id or resolver.default_para_style
    for st in resolver._style_chain(sid):
        spr = st.find(qn("w:pPr"))
        if spr is not None:
            _overlay(style_ppr, read_ppr(spr), {}, "style")
    _overlay(eff, style_ppr, owner, "style")

    # 3) 编号层——用"直接优先于样式链"的 numId/ilvl 去查级别 pPr
    direct_ppr = read_ppr(direct_ppr_el) if direct_ppr_el is not None else {}
    num_id = direct_ppr.get("num_id")
    if num_id is None:
        num_id = style_ppr.get("num_id") or eff.get("num_id")
    ilvl = direct_ppr.get("ilvl")
    if ilvl is None:
        ilvl = style_ppr.get("ilvl") or eff.get("ilvl")
    if num_levels and num_id and num_id != "0":
        try:
            il = int(ilvl) if ilvl is not None else 0
        except (ValueError, TypeError):
            il = 0
        lvl_ppr = num_levels.get(num_id, {}).get(il)
        if lvl_ppr:
            _overlay(eff, lvl_ppr, owner, "numbering")

    # 4) 直接属性（最高优先级）
    if direct_ppr:
        _overlay(eff, direct_ppr, owner, "direct")

    return eff, owner


def contested_indent_report(resolver, style_id, direct_ppr_el,
                            mark_rpr_el=None, num_levels=None):
    """只挑缩进类键（INDENT_KEYS）汇报归属：{键: 供给层}。

    缩进是方案C 头号战场（首行缩进/左缩进/悬挂缩进在样式层、编号层、直接层间
    抢值——见 §2.1/§2.3）。返回值只含"确有某层设置了"的键，便于对抗性 fixtures
    断言"竞争值坍缩到预期层"。"""
    _eff, owner = resolve_ppr_with_provenance(
        resolver, style_id, direct_ppr_el, mark_rpr_el, num_levels)
    return {k: owner[k] for k in INDENT_KEYS if k in owner}
