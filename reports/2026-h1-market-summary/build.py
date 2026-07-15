import math
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.chart.data import CategoryChartData
from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION, XL_TICK_LABEL_POSITION
from pptx.oxml.ns import qn

from style import (
    PAPER, PAPER_2, PAPER_3, BORDER, INK, INK_2, INK_3,
    PRIMARY, PRIMARY_SOFT, PRIMARY_DEEP, COMPLETE, COMPLETE_SOFT,
    ASSET, ASSET_SOFT, RISK, RISK_SOFT,
    DISPLAY_FONT, BODY_FONT, MONO_FONT, SLIDE_W, SLIDE_H, MARGIN, CONTENT_W,
    set_bg, rect, oval, connector_line, rich, text, bullets, chip, badge_num,
    header, footer, bottom_conclusion, card, polar, E,
)

prs = Presentation()
prs.slide_width = SLIDE_W
prs.slide_height = SLIDE_H
BLANK = prs.slide_layouts[6]


def new_slide():
    s = prs.slides.add_slide(BLANK)
    set_bg(s, PAPER)
    return s


def add_cell_borders(cell, color_hex='DFD9CB'):
    """OOXML CT_TableCellProperties requires lnL/lnR/lnT/lnB to precede the
    fill element — appending them after cell.fill.solid() produces invalid
    XML that LibreOffice/PowerPoint refuse to open. Insert at the front."""
    tcPr = cell._tc.get_or_add_tcPr()
    for tag in ('a:lnB', 'a:lnT', 'a:lnR', 'a:lnL'):
        ln = tcPr.makeelement(qn(tag), {'w': '6350', 'cap': 'flat'})
        fill = ln.makeelement(qn('a:solidFill'), {})
        clr = fill.makeelement(qn('a:srgbClr'), {'val': color_hex})
        fill.append(clr)
        ln.append(fill)
        tcPr.insert(0, ln)


def plain_table(slide, left, top, width, height, data, col_ratios,
                 header_fill=PRIMARY, header_color=PAPER, body_font=11,
                 header_font=11, row_h=None, zebra=PAPER_2, zebra2=PAPER,
                 body_color=INK, align_cols=None):
    rows, cols = len(data), len(data[0])
    left, top, width, height = E(left), E(top), E(width), E(height)
    gtable = slide.shapes.add_table(rows, cols, left, top, width, height).table
    gtable.first_row = False
    gtable.horz_banding = False
    total = sum(col_ratios)
    for i, ratio in enumerate(col_ratios):
        gtable.columns[i].width = Emu(int(width * ratio / total))
    if row_h:
        for r in range(rows):
            gtable.rows[r].height = row_h
    for r in range(rows):
        for c in range(cols):
            cell = gtable.cell(r, c)
            cell.margin_left = Pt(8)
            cell.margin_right = Pt(8)
            cell.margin_top = Pt(4)
            cell.margin_bottom = Pt(4)
            cell.vertical_anchor = MSO_ANCHOR.MIDDLE
            is_header = r == 0
            cell.fill.solid()
            if is_header:
                cell.fill.fore_color.rgb = header_fill
            else:
                cell.fill.fore_color.rgb = zebra if r % 2 == 1 else zebra2
            tf = cell.text_frame
            tf.word_wrap = True
            p = tf.paragraphs[0]
            p.alignment = (align_cols[c] if align_cols else PP_ALIGN.LEFT)
            run = p.add_run()
            run.text = str(data[r][c])
            run.font.size = Pt(header_font if is_header else body_font)
            run.font.bold = is_header
            run.font.name = BODY_FONT
            run.font.color.rgb = header_color if is_header else body_color
    # thin uniform borders
    for r in range(rows):
        for c in range(cols):
            add_cell_borders(gtable.cell(r, c))
    return gtable


# ============================================================ PAGE 01
def page01():
    s = new_slide()
    # abstract scene: left thin "ocean" horizontal lines, right arc + satellite dots
    for i in range(5):
        y = Inches(5.35) + Inches(0.16) * i
        w = Inches(3.6) - Inches(0.22) * i
        connector_line(s, MARGIN, y, MARGIN + w, y, color=PRIMARY_SOFT, weight=1.4)
    connector_line(s, MARGIN, Inches(5.35 - 0.16), MARGIN + Inches(3.6), Inches(5.35 - 0.16),
                    color=PRIMARY, weight=1.6)
    # right: earth arc + orbit + satellite nodes
    cx, cy, r = Inches(12.3), Inches(8.1), Inches(3.7)
    big = oval(s, cx, cy, r, fill=None, line=PRIMARY, line_w=1.4)
    orbit = oval(s, cx, cy, r + Inches(0.55), fill=None, line=BORDER, line_w=1.0)
    for deg in (-150, -128, -108):
        x, y = polar(cx, cy, r + Inches(0.55), deg)
        oval(s, x, y, Emu(int(Inches(0.05))), fill=COMPLETE)
    text(s, MARGIN, Inches(2.35), Inches(9.5), Inches(0.5), "遨海科技市场部",
         size=26, bold=True, color=INK_2, font=DISPLAY_FONT, space_after=0)
    text(s, MARGIN, Inches(2.85), Inches(9.8), Inches(1.1), "2026 年半年工作总结",
         size=48, bold=True, color=PRIMARY, font=DISPLAY_FONT, space_after=0)
    connector_line(s, MARGIN, Inches(4.05), MARGIN + Inches(1.1), Inches(4.05),
                    color=RISK, weight=2.2)
    text(s, MARGIN, Inches(4.2), Inches(8.5), Inches(0.4),
         "从流程建立到 MTL 精细化运营", size=16, color=INK_2, space_after=0)
    text(s, MARGIN, Inches(6.75), Inches(6), Inches(0.35), "市场部｜2026 年 7 月",
         size=11.5, color=INK_3, font=MONO_FONT, space_after=0)
    text(s, SLIDE_W - MARGIN - Inches(1.5), Inches(0.42), Inches(1.5), Inches(0.32),
         "遨海科技", size=11, color=INK_2, bold=True, align=PP_ALIGN.RIGHT, space_after=0)


# ============================================================ PAGE 02
def page02():
    s = new_slide()
    cur = header(s, 2, None,
                 "公司进入“航海基本盘 + 航天突破”并行发展的新阶段",
                 "新的技术突破和业务空间，需要一个能够持续识别机会、组织资源、"
                 "沉淀资产并推动协同的经营前端。", title_size=22)
    cols = [
        ("01", "航海基本盘", COMPLETE, [
            "客户、区域项目、行业需求和政策窗口持续出现。",
            "方案、资质、案例和客户信息需要持续复用。",
        ]),
        ("02", "航天新突破", RISK, [
            ("VDES 卫星载荷完成在轨测试，", "验证效果达到预期。"),
            ("相关特种领域项目取得实质性突破。", ""),
            ("管理层进一步明确航天领域的发展方向。", ""),
        ]),
        ("03", "市场部使命", PRIMARY, [
            "更早发现需求与项目窗口。",
            "更快组织产品、方案和技术资源。",
            "把单点突破转化为可持续的市场机会。",
        ]),
    ]
    gap = Inches(0.28)
    cw = (CONTENT_W - gap * 2) / 3
    top = cur
    ch = Inches(3.55)
    railY = top + ch + Inches(0.28)
    connector_line(s, MARGIN, railY, SLIDE_W - MARGIN, railY, color=BORDER, weight=1.0)
    for i, (num, name, accent, items) in enumerate(cols):
        left = MARGIN + i * (cw + gap)
        card(s, left, top, cw, ch)
        rect(s, left, top, cw, Inches(0.06), fill=accent)
        badge_num(s, left + Inches(0.36), top + Inches(0.42), Inches(0.19), num,
                  fill=accent, size=11)
        text(s, left + Inches(0.62), top + Inches(0.24), cw - Inches(0.8), Inches(0.4),
             name, size=15.5, bold=True, color=INK, font=DISPLAY_FONT, space_after=0)
        bullets(s, left + Inches(0.3), top + Inches(0.9), cw - Inches(0.6), ch - Inches(1.1),
                items, size=11.5, marker_color=accent)
        cx = left + cw / 2
        oval(s, cx, railY, Inches(0.07), fill=accent)
    bottom_conclusion(s, "市场部不是材料制作部门，而是公司的经营前端和市场资产运营者。")
    footer(s, 2)


# ============================================================ PAGE 03
def page03():
    s = new_slide()
    cur = header(s, 3, None,
                 "市场部负责 MTL：把市场信息逐步培育为可移交的合格线索",
                 "市场部负责信息归集、线索识别、持续培育和价值判断；线索移交后的机会推进、"
                 "投标签约与回款由销售及相关部门承接。", title_size=20)

    stations = [
        ("01", "信息归集", "政策、客户、伙伴、标讯、行业信息统一入池", "可追溯信息"),
        ("02", "线索识别", "判断客户、需求方向和适配性", "初步线索"),
        ("03", "线索培育", "补充预算、范围、方案、时间窗和下一步", "培育线索／合格线索"),
        ("04", "价值判断", "判断优先级、资源条件和移交建议", "重点机会／移交件"),
    ]
    top = cur
    ch = Inches(1.7)
    sw = Inches(2.08)
    gap = Inches(0.16)
    rail_y = top + Inches(0.34)
    for i in range(4):
        left = MARGIN + i * (sw + gap)
        card(s, left, top, sw, ch, fill=PAPER_2)
        badge_num(s, left + Inches(0.34), top + Inches(0.34), Inches(0.19), stations[i][0],
                  fill=PRIMARY)
        text(s, left + Inches(0.6), top + Inches(0.2), sw - Inches(0.8), Inches(0.32),
             stations[i][1], size=13.5, bold=True, color=INK, font=DISPLAY_FONT, space_after=0)
        text(s, left + Inches(0.24), top + Inches(0.66), sw - Inches(0.48), Inches(0.62),
             stations[i][2], size=10, color=INK_2, space_after=0, line_spacing=1.2)
        chip(s, left + Inches(0.24), top + Inches(1.32), sw - Inches(0.48), Inches(0.28),
             stations[i][3], fill=COMPLETE_SOFT, text_color=COMPLETE, size=9.5)
        if i < 3:
            connector_line(s, left + sw, top + ch / 2, left + sw + gap, top + ch / 2,
                            color=PRIMARY, weight=1.6)

    gate_x = MARGIN + 4 * sw + 3 * gap + Inches(0.12)
    connector_line(s, gate_x, top - Inches(0.05), gate_x, top + ch + Inches(0.05),
                    color=INK_3, weight=1.2, dash='dash')
    text(s, gate_x - Inches(0.02), top - Inches(0.30), Inches(1.9), Inches(0.26),
         "MTL/LTC 工作界面", size=9, color=INK_3, font=MONO_FONT, space_after=0)
    outline_labels = ["机会", "投标", "合同", "回款"]
    ow = Inches(0.68)
    for j, lab in enumerate(outline_labels):
        oleft = gate_x + Inches(0.16) + j * (ow + Inches(0.08))
        rect(s, oleft, top + Inches(0.55), ow, Inches(0.6), fill=None, line=BORDER,
             line_w=1.0, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
        text(s, oleft, top + Inches(0.55), ow, Inches(0.6), lab, size=10.5, color=INK_3,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE, space_after=0)

    lvl_top = top + ch + Inches(0.3)
    levels = [
        ("初步线索", "客户、需求方向、信息来源明确。"),
        ("培育线索", "已有跟进动作，关键经营信息仍需补充。"),
        ("合格线索", "预算或资金来源、采购范围、时间窗、金额区间基本明确。"),
        ("重点机会", "竞争格局、决策关系、下一步和责任分工清晰。"),
    ]
    text(s, MARGIN, lvl_top, Inches(2), Inches(0.28), "线索等级", size=12, bold=True,
         color=PRIMARY, font=DISPLAY_FONT, space_after=0)
    lvl_top2 = lvl_top + Inches(0.32)
    lw = (CONTENT_W - Inches(0.3) * 3) / 4
    for i, (name, desc) in enumerate(levels):
        left = MARGIN + i * (lw + Inches(0.3))
        card(s, left, lvl_top2, lw, Inches(0.88), fill=PAPER_2)
        text(s, left + Inches(0.16), lvl_top2 + Inches(0.1), lw - Inches(0.32), Inches(0.26),
             name, size=11.5, bold=True, color=PRIMARY_DEEP, space_after=0)
        text(s, left + Inches(0.16), lvl_top2 + Inches(0.36), lw - Inches(0.32), Inches(0.48),
             desc, size=9, color=INK_2, line_spacing=1.15, space_after=0)

    note_top = lvl_top2 + Inches(1.0)
    rect(s, MARGIN, note_top, CONTENT_W, Inches(0.34), fill=RISK_SOFT)
    text(s, MARGIN + Inches(0.18), note_top, CONTENT_W - Inches(0.36), Inches(0.34),
         "工作界面：市场部承诺线索质量；接收、回流和状态维护规则提请公司确认。",
         size=10.5, color=RISK, bold=False, anchor=MSO_ANCHOR.MIDDLE, space_after=0)
    footer(s, 3)


# ============================================================ PAGE 04
def page04():
    s = new_slide()
    cur = header(s, 4, None, "5 月建立运行规则，6 月开始用真实任务检验流程",
                 "两个月完成制度、训练、六池和调度机制建设，并开始通过实际线索和标讯检验"
                 "经营前端流程。", title_size=22)

    top = cur
    ch = Inches(2.55)
    colw = (CONTENT_W - Inches(0.4)) / 2
    tracks = [
        ("5 月", "建立规则和基本能力", PRIMARY, [
            "完成市场部基础制度、职责和工作流程。",
            "组织公司级销售初级训练营。",
            "形成统一课件、六类场景作战卡和客户地图。",
        ]),
        ("6 月", "建立数据和经营资产", COMPLETE, [
            "归集并处理市场信息和标讯。",
            "建立六类经营资产框架，六池建设任务完成 44/50。",
            "建立标前评估和 Go/No-Go 判断机制。",
            "将 17 项临时重要事项纳入统一动态调度。",
        ]),
    ]
    for i, (mon, sub, accent, items) in enumerate(tracks):
        left = MARGIN + i * (colw + Inches(0.4))
        card(s, left, top, colw, ch)
        railx = left + Inches(0.34)
        connector_line(s, railx, top + Inches(0.3), railx, top + ch - Inches(0.22),
                        color=accent, weight=1.8)
        text(s, left + Inches(0.55), top + Inches(0.16), colw - Inches(0.8), Inches(0.34),
             mon, size=18, bold=True, color=accent, font=DISPLAY_FONT, space_after=0)
        text(s, left + Inches(0.55), top + Inches(0.5), colw - Inches(0.8), Inches(0.28),
             sub, size=11.5, color=INK_2, space_after=0)
        n = len(items)
        step = (ch - Inches(1.0)) / n
        for j, it in enumerate(items):
            ny = top + Inches(0.95) + step * j
            oval(s, railx, ny + Inches(0.09), Inches(0.055), fill=accent)
            bullets(s, left + Inches(0.55), ny, colw - Inches(0.8), step, [it],
                    size=11, marker_color=accent, space_after=0)

    chip_top = top + ch + Inches(0.28)
    chips = ["六池存数据", "月例会定优先级", "任务平台跟责任节点", "知识平台沉淀模板与版本"]
    cw2 = (CONTENT_W - Inches(0.2) * 3) / 4
    for i, c in enumerate(chips):
        left = MARGIN + i * (cw2 + Inches(0.2))
        chip(s, left, chip_top, cw2, Inches(0.42), c, fill=ASSET_SOFT, text_color=ASSET,
             size=10)

    bottom_conclusion(s, "上半年的关键成果，是让市场工作从分散响应进入统一流程，"
                          "并产生了第一批可复盘数据。")
    footer(s, 4)


# ============================================================ PAGE 05
def page05():
    s = new_slide()
    cur = header(s, 5, None, "上半年形成 9 条有效线索，来源结构和成熟度已初步可见",
                 "线索已经开始产生，但来源集中、成熟度偏低，尚未形成金额判断和规范移交。",
                 title_size=22)

    top = cur
    left_w = Inches(6.55)
    right_left = MARGIN + left_w + Inches(0.35)
    right_w = CONTENT_W - left_w - Inches(0.35)
    ch = Inches(3.3)

    card(s, MARGIN, top, left_w, ch)
    text(s, MARGIN + Inches(0.24), top + Inches(0.16), left_w - Inches(0.48), Inches(0.28),
         "线索来源结构", size=12.5, bold=True, color=PRIMARY, font=DISPLAY_FONT,
         space_after=0)
    chart_data = CategoryChartData()
    chart_data.categories = ["公司领导或总经办转入", "市场部主动开发", "既有客户或销售转入",
                              "合作伙伴或上下游企业", "政策、协会或政府渠道",
                              "招投标公开信息转化"]
    chart_data.add_series("有效线索数", (4, 3, 1, 1, 0, 0))
    gframe = s.shapes.add_chart(XL_CHART_TYPE.BAR_CLUSTERED,
                                 E(MARGIN + Inches(0.15)), E(top + Inches(0.52)),
                                 E(left_w - Inches(0.3)), E(ch - Inches(0.7)), chart_data)
    chart = gframe.chart
    chart.has_legend = False
    chart.has_title = False
    plot = chart.plots[0]
    plot.has_data_labels = True
    plot.gap_width = 60
    dl = plot.data_labels
    dl.font.size = Pt(10)
    dl.font.color.rgb = INK
    dl.font.name = MONO_FONT
    series = plot.series[0]
    series.format.fill.solid()
    series.format.fill.fore_color.rgb = PRIMARY
    cat_ax = chart.category_axis
    cat_ax.tick_labels.font.size = Pt(9.5)
    cat_ax.tick_labels.font.name = BODY_FONT
    cat_ax.format.line.color.rgb = BORDER
    val_ax = chart.value_axis
    val_ax.visible = False
    val_ax.has_major_gridlines = False
    text(s, MARGIN + Inches(0.24), top + ch - Inches(0.02), left_w - Inches(0.48), Inches(0.0),
         "", size=1, space_after=0)  # spacer no-op

    note_y = top + ch + Inches(0.06)
    rect(s, MARGIN, note_y, left_w, Inches(0.34), fill=None)
    text(s, MARGIN + Inches(0.05), note_y, left_w, Inches(0.34),
         "77% 的有效线索来自管理层导入和市场部主动开发。", size=11, bold=True,
         color=PRIMARY_DEEP, space_after=0)

    # right: maturity + advance-actions stack
    text(s, right_left, top, right_w, Inches(0.28), "线索成熟度", size=12.5, bold=True,
         color=PRIMARY, font=DISPLAY_FONT, space_after=0)
    seg_top = top + Inches(0.38)
    seg_h = Inches(0.5)
    total = 9
    segs = [
        ("5 条", "初步培育，信息仍不完整", 5, PRIMARY_SOFT, PRIMARY_DEEP),
        ("3 条", "市场部持续主跟进", 3, COMPLETE_SOFT, COMPLETE),
        ("1 条", "事实移交北京分公司主跟进（正式移交记录待补）", 1, RISK_SOFT, RISK),
    ]
    y = seg_top
    for label, desc, n, fill, tc in segs:
        h = seg_h + Inches(0.28) * (0 if len(desc) < 16 else 1)
        rect(s, right_left, y, right_w, h, fill=fill, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
        text(s, right_left + Inches(0.18), y, Inches(0.9), h, label, size=15, bold=True,
             color=tc, font=MONO_FONT, anchor=MSO_ANCHOR.MIDDLE, space_after=0)
        text(s, right_left + Inches(1.05), y, right_w - Inches(1.2), h, desc, size=10,
             color=INK_2, anchor=MSO_ANCHOR.MIDDLE, line_spacing=1.12, space_after=0)
        y += h + Inches(0.14)

    chip(s, right_left, y + Inches(0.02), right_w, Inches(0.36),
         "9 条线索均未形成金额或金额区间判断", fill=RISK_SOFT, text_color=RISK, size=10.5)
    y += Inches(0.5)

    text(s, right_left, y, right_w, Inches(0.26), "推进动作", size=11.5, bold=True,
         color=PRIMARY, font=DISPLAY_FONT, space_after=0)
    y += Inches(0.32)
    stats = [("4 条", "已明确下一步动作和计划时间"), ("2 条", "有下一步动作，但无明确时间"),
             ("3 条", "尚未形成明确下一步动作")]
    sw = (right_w - Inches(0.16) * 2) / 3
    for i, (n, d) in enumerate(stats):
        lft = right_left + i * (sw + Inches(0.16))
        card(s, lft, y, sw, Inches(0.75), fill=PAPER_2)
        text(s, lft, y + Inches(0.06), sw, Inches(0.28), n, size=14, bold=True, color=PRIMARY,
             font=MONO_FONT, align=PP_ALIGN.CENTER, space_after=0)
        text(s, lft + Inches(0.06), y + Inches(0.36), sw - Inches(0.12), Inches(0.38), d,
             size=8, color=INK_2, align=PP_ALIGN.CENTER, line_spacing=1.05, space_after=0)

    bottom_conclusion(s, "问题不是“没有线索”，而是渠道尚未稳定、线索还不够成熟、经营价值还无法量化。")
    footer(s, 5)


# ============================================================ PAGE 06
def page06():
    s = new_slide()
    cur = header(s, 6, None, "实际运行暴露出四类问题：渠道、成熟度、价值判断和工作界面",
                 "这些问题不是对上半年工作的否定，而是下半年进行精细化运营的真实起点。",
                 title_size=22)

    top = cur
    quads = [
        ("01", "渠道结构尚未稳定", ["44% 依赖管理层转入。",
                               "政策、协会和政府渠道尚未形成有效线索。",
                               "合作伙伴渠道只有 1 条，尚未形成稳定贡献。"]),
        ("02", "线索成熟度不足", ["5/9 仍处于信息补全阶段。",
                              "5/9 缺少明确推进时间，其中 3 条没有下一步动作。"]),
        ("03", "经营价值无法量化", ["9/9 均无法估算金额。",
                                "4 条重点跟进线索仍缺预算、采购范围、方案配置和立项时间。"]),
        ("04", "工作界面没有闭环", ["1 条线索已事实移交北分，但没有正式移交和接收记录。",
                               "移交后的状态回传、回流和统计规则尚未形成。"]),
    ]
    gap = Inches(0.24)
    qw = (CONTENT_W - Inches(2.15) - gap) / 2
    qh = Inches(1.72)
    for i, (num, title_s, items) in enumerate(quads):
        r, c = divmod(i, 2)
        left = MARGIN + c * (qw + gap)
        t = top + r * (qh + gap)
        card(s, left, t, qw, qh)
        badge_num(s, left + Inches(0.32), t + Inches(0.32), Inches(0.18), num, fill=PRIMARY)
        text(s, left + Inches(0.56), t + Inches(0.16), qw - Inches(0.75), Inches(0.32),
             title_s, size=13, bold=True, color=INK, font=DISPLAY_FONT, space_after=0)
        bullets(s, left + Inches(0.3), t + Inches(0.56), qw - Inches(0.55), qh - Inches(0.7),
                items, size=10, space_after=4, line_spacing=1.12)

    gx = MARGIN + 2 * qw + gap + Inches(0.15)
    gauge_w = Inches(2.0)
    card(s, gx, top, gauge_w, qh * 2 + gap, fill=PAPER_2)
    text(s, gx + Inches(0.18), top + Inches(0.16), gauge_w - Inches(0.36), Inches(0.5),
         "标讯侧补充诊断", size=11.5, bold=True, color=PRIMARY, font=DISPLAY_FONT,
         line_spacing=1.1, space_after=0)
    gcx = gx + gauge_w / 2
    gcy = top + Inches(1.55)
    gr = Inches(0.62)
    oval(s, gcx, gcy, gr, fill=None, line=BORDER, line_w=6.0)
    # small orange wedge stub to indicate 0/6 (near-zero) via a short arc line
    connector_line(s, gcx, gcy - gr, gcx + Inches(0.12), gcy - gr + Inches(0.02),
                    color=RISK, weight=6.0)
    text(s, gx, gcy - Inches(0.24), gauge_w, Inches(0.5), "0/6", size=20, bold=True,
         color=RISK, font=MONO_FONT, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE,
         space_after=0)
    text(s, gx + Inches(0.14), gcy + gr + Inches(0.18), gauge_w - Inches(0.28), Inches(0.7),
         "6 项标讯均在公告发布后进入评估，标讯前置识别率为 0/6；具体 No-Go 原因待逐条回填。",
         size=8.5, color=INK_2, align=PP_ALIGN.CENTER, line_spacing=1.15, space_after=0)

    bottom_conclusion(s, "上半年问题的共同指向：MTL 已经开始运行，但还缺少数据完整度、"
                          "渠道质量和持续调优能力。")
    footer(s, 6)


# ============================================================ PAGE 07
def page07():
    s = new_slide()
    cur = header(s, 7, None, "上半年发现的问题，就是下半年精细化改进的依据",
                 "下半年每项动作都必须对应一个已发生的问题，并通过新的运行数据验证是否有效。",
                 title_size=22)

    rows = [
        ["上半年暴露的问题", "下半年改进动作", "验证结果"],
        ["线索来源集中、稳定渠道不足", "建立渠道分类、重点主体和伙伴运营机制",
         "各渠道有效线索数、成熟度和推进速度"],
        ["线索信息不完整、推进节奏不清", "统一字段、阶段、下一步和更新时间要求",
         "数据完整率、阶段变化、超期情况"],
        ["9 条线索无法估值", "建立金额区间、预算和价值评估规则",
         "价值评估覆盖率、管道金额基线"],
        ["事实移交但无正式记录", "建立移交、接收、回流和状态更新规则",
         "正式移交率、接收状态完整率"],
        ["标讯均在公告后识别", "跟踪重点主体、政策规划、标准和伙伴信息",
         "标讯前置识别率"],
    ]
    top = cur
    th = Inches(2.85)
    plain_table(s, MARGIN, top, CONTENT_W, th, rows, col_ratios=[1, 1, 1],
                body_font=10.5, header_font=11)

    conv_y = top + th + Inches(0.28)
    rect(s, MARGIN, conv_y, CONTENT_W, Inches(0.62), fill=PRIMARY_SOFT)
    text(s, MARGIN + Inches(0.15), conv_y, Inches(1.6), Inches(0.62), "核心转折", size=11,
         bold=True, color=PRIMARY, font=MONO_FONT, anchor=MSO_ANCHOR.MIDDLE, space_after=0)
    text(s, MARGIN + Inches(1.7), conv_y, CONTENT_W - Inches(1.9), Inches(0.62),
         "建设重点从“把表建出来”，转为“用运行数据持续发现问题、调整规则和重新配置资源”。",
         size=13, bold=True, color=PRIMARY_DEEP, anchor=MSO_ANCHOR.MIDDLE, line_spacing=1.15,
         space_after=0)
    footer(s, 7)


# ============================================================ PAGE 08
def page08():
    s = new_slide()
    cur = header(s, 8, None, "下半年推动 MTL 进入“数据—复盘—调优—再运行”的闭环",
                 "精细化运营不是增加台账，而是让每一轮运行都产生数据，每一次复盘都能够改变"
                 "下一轮规则和资源投入。", title_size=21)

    top = cur
    steps = ["信息入池", "线索分级", "跟进运行", "结果记录", "月度复盘", "规则调优"]
    descs = ["统一来源、时间和证据。", "初步、培育、合格、重点机会。",
             "责任人、下一步、计划时间和更新记录。", "推进、停滞、移交、回流或关闭。",
             "分析渠道质量、成熟度变化和问题原因。", "调整字段、分级标准、渠道投入和协同方式。"]
    loop_left = MARGIN
    loop_w = Inches(7.6)
    loop_h = Inches(3.9)
    cx = loop_left + loop_w / 2
    cy = top + loop_h / 2 + Inches(0.1)
    rx = Inches(3.05)
    ry = Inches(1.48)
    node_w, node_h = Inches(1.95), Inches(0.92)
    pts = []
    for i in range(6):
        deg = -90 + i * 60
        x, y = polar(cx, cy, rx, deg)
        y2 = Emu(int(cy + ry * math.sin(math.radians(deg))))
        pts.append((x, y2))
    for i in range(6):
        x1, y1 = pts[i]
        x2, y2 = pts[(i + 1) % 6]
        connector_line(s, x1, y1, x2, y2, color=PRIMARY_SOFT, weight=1.6)
    center_r = Inches(0.62)
    oval(s, cx, cy, center_r, fill=PRIMARY, line=None)
    text(s, cx - center_r, cy - Inches(0.14), center_r * 2, Inches(0.3), "MTL",
         size=13, bold=True, color=PAPER, font=MONO_FONT, align=PP_ALIGN.CENTER,
         space_after=0)
    text(s, cx - center_r, cy + Inches(0.06), center_r * 2, Inches(0.24), "持续调优",
         size=8.5, color=PAPER, align=PP_ALIGN.CENTER, space_after=0)
    for i, (x, y) in enumerate(pts):
        left = Emu(int(x - node_w / 2))
        t = Emu(int(y - node_h / 2))
        card(s, left, t, node_w, node_h, fill=PAPER_2)
        badge_num(s, left + Inches(0.22), t + Inches(0.22), Inches(0.15), i + 1, fill=COMPLETE,
                  size=9.5)
        text(s, left + Inches(0.42), t + Inches(0.09), node_w - Inches(0.55), Inches(0.26),
             steps[i], size=11, bold=True, color=INK, font=DISPLAY_FONT, space_after=0)
        text(s, left + Inches(0.16), t + Inches(0.42), node_w - Inches(0.3), Inches(0.46),
             descs[i], size=8, color=INK_2, line_spacing=1.1, space_after=0)

    right_left = loop_left + loop_w + Inches(0.3)
    right_w = CONTENT_W - loop_w - Inches(0.3)
    text(s, right_left, top, right_w, Inches(0.28), "运行载体", size=12.5, bold=True,
         color=PRIMARY, font=DISPLAY_FONT, space_after=0)
    carriers = [("六池", "结构化经营数据"), ("市场月例会", "诊断和优先级调整"),
                ("任务平台", "责任、节点和状态"), ("知识平台", "模板、版本、复盘和规则沉淀")]
    cy2 = top + Inches(0.4)
    ch2 = Inches(0.82)
    for i, (name, desc) in enumerate(carriers):
        t = cy2 + i * (ch2 + Inches(0.1))
        card(s, right_left, t, right_w, ch2, fill=ASSET_SOFT, line=None)
        text(s, right_left + Inches(0.18), t + Inches(0.08), right_w - Inches(0.36),
             Inches(0.26), name, size=11.5, bold=True, color=INK, font=DISPLAY_FONT,
             space_after=0)
        text(s, right_left + Inches(0.18), t + Inches(0.36), right_w - Inches(0.36),
             Inches(0.4), desc, size=9, color=INK_2, line_spacing=1.15, space_after=0)

    bottom_conclusion(s, "台账记录事实，复盘发现规律，调优改变下一轮行动。")
    footer(s, 8)


# ============================================================ PAGE 09
def page09():
    s = new_slide()
    cur = header(s, 9, None, "先把每条线索记录清楚，才能判断问题出在哪里",
                 "下半年不追求虚高线索数量，先保证每条线索有来源、有阶段、有动作、有时间、"
                 "有价值判断。", title_size=21)

    top = cur
    left_w = Inches(5.6)
    right_left = MARGIN + left_w + Inches(0.35)
    right_w = CONTENT_W - left_w - Inches(0.35)
    ch = Inches(4.05)

    card(s, MARGIN, top, left_w, ch, fill=PAPER_2)
    rect(s, MARGIN, top, left_w, Inches(0.4), fill=PRIMARY, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    rect(s, MARGIN, top + Inches(0.2), left_w, Inches(0.2), fill=PRIMARY)
    text(s, MARGIN + Inches(0.2), top, left_w - Inches(0.4), Inches(0.4), "线索对象详情页 · 字段标准",
         size=11.5, bold=True, color=PAPER, anchor=MSO_ANCHOR.MIDDLE, space_after=0)
    fields = ["来源渠道与原始证据", "客户、需求方向和适配产品", "当前阶段与阶段变化日期",
              "责任人、下一步动作和计划时间", "预算或资金来源、采购范围和立项状态",
              "金额或金额区间；无法估值时记录原因和补全动作", "移交、接收、回流或关闭状态"]
    fy = top + Inches(0.56)
    step = (ch - Inches(0.7)) / len(fields)
    for i, f in enumerate(fields):
        yy = fy + step * i
        oval(s, MARGIN + Inches(0.3), yy + step / 2, Inches(0.05), fill=COMPLETE)
        text(s, MARGIN + Inches(0.46), yy, left_w - Inches(0.7), step, f, size=10.5,
             color=INK, anchor=MSO_ANCHOR.MIDDLE, line_spacing=1.1, space_after=0)
        if i < len(fields) - 1:
            connector_line(s, MARGIN + Inches(0.2), yy + step, MARGIN + left_w - Inches(0.2),
                            yy + step, color=BORDER, weight=0.6)

    text(s, right_left, top, right_w, Inches(0.26), "三项管理规则", size=12.5, bold=True,
         color=PRIMARY, font=DISPLAY_FONT, space_after=0)
    rules = ["无下一步，不算持续跟进。", "无计划时间，必须进入月度异常清单。",
             "无法估值可以存在，但必须说明缺什么、由谁补、何时再判断。"]
    ry = top + Inches(0.38)
    for i, r in enumerate(rules):
        rh = Inches(0.58)
        card(s, right_left, ry, right_w, rh, fill=PRIMARY_SOFT, line=None)
        badge_num(s, right_left + Inches(0.28), ry + rh / 2, Inches(0.15), i + 1, fill=PRIMARY,
                  size=9.5)
        text(s, right_left + Inches(0.5), ry, right_w - Inches(0.66), rh, r, size=10.5,
             bold=True, color=PRIMARY_DEEP, anchor=MSO_ANCHOR.MIDDLE, line_spacing=1.1,
             space_after=0)
        ry += rh + Inches(0.12)

    ry += Inches(0.1)
    text(s, right_left, ry, right_w, Inches(0.26), "阶段目标", size=12.5, bold=True,
         color=PRIMARY, font=DISPLAY_FONT, space_after=0)
    ry += Inches(0.34)
    goals = [("7 月", "完成 9 条存量线索的数据回填和重新分级"),
             ("8 月", "形成第一版线索金额及成熟度基线"),
             ("9 月起", "所有新增线索按统一规则运行")]
    gh = Inches(0.62)
    for label, desc in goals:
        card(s, right_left, ry, right_w, gh, fill=PAPER_2)
        text(s, right_left + Inches(0.16), ry + Inches(0.06), Inches(0.9), Inches(0.26),
             label, size=11.5, bold=True, color=COMPLETE, font=MONO_FONT, space_after=0)
        text(s, right_left + Inches(0.16), ry + Inches(0.32), right_w - Inches(0.32),
             Inches(0.26), desc, size=9, color=INK_2, space_after=0)
        ry += gh + Inches(0.1)

    footer(s, 9)


# ============================================================ PAGE 10
def page10():
    s = new_slide()
    cur = header(s, 10, None, "不只看渠道带来多少线索，更要看线索质量和推进结果",
                 "渠道投入依据真实产出动态调整：比较不同渠道的有效率、成熟度、推进速度、"
                 "价值信息和移交结果。", title_size=21)

    top = cur
    rows = [
        ["评价维度", "回答的问题"],
        ["有效线索数", "渠道有没有产出"],
        ["成熟度分布", "带来的只是信息，还是可推进机会"],
        ["下一步及时率", "线索能不能进入行动"],
        ["价值评估覆盖率", "能不能判断金额和优先级"],
        ["移交与后续状态", "线索有没有被组织承接"],
    ]
    th = Inches(2.15)
    plain_table(s, MARGIN, top, CONTENT_W, th, rows, col_ratios=[1, 2.6], body_font=11,
                header_font=11.5)

    tr_top = top + th + Inches(0.26)
    text(s, MARGIN, tr_top, Inches(4), Inches(0.26), "下半年重点渠道动作", size=12, bold=True,
         color=PRIMARY, font=DISPLAY_FONT, space_after=0)
    channels = [
        ("管理层转入", "规范承接和反馈，确保重要线索及时形成动作。"),
        ("市场部主动开发", "复盘 3 条线索来源，形成可复制的触达方式。"),
        ("销售与既有客户", "建立统一线索入口和双向状态反馈。"),
        ("合作伙伴", "明确价值交换内容、唯一对接人和真实线索产出。"),
        ("政策与前置信息", "跟踪重点主体、政策规划和公开标准，改善 0/6 的标讯前置识别。"),
    ]
    tr_top2 = tr_top + Inches(0.32)
    cw = (CONTENT_W - Inches(0.15) * 4) / 5
    ch2 = Inches(1.0)
    station_y = tr_top2 + ch2 + Inches(0.16)
    rect(s, MARGIN, station_y, CONTENT_W, Inches(0.38), fill=PRIMARY,
         shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    text(s, MARGIN, station_y, CONTENT_W, Inches(0.38), "渠道质量评估台", size=11.5,
         bold=True, color=PAPER, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE,
         space_after=0)
    for i, (name, desc) in enumerate(channels):
        left = MARGIN + i * (cw + Inches(0.15))
        card(s, left, tr_top2, cw, ch2, fill=PAPER_2)
        text(s, left + Inches(0.1), tr_top2 + Inches(0.06), cw - Inches(0.2), Inches(0.26),
             name, size=10, bold=True, color=INK, space_after=0)
        text(s, left + Inches(0.1), tr_top2 + Inches(0.32), cw - Inches(0.2), Inches(0.62),
             desc, size=7.8, color=INK_2, line_spacing=1.1, space_after=0)
        connector_line(s, left + cw / 2, tr_top2 + ch2, left + cw / 2, station_y,
                        color=BORDER, weight=1.0)

    bottom_conclusion(s, "高质量渠道增加投入；长期只有信息、没有推进的渠道降低优先级；"
                          "渠道规则按月复盘、按季度调整。",
                       top=station_y + Inches(0.34))
    footer(s, 10)


# ============================================================ PAGE 11
def page11():
    s = new_slide()
    cur = header(s, 11, None, "每月复盘运行，每季度调整规则，年底形成可验证的 MTL 能力",
                 "下半年验收的不是表格数量，而是数据完整、问题可追溯、规则有调整、线索有进展。",
                 title_size=20)

    top = cur
    left_w = Inches(4.55)
    right_left = MARGIN + left_w + Inches(0.32)
    right_w = CONTENT_W - left_w - Inches(0.32)

    text(s, MARGIN, top, left_w, Inches(0.26), "月度复盘固定议题", size=12, bold=True,
         color=PRIMARY, font=DISPLAY_FONT, space_after=0)
    topics = ["新增信息和有效线索来自哪些渠道。", "线索阶段发生了什么变化。",
              "哪些线索缺动作、缺时间或长期停滞。", "哪些信息已经完成金额或价值判断。",
              "哪些线索发生移交、回流或关闭。", "标讯 No-Go 和停滞原因如何分布。",
              "下个月需要调整哪些规则和资源。"]
    card(s, MARGIN, top + Inches(0.32), left_w, Inches(3.55), fill=PAPER_2)
    ty = top + Inches(0.48)
    step = Inches(3.55 - 0.3) / len(topics)
    for i, tpc in enumerate(topics):
        yy = ty + step * i
        badge_num(s, MARGIN + Inches(0.32), yy + step / 2, Inches(0.14), i + 1, fill=PRIMARY,
                  size=9)
        text(s, MARGIN + Inches(0.54), yy, left_w - Inches(0.75), step, tpc, size=10,
             color=INK, anchor=MSO_ANCHOR.MIDDLE, line_spacing=1.12, space_after=0)

    text(s, right_left, top, right_w, Inches(0.26), "里程碑", size=12, bold=True,
         color=PRIMARY, font=DISPLAY_FONT, space_after=0)
    rows = [
        ["时间", "交付"],
        ["7 月底", "9 条存量线索回填、重新分级；6 项 No-Go 原因完成分类"],
        ["8 月底", "第一版线索金额与成熟度基线；工作界面试运行"],
        ["9 月起", "月度渠道质量和前置识别指标可测"],
        ["Q4  v1.0→v1.2", "完成至少 2-3 轮规则或资源调优，并保留调整记录"],
        ["年底", "形成 MTL 年度运行报告和 2027 目标建议"],
    ]
    mh = Inches(1.95)
    plain_table(s, right_left, top + Inches(0.32), right_w, mh, rows,
                col_ratios=[1, 2.6], body_font=9.3, header_font=10)

    deliv_top = top + Inches(0.32) + mh + Inches(0.22)
    text(s, right_left, deliv_top, right_w, Inches(0.24), "年底六项交付", size=11.5,
         bold=True, color=PRIMARY, font=DISPLAY_FONT, space_after=0)
    delivs = ["MTL 数据字典和线索分级规则", "渠道质量与贡献分析", "线索成熟度和金额基线",
              "标讯前置识别和 No-Go 原因分析", "MTL/LTC 移交与回流运行记录",
              "2-3 轮调优记录及对应效果"]
    dtop = deliv_top + Inches(0.3)
    dw = (right_w - Inches(0.12)) / 2
    dh = Inches(0.42)
    for i, d in enumerate(delivs):
        r, c = divmod(i, 2)
        left = right_left + c * (dw + Inches(0.12))
        yy = dtop + r * (dh + Inches(0.1))
        chip(s, left, yy, dw, dh, d, fill=ASSET_SOFT, text_color=INK, bold=False, size=8.3)

    bottom_conclusion(s, "年底不仅知道“有多少线索”，还要知道线索为什么来、为什么停、怎样提升。")
    footer(s, 11)


# ============================================================ PAGE 12
def page12():
    s = new_slide()
    cur = header(s, 12, None, "请管理层确认运行规则，市场部承诺按月交付真实数据和改进结果",
                 "精细化运营需要明确的职责边界和工作界面；规则确认后，市场部对数据质量、"
                 "复盘节奏和调优结果负责。", title_size=19)

    top = cur
    rows = [
        ["请管理层确认", "市场部对等承诺"],
        ["市场部 MTL 职责边界和线索等级定义", "统一入口、统一分级，数据依据可追溯"],
        ["MTL/LTC 移交、接收、回流和状态维护规则", "按等级提供移交件，未成熟线索继续培育"],
        ["产品、工程、交付的专业审核接口", "使用标准输入、提供合理提前量、记录审核结论"],
        ["北京、青岛、辽宁等区域协同口径", "不介入既有客户归属，聚焦线索供给和专业支撑"],
        ["2026 H2 以精细化运行交付为主要评价依据",
         "月度报告，年底提交运行基线、调优记录和 2027 目标建议"],
    ]
    th = Inches(2.55)
    gtable = s.shapes.add_table(len(rows), 2, E(MARGIN), E(top), E(CONTENT_W), E(th)).table
    gtable.first_row = False
    gtable.horz_banding = False
    gtable.columns[0].width = Emu(int(CONTENT_W * 0.5))
    gtable.columns[1].width = Emu(int(CONTENT_W * 0.5))
    for r in range(len(rows)):
        for c in range(2):
            cell = gtable.cell(r, c)
            cell.margin_left = Pt(10)
            cell.margin_right = Pt(10)
            cell.margin_top = Pt(5)
            cell.margin_bottom = Pt(5)
            cell.vertical_anchor = MSO_ANCHOR.MIDDLE
            cell.fill.solid()
            if r == 0:
                cell.fill.fore_color.rgb = PRIMARY if c == 0 else COMPLETE
                fc = PAPER
            else:
                cell.fill.fore_color.rgb = PRIMARY_SOFT if c == 0 else COMPLETE_SOFT
                fc = INK
            tf = cell.text_frame
            tf.word_wrap = True
            p = tf.paragraphs[0]
            run = p.add_run()
            run.text = str(rows[r][c])
            run.font.size = Pt(11.5 if r == 0 else 10)
            run.font.bold = (r == 0)
            run.font.name = BODY_FONT
            run.font.color.rgb = fc
    for r in range(len(rows)):
        for c in range(2):
            add_cell_borders(gtable.cell(r, c), color_hex='FCFBF8')

    prom_top = top + th + Inches(0.16)
    text(s, MARGIN, prom_top, CONTENT_W, Inches(0.24), "市场部三项明确承诺", size=11.5,
         bold=True, color=PRIMARY, font=DISPLAY_FONT, space_after=0)
    proms = ["不以虚高线索数量代替质量。", "不以临时任务多为完成率辩护，所有重要事项进入统一调度。",
             "不只报告问题，每次复盘必须形成责任、动作和规则调整。"]
    pw = (CONTENT_W - Inches(0.2) * 2) / 3
    py = prom_top + Inches(0.3)
    ph = Inches(0.56)
    for i, pr in enumerate(proms):
        left = MARGIN + i * (pw + Inches(0.2))
        card(s, left, py, pw, ph, fill=PAPER_2)
        badge_num(s, left + Inches(0.24), py + ph / 2, Inches(0.14), i + 1, fill=RISK, size=9)
        text(s, left + Inches(0.44), py, pw - Inches(0.6), ph, pr, size=9.3, bold=True,
             color=INK, anchor=MSO_ANCHOR.MIDDLE, line_spacing=1.1, space_after=0)

    close_top = py + ph + Inches(0.16)
    rect(s, MARGIN, close_top, CONTENT_W, Inches(0.5), fill=PRIMARY)
    text(s, MARGIN + Inches(0.24), close_top, CONTENT_W - Inches(0.48), Inches(0.5),
         "上半年让问题变得可见，下半年让改进变得可验证。", size=13.5, bold=True, color=PAPER,
         anchor=MSO_ANCHOR.MIDDLE, space_after=0)

    dec_top = close_top + Inches(0.62)
    fields = ["责任主体：___________________", "完成时间：___________________",
              "验证方式：___________________"]
    fw = (CONTENT_W - Inches(0.2) * 2) / 3
    for i, f in enumerate(fields):
        left = MARGIN + i * (fw + Inches(0.2))
        text(s, left, dec_top, fw, Inches(0.3), f, size=10, color=INK_2, font=MONO_FONT,
             space_after=0)
    footer(s, 12)


for fn in (page01, page02, page03, page04, page05, page06, page07, page08, page09,
           page10, page11, page12):
    fn()

import os
OUT = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                    "遨海科技市场部2026年半年工作总结.pptx")
prs.save(OUT)
print("saved", OUT)
