# -*- coding: utf-8 -*-
"""
build_full_counseling_deck.py
中投區國三升學志願選填專家深度診斷與戰略指南 (70吋電視滿版大字・三風格簡報生成器)
- 左側 45%：OCR 乾淨大字表（字大滿版）
- 右側 55%：只保留「人物故事梗概」並放大滿版（粗體鮮明）
- 專家急救處方頁：文字精簡粗體鮮明、滿版色塊卡片化分離、落實「3秒原則」
- 同步生成：
  1. 風格 1：高對比極簡看板風 (中投區國三志願選填_風格1_高對比極簡看板風.pptx / .pdf)
  2. 風格 2：沉浸式暗色科技風 (中投區國三志願選填_風格2_沉浸式暗色科技風.pptx / .pdf)
  3. 風格 3：大色塊幾何雜誌風 (中投區國三志願選填_風格3_大色塊幾何雜誌風.pptx / .pdf)
  4. 三風格黃金融合旗艦版 (中投區國三升學志願選填分析_專家深度診斷與策略指南_116版.pptx / .pdf)
"""

import os
import sys
import pptx
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
import comtypes.client

sys.path.insert(0, os.path.abspath('.'))
sys.stdout.reconfigure(encoding='utf-8')

from student_data import STUDENTS_DATA

FONT_NAME = "Microsoft JhengHei"

# ==============================================================================
# 三大視覺風格主題色票定義 (Themes for Style 1, Style 2, Style 3)
# ==============================================================================
THEMES = {
    "style1_billboard": {
        "name": "風格1：高對比極簡看板風",
        "slide_bg": RGBColor(254, 249, 195),       # 暖黃高對比底 (#FEF9C3)
        "slide_bg_alt": RGBColor(255, 255, 255),   # 純白高對比底
        "cover_bg": RGBColor(250, 204, 21),        # 強烈看板亮黃 (#FACC15)
        "cover_text": RGBColor(15, 23, 42),        # 深海藍炭黑 (#0F172A)
        "cover_sub": RGBColor(30, 41, 59),
        "header_bg": RGBColor(15, 23, 42),         # 頂部深海藍看板橫幅
        "header_tag": RGBColor(250, 204, 21),      # 亮黃標籤
        "header_title": RGBColor(255, 255, 255),   # 純白大標
        "card_bg_1": RGBColor(255, 255, 255),      # 純白看板卡
        "card_bg_2": RGBColor(254, 240, 138),      # 飽和暖黃看板卡
        "card_bg_3": RGBColor(224, 242, 254),      # 高對比亮藍看板卡
        "card_bg_4": RGBColor(254, 226, 226),      # 高對比亮紅看板卡
        "card_bg_5": RGBColor(209, 250, 229),      # 高對比亮綠看板卡
        "border_col": RGBColor(15, 23, 42),        # 粗黑看板邊框
        "border_width": 2.5,
        "text_main": RGBColor(15, 23, 42),         # 深海藍炭黑
        "text_sub": RGBColor(30, 41, 59),
        "accent_primary": RGBColor(29, 78, 216),   # 飽和寶藍
        "accent_danger": RGBColor(220, 38, 38),    # 飽和警示紅
        "accent_success": RGBColor(4, 120, 87),    # 飽和翠綠
        "accent_warn": RGBColor(180, 83, 9),       # 飽和琥珀
        "table_hdr_bg": RGBColor(15, 23, 42),
        "table_hdr_fg": RGBColor(250, 204, 21),
        "table_row_even": RGBColor(255, 255, 255),
        "table_row_odd": RGBColor(254, 252, 232),
        "table_row_trap": RGBColor(254, 202, 202),
        "table_trap_fg": RGBColor(185, 28, 28),
        "banner_bg": RGBColor(15, 23, 42),
        "banner_fg": RGBColor(250, 204, 21),
    },
    "style2_darktech": {
        "name": "風格2：沉浸式暗色科技風",
        "slide_bg": RGBColor(30, 30, 30),          # 深石墨灰 (#1E1E1E)
        "slide_bg_alt": RGBColor(24, 24, 27),      # 暗曜石底 (#18181B)
        "cover_bg": RGBColor(18, 18, 20),          # 極致暗黑科技底
        "cover_text": RGBColor(0, 240, 255),       # 電光青 (#00F0FF)
        "cover_sub": RGBColor(255, 159, 28),       # 亮琥珀橘 (#FF9F1C)
        "header_bg": RGBColor(39, 39, 42),         # 暗色科技頂欄
        "header_tag": RGBColor(0, 240, 255),       # 電光青
        "header_title": RGBColor(255, 255, 255),   # 純白特粗
        "card_bg_1": RGBColor(42, 42, 46),         # 暗灰科技面板1
        "card_bg_2": RGBColor(35, 43, 53),         # 暗青科技面板2
        "card_bg_3": RGBColor(48, 36, 42),         # 暗紅科技面板3
        "card_bg_4": RGBColor(32, 46, 40),         # 暗綠科技面板4
        "card_bg_5": RGBColor(46, 40, 30),         # 暗橘科技面板5
        "border_col": RGBColor(0, 240, 255),       # 霓虹電光青外框
        "border_width": 2.0,
        "text_main": RGBColor(255, 255, 255),      # 純白高對比字
        "text_sub": RGBColor(226, 232, 240),       # 亮銀灰字
        "accent_primary": RGBColor(0, 240, 255),   # 電光青
        "accent_danger": RGBColor(255, 77, 109),   # 霓虹洋紅
        "accent_success": RGBColor(52, 211, 153),  # 霓虹萊姆綠
        "accent_warn": RGBColor(255, 159, 28),     # 亮琥珀橘
        "table_hdr_bg": RGBColor(0, 180, 216),     # 電光藍表頭
        "table_hdr_fg": RGBColor(15, 23, 42),
        "table_row_even": RGBColor(42, 42, 46),
        "table_row_odd": RGBColor(52, 52, 58),
        "table_row_trap": RGBColor(90, 24, 42),    # 暗紅霓虹警示底
        "table_trap_fg": RGBColor(255, 159, 28),   # 亮琥珀橘警示字
        "banner_bg": RGBColor(255, 159, 28),       # 亮琥珀橘底
        "banner_fg": RGBColor(18, 18, 20),         # 深黑字
    },
    "style3_colorblock": {
        "name": "風格3：大色塊幾何雜誌風",
        "slide_bg": RGBColor(241, 245, 249),       # 莫蘭迪淺灰底 (#F1F5F9)
        "slide_bg_alt": RGBColor(248, 250, 252),
        "cover_bg": RGBColor(30, 58, 138),         # 雜誌深皇家藍
        "cover_text": RGBColor(254, 243, 199),     # 奶油杏黃字
        "cover_sub": RGBColor(186, 230, 253),      # 莫蘭迪霧藍字
        "header_bg": RGBColor(30, 41, 59),         # 雜誌深石板藍
        "header_tag": RGBColor(125, 211, 252),     # 粉彩天藍
        "header_title": RGBColor(255, 255, 255),
        "card_bg_1": RGBColor(219, 234, 254),      # 莫蘭迪霧藍塊 (#DBEAFE)
        "card_bg_2": RGBColor(254, 243, 199),      # 暖杏奶油色塊 (#FEF3C7)
        "card_bg_3": RGBColor(255, 228, 230),      # 粉彩珊瑚紅色塊 (#FFE4E6)
        "card_bg_4": RGBColor(209, 250, 229),      # 薄荷鼠尾草綠色塊 (#D1FAE5)
        "card_bg_5": RGBColor(243, 232, 255),      # 莫蘭迪薰衣草紫色塊 (#F3E8FF)
        "border_col": RGBColor(255, 255, 255),     # 雜誌幾何白邊分割
        "border_width": 2.5,
        "text_main": RGBColor(15, 23, 42),         # 大號深色特粗體
        "text_sub": RGBColor(30, 41, 59),
        "accent_primary": RGBColor(3, 105, 161),    # 深海藍標籤
        "accent_danger": RGBColor(190, 18, 60),    # 深珊瑚紅標籤
        "accent_success": RGBColor(4, 120, 87),    # 深森林綠標籤
        "accent_warn": RGBColor(180, 83, 9),       # 深焦糖橘標籤
        "table_hdr_bg": RGBColor(30, 58, 138),
        "table_hdr_fg": RGBColor(255, 255, 255),
        "table_row_even": RGBColor(255, 255, 255),
        "table_row_odd": RGBColor(224, 242, 254),
        "table_row_trap": RGBColor(254, 205, 211),
        "table_trap_fg": RGBColor(159, 18, 57),
        "banner_bg": RGBColor(6, 95, 70),
        "banner_fg": RGBColor(236, 253, 245),
    }
}


def add_bg(slide, color):
    bg = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, Inches(13.333), Inches(7.5))
    bg.fill.solid()
    bg.fill.fore_color.rgb = color
    bg.line.fill.background()
    return bg


def add_block(slide, left, top, width, height, bg_color, border_color=None, border_width=2.0, rounded=False):
    shape_type = MSO_SHAPE.ROUNDED_RECTANGLE if rounded else MSO_SHAPE.RECTANGLE
    card = slide.shapes.add_shape(shape_type, left, top, width, height)
    card.fill.solid()
    card.fill.fore_color.rgb = bg_color
    if border_color and border_width > 0:
        card.line.color.rgb = border_color
        card.line.width = Pt(border_width)
    else:
        card.line.fill.background()
    return card


def add_top_bar(slide, th, tag_text, title_text, badge_text=None, is_danger=True):
    """滿版高對比頂部導航橫幅 (高度 0.92 吋)"""
    bar = add_block(slide, Inches(0.25), Inches(0.14), Inches(12.833), Inches(0.88),
                    th["header_bg"], th["border_col"], th["border_width"], rounded=False)

    tx = slide.shapes.add_textbox(Inches(0.45), Inches(0.16), Inches(8.8), Inches(0.82))
    tf = tx.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_top = tf.margin_right = tf.margin_bottom = 0

    p1 = tf.paragraphs[0]
    p1.text = tag_text
    p1.font.name = FONT_NAME
    p1.font.size = Pt(13)
    p1.font.bold = True
    p1.font.color.rgb = th["header_tag"]

    p2 = tf.add_paragraph()
    p2.text = title_text
    p2.font.name = FONT_NAME
    p2.font.size = Pt(20)
    p2.font.bold = True
    p2.font.color.rgb = th["header_title"]

    if badge_text:
        b_bg = th["accent_danger"] if is_danger else th["accent_primary"]
        b_fg = RGBColor(255, 255, 255) if th["name"] != "風格2：沉浸式暗色科技風" else RGBColor(18, 18, 20)
        if th["name"] == "風格2：沉浸式暗色科技風" and is_danger:
            b_fg = RGBColor(255, 255, 255)
        badge = add_block(slide, Inches(9.35), Inches(0.24), Inches(3.55), Inches(0.68),
                          b_bg, th["header_title"], 1.5, rounded=True)
        btf = badge.text_frame
        btf.word_wrap = True
        btf.vertical_anchor = MSO_ANCHOR.MIDDLE
        bp = btf.paragraphs[0]
        bp.alignment = PP_ALIGN.CENTER
        bp.text = badge_text
        bp.font.name = FONT_NAME
        bp.font.size = Pt(15)
        bp.font.bold = True
        bp.font.color.rgb = b_fg


def build_deck_for_mode(mode_key, out_pptx):
    """
    mode_key:
      - 'style1_billboard': 全簡報採用 風格1（高對比極簡看板風）
      - 'style2_darktech':  全簡報採用 風格2（沉浸式暗色科技風）
      - 'style3_colorblock':全簡報採用 風格3（大色塊幾何雜誌風）
      - 'hybrid_master':    三風格黃金融合旗艦版（觀念提問=風格1，比序系統=風格2，16學生案例=風格3）
    """
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    blank_layout = prs.slide_layouts[6]

    def get_theme(section_type):
        if mode_key in THEMES:
            return THEMES[mode_key]
        # hybrid_master mode maps best style to each section:
        if section_type == "concept":
            return THEMES["style1_billboard"]
        elif section_type == "system":
            return THEMES["style2_darktech"]
        else:
            return THEMES["style3_colorblock"]

    # ==========================================================================
    # SLIDE 1: 滿版大字報封面 (Concept -> Style 1 in Hybrid)
    # ==========================================================================
    th1 = get_theme("concept")
    s1 = prs.slides.add_slide(blank_layout)
    add_bg(s1, th1["cover_bg"])

    # 巨型幾何裝飾框
    frame1 = add_block(s1, Inches(0.45), Inches(0.45), Inches(12.433), Inches(6.6),
                       th1["cover_bg"], th1["cover_text"], 4.0, rounded=False)

    # 頂部標籤色塊
    top_pill = add_block(s1, Inches(0.9), Inches(0.9), Inches(6.8), Inches(0.75),
                         th1["header_bg"], None, 0, rounded=False)
    ptf = top_pill.text_frame
    ptf.vertical_anchor = MSO_ANCHOR.MIDDLE
    pp = ptf.paragraphs[0]
    pp.text = "  🎯 116學年度中投區免試入學 ｜ 70吋電視滿版大字戰略版"
    pp.font.name = FONT_NAME
    pp.font.size = Pt(18)
    pp.font.bold = True
    pp.font.color.rgb = th1["header_tag"]

    # 主標題（佔據版面三分之一以上，特粗超大字）
    tx1 = s1.shapes.add_textbox(Inches(0.9), Inches(1.85), Inches(11.5), Inches(3.4))
    tf1 = tx1.text_frame
    tf1.word_wrap = True
    p1 = tf1.paragraphs[0]
    p1.text = "選對策略，比考高分更關鍵！\n16位國三生志願選填深度診斷"
    p1.font.name = FONT_NAME
    p1.font.size = Pt(48)
    p1.font.bold = True
    p1.font.color.rgb = th1["cover_text"]
    p1.line_spacing = 1.25

    # 下方三大口號色塊列
    slogans = [
        ("⚡ 3秒看懂演算法", "13階比序＋成交不可逆", th1["card_bg_1"]),
        ("🔍 左45% OCR大字表", "10大志願門檻與下場全公開", th1["card_bg_3"]),
        ("🚀 右55% 故事＆急救處方", "坐在學生身邊的接地氣雙贏解方", th1["card_bg_5"])
    ]
    for idx, (st_t, st_d, s_bg) in enumerate(slogans):
        bx = Inches(0.9 + idx * 3.9)
        b_card = add_block(s1, bx, Inches(5.35), Inches(3.7), Inches(1.3),
                           s_bg, th1["border_col"], 2.5, rounded=False)
        btf = b_card.text_frame
        btf.word_wrap = True
        btf.vertical_anchor = MSO_ANCHOR.MIDDLE
        bp1 = btf.paragraphs[0]
        bp1.text = st_t
        bp1.font.name = FONT_NAME
        bp1.font.size = Pt(18)
        bp1.font.bold = True
        bp1.font.color.rgb = th1["text_main"]
        bp1.alignment = PP_ALIGN.CENTER
        bp2 = btf.add_paragraph()
        bp2.text = st_d
        bp2.font.name = FONT_NAME
        bp2.font.size = Pt(13.5)
        bp2.font.bold = True
        bp2.font.color.rgb = th1["text_sub"]
        bp2.alignment = PP_ALIGN.CENTER

    # ==========================================================================
    # SLIDE 2: 116 官方 13 階比序演算法 (System -> Style 2 Dark Tech in Hybrid)
    # ==========================================================================
    th2 = get_theme("system")
    s2 = prs.slides.add_slide(blank_layout)
    add_bg(s2, th2["slide_bg"])
    add_top_bar(s2, th2, "官方比序演算法 ｜ 116學年度中投區免試入學作業要點",
                "電腦分發 4 大階段 ＆ 3 大保命鐵律（3秒看懂！）", "⚡ 達標即結案不可逆", is_danger=True)

    # 左側 48%：4大比序天梯滿版色塊
    phases = [
        ("階 1 ｜ 總積分 100 分", "志願序30 ＋ 就近10 ＋ 弱勢3 ＋ 多元27 ＋ 會考30", th2["card_bg_1"], th2["accent_primary"]),
        ("階 2 ｜ 分項積分水平比", "志願序(30) ➔ 就近(10) ➔ 弱勢(3) ➔ 多元(27) ➔ 會考(30)", th2["card_bg_2"], th2["accent_warn"]),
        ("階 3 ｜ 會考核心大對決", "總積點(111點) ➔ +號總數 ➔ 單科積分(國數英社自)", th2["card_bg_3"], th2["accent_danger"]),
        ("階 4 ｜ 校序順次與增額", "校序順次 ➔ 志願序順次 ➔ 仍同分一律【增額錄取不抽籤】", th2["card_bg_4"], th2["accent_success"])
    ]
    for idx, (pt, pd, cbg, acol) in enumerate(phases):
        py = Inches(1.15 + idx * 1.54)
        add_block(s2, Inches(0.25), py, Inches(6.15), Inches(1.42), cbg, acol, 2.5)
        tx = s2.shapes.add_textbox(Inches(0.45), py + Inches(0.10), Inches(5.8), Inches(1.25))
        tf = tx.text_frame
        tf.word_wrap = True
        p1 = tf.paragraphs[0]
        p1.text = pt
        p1.font.name = FONT_NAME
        p1.font.size = Pt(20)
        p1.font.bold = True
        p1.font.color.rgb = acol
        p2 = tf.add_paragraph()
        p2.text = pd
        p2.font.name = FONT_NAME
        p2.font.size = Pt(15.5)
        p2.font.bold = True
        p2.font.color.rgb = th2["text_main"]
        p2.space_before = Pt(4)

    # 右側 52%：3大黃金戰略鐵律滿版大字卡
    rules = [
        ("🚨 鐵律 1：由上往下，達標立刻關門！",
         "電腦從第 1 志願往下掃描，只要某一志願達標立刻「成交鎖檔」！後面排再多夢幻名校，電腦連看都不看一眼！",
         th2["card_bg_3"], th2["accent_danger"]),
        ("🛡️ 鐵律 2：前 10 志願同拿 30 分滿分！",
         "第 1 到第 10 志願全部同為 30 分！前 3 志願大膽填「跳一階的公立高中或公立高職名科」，沒上往下滑完全零扣分！",
         th2["card_bg_4"], th2["accent_success"]),
        ("✨ 鐵律 3：同校連續選填不同科＝不扣分！",
         "連續填同一所高職的不同科別（如連填霧峰農工4科、或台中家商4科），全部只算「同 1 個志願序」！包牌必備神技！",
         th2["card_bg_5"], th2["accent_warn"])
    ]
    for idx, (rt, rd, cbg, acol) in enumerate(rules):
        ry = Inches(1.15 + idx * 2.06)
        add_block(s2, Inches(6.55), ry, Inches(6.533), Inches(1.94), cbg, acol, 2.5)
        tx = s2.shapes.add_textbox(Inches(6.75), ry + Inches(0.12), Inches(6.15), Inches(1.72))
        tf = tx.text_frame
        tf.word_wrap = True
        p1 = tf.paragraphs[0]
        p1.text = rt
        p1.font.name = FONT_NAME
        p1.font.size = Pt(20)
        p1.font.bold = True
        p1.font.color.rgb = acol
        p2 = tf.add_paragraph()
        p2.text = rd
        p2.font.name = FONT_NAME
        p2.font.size = Pt(15.5)
        p2.font.bold = True
        p2.font.color.rgb = th2["text_main"]
        p2.space_before = Pt(6)

    # ==========================================================================
    # SLIDE 3: 中投區各梯隊真實錄取門檻滿版大字總表 (System -> Style 2)
    # ==========================================================================
    s3 = prs.slides.add_slide(blank_layout)
    add_bg(s3, th2["slide_bg"])
    add_top_bar(s3, th2, "中投區真實行情庫 ｜ 高中＆高職完整對應",
                "中投區各梯隊學校會考等第 ＆ 積點落點大字總表", "📊 懂行情才不會亂填", is_danger=False)

    t_shape = s3.shapes.add_table(6, 4, Inches(0.25), Inches(1.12), Inches(12.833), Inches(6.18))
    t = t_shape.table
    t.columns[0].width = Inches(2.2)
    t.columns[1].width = Inches(4.1)
    t.columns[2].width = Inches(2.3)
    t.columns[3].width = Inches(4.233)

    headers = ["梯隊分級", "代表公立高中 ＆ 公立高職名科", "會考等第 (積點)", "接地氣選填實戰攻略"]
    for c_i, h in enumerate(headers):
        cell = t.cell(0, c_i)
        cell.text = h
        cell.fill.solid()
        cell.fill.fore_color.rgb = th2["table_hdr_bg"]
        cell.text_frame.vertical_anchor = MSO_ANCHOR.MIDDLE
        for p in cell.text_frame.paragraphs:
            p.font.name = FONT_NAME
            p.font.size = Pt(16)
            p.font.bold = True
            p.font.color.rgb = th2["table_hdr_fg"]
            p.alignment = PP_ALIGN.CENTER

    rows_s3 = [
        ("第一梯隊\n頂尖霸主", "台中一中 / 台中女中 (94~97點)\n興大附中 / 中科實中 (86~94點)\n台中高工 (電機/資訊/電子 71~90點)", "5A ~ 4A1B\n(71 ~ 97 點)", "一中女中需5A高積點；4A1B(86點+)全力衝刺附中與中科；高工資電是理工技職霸主！"),
        ("第二梯隊\n前段王牌", "文華高中 (男85-88點 / 女82-85點)\n惠文高中(76-77點) / 台中二中(69-74點)\n台中家商(應英/國貿) / 高工(圖傳/機械)", "4A1B ~ 2A3B\n(62 ~ 88 點)", "84-88點女生填文華極穩！惠文分數高於二中，切勿把二中排在文華與惠文前面！"),
        ("第三梯隊\n優質公立", "忠明(71-76) / 西苑 / 東山(70-74點)\n市立大里高中 (61~74點)\n興大附農(餐飲/觀光) / 家商(服飾/資處)", "3A2B ~ 1A4B\n(55 ~ 76 點)", "1A4B~2A3B(60~68點)是興大附農、台中家商與市立大里的黃金主戰場！"),
        ("第四梯隊\n公立主力", "市立長億高中 (53~64點・攔截大水壩)\n豐原高商 (50~62點)\n國立霧峰農工 / 沙鹿高工 / 草屯商工", "1A4B ~ 2B3C\n(33 ~ 64 點)", "5B或2B3C想讀工科/餐旅，前3志願務必先衝刺『霧峰農工、草屯商工』等公立高職！"),
        ("第五梯隊\n穩健保底", "新社高中 / 竹山高中 (31~42點)\n明道 / 僑泰 / 新民 / 青年 / 慈明各科", "5B ~ 5C\n(20 ~ 45 點)", "私校特色科放第6~10志願連填保底，千萬別放第1~4志願提早攔截公立機會！")
    ]
    for r_i, r_data in enumerate(rows_s3):
        for c_i, val in enumerate(r_data):
            cell = t.cell(r_i + 1, c_i)
            cell.text = val
            cell.fill.solid()
            cell.fill.fore_color.rgb = th2["table_row_even"] if r_i % 2 == 0 else th2["table_row_odd"]
            cell.text_frame.vertical_anchor = MSO_ANCHOR.MIDDLE
            for p in cell.text_frame.paragraphs:
                p.font.name = FONT_NAME
                p.font.size = Pt(14)
                p.font.bold = True
                p.font.color.rgb = th2["text_main"] if c_i != 2 else th2["accent_warn"]
                if c_i in [0, 2]:
                    p.alignment = PP_ALIGN.CENTER

    # ==========================================================================
    # SLIDE 4: 蔡幸君老師適性心法 ＆ 2:5:3 滿版看板 (Concept -> Style 1)
    # ==========================================================================
    s4 = prs.slides.add_slide(blank_layout)
    add_bg(s4, th1["slide_bg"])
    add_top_bar(s4, th1, "適性揚才核心心法 ｜ 蔡幸君老師志願選填哲學",
                "拒絕盲目跟風！落實「想好、填好」與前10志願「2:5:3」黃金比例", "💡 走對跑道人人是第一", is_danger=False)

    # 上半部左右雙看板
    b_left = add_block(s4, Inches(0.25), Inches(1.12), Inches(6.3), Inches(2.55), th1["card_bg_1"], th1["border_col"], 3.0)
    tx_bl = s4.shapes.add_textbox(Inches(0.45), Inches(1.22), Inches(5.9), Inches(2.35))
    tf_bl = tx_bl.text_frame
    tf_bl.word_wrap = True
    p = tf_bl.paragraphs[0]
    p.text = "🧭 兩大提問：想好 ＆ 填好"
    p.font.name = FONT_NAME
    p.font.size = Pt(24)
    p.font.bold = True
    p.font.color.rgb = th1["accent_primary"]
    p2 = tf_bl.add_paragraph()
    p2.text = "• 愛實作討厭背書 ➔ 鎖定【公立高職＋私校王牌工科/餐旅】\n• 愛讀書學術研究 ➔ 鎖定【公立普通高中梯隊】\n• 絕不拿普通高中去塞技職生的志願表！"
    p2.font.name = FONT_NAME
    p2.font.size = Pt(16)
    p2.font.bold = True
    p2.font.color.rgb = th1["text_main"]
    p2.space_before = Pt(8)
    p2.line_spacing = 1.3

    b_right = add_block(s4, Inches(6.75), Inches(1.12), Inches(6.333), Inches(2.55), th1["card_bg_4"], th1["accent_danger"], 3.0)
    tx_br = s4.shapes.add_textbox(Inches(6.95), Inches(1.22), Inches(5.95), Inches(2.35))
    tf_br = tx_br.text_frame
    tf_br.word_wrap = True
    p = tf_br.paragraphs[0]
    p.text = "💥 落榜悔恨三大危機公式"
    p.font.name = FONT_NAME
    p.font.size = Pt(24)
    p.font.bold = True
    p.font.color.rgb = th1["accent_danger"]
    p2 = tf_br.add_paragraph()
    p2.text = "【操作手滑】（點錯農機生機科、連填兩次竹山長億）\n＋【順序倒填】（私校放第1、長億放第3擋死文華）\n＋【脫離現實】（跨縣市去彰化田中、南投水里深山）"
    p2.font.name = FONT_NAME
    p2.font.size = Pt(16)
    p2.font.bold = True
    p2.font.color.rgb = th1["text_main"]
    p2.space_before = Pt(8)
    p2.line_spacing = 1.3

    # 下半部三欄滿版色塊：2 : 5 : 3 黃金結構
    cols_s4 = [
        ("🎯 第 1~3 志願：夢想衝刺", "填高 5~10 點的同群公立名校",
         "前10志願同為30分滿分！2B3C愛修車先衝『霧峰農工汽修』；1A4B愛設計先衝『興大附農/台中家商』！沒上零扣分！",
         th1["card_bg_2"], th1["accent_warn"]),
        ("🛡️ 第 4~7 志願：務實主戰場", "嚴格按積點由高到低順排",
         "核心錄取區！必須100%符合「前面分數高、後面分數低」，絕不把低分學校插在前面當攔路虎！",
         th1["card_bg_3"], th1["accent_primary"]),
        ("⚓ 第 8~10 志願：鐵壁保底", "同校連填不同科＝只算1志願序",
         "放100%穩上的公立保底或市區校車私校（如青年/僑泰汽修、新民日語），確保零落榜！",
         th1["card_bg_5"], th1["accent_success"])
    ]
    for idx, (ct, cs, cd, cbg, acol) in enumerate(cols_s4):
        cx = Inches(0.25 + idx * 4.31)
        add_block(s4, cx, Inches(3.82), Inches(4.20), Inches(3.48), cbg, th1["border_col"], 2.5)
        tx = s4.shapes.add_textbox(cx + Inches(0.2), Inches(3.95), Inches(3.8), Inches(3.2))
        tf = tx.text_frame
        tf.word_wrap = True
        p1 = tf.paragraphs[0]
        p1.text = ct
        p1.font.name = FONT_NAME
        p1.font.size = Pt(21)
        p1.font.bold = True
        p1.font.color.rgb = acol
        p2 = tf.add_paragraph()
        p2.text = cs
        p2.font.name = FONT_NAME
        p2.font.size = Pt(16)
        p2.font.bold = True
        p2.font.color.rgb = th1["text_main"]
        p2.space_before = Pt(6)
        p3 = tf.add_paragraph()
        p3.text = cd
        p3.font.name = FONT_NAME
        p3.font.size = Pt(14.5)
        p3.font.bold = True
        p3.font.color.rgb = th1["text_sub"]
        p3.space_before = Pt(8)
        p3.line_spacing = 1.25

    # ==========================================================================
    # SLIDE 5: 16位同學 7 大典型盲點分流總表 (Style 3 Color-Block Magazine)
    # ==========================================================================
    th3 = get_theme("cases")
    s5 = prs.slides.add_slide(blank_layout)
    add_bg(s5, th3["slide_bg"])
    add_top_bar(s5, th3, "全班 16 份模擬卷大體檢 ｜ 七大盲點對號入座",
                "16 位趣味角色選填盲點全景雷達圖（看看誰跟你一樣？）", "🔍 7大盲點一次排雷", is_danger=True)

    cats_s5 = [
        ("1. 系統手滑誤勾 / 重複連填", "#06 手滑巴菲特 (誤勾農機生機科) ｜ #08 影分身謝雙胞 (竹山連填2次) ｜ #10 雙胞胎郭孟孟 (長億連填2次)", th3["card_bg_2"], th3["accent_warn"]),
        ("2. 局部倒填 ＆ 缺保底安全網", "#05 鋼索哥柯傲天 (86點把二中排在文華前、東山排惠文前，且後方無社區保底網！)", th3["card_bg_3"], th3["accent_danger"]),
        ("3. 毀滅倒填 / 私校門神 / 水壩", "#02 倒立哥陳步步 (長億填第1) ｜ #03 動漫小迷妹 (新民擋第3) ｜ #12 門神小卡卡 ｜ #13 林芸芸 & #16 小安 (長億擋文華)", th3["card_bg_3"], th3["accent_danger"]),
        ("4. 方向模糊 ＆ 跨群大包牌", "#07 包牌王劉阿豪 (5B漏填霧峰農工、慈明全包) ｜ #09 斜槓小當家 (餐飲美容汽車普通科5群大雜燴)", th3["card_bg_5"], th3["accent_primary"]),
        ("5. 低估實力（放棄公立高職）", "#01 修車阿樂 (2B3C第1填電商/漏填霧農工科) ｜ #08 謝雙胞 (68點第1填霧農棄附農) ｜ #14 灰姑娘小廷 (62點被11私校包圍)", th3["card_bg_1"], th3["accent_primary"]),
        ("6. 親子諜對諜（賭氣敷衍選填）", "#11 電競熱舞翰 (66點第1填新民表藝、把台中高工塞第8敷衍爸媽，中間亂塞4所幼保科)", th3["card_bg_2"], th3["accent_warn"]),
        ("7. 跨縣市魔鬼通勤 ＆ 高學費", "#04 剪刀手愛美麗 (填到彰化田中達德、苗栗卓蘭) ｜ #15 米其林仙子 (64點第2填高學費葳格截殺霧農、第8填南投水里深山)", th3["card_bg_4"], th3["accent_success"])
    ]
    for idx, (c_title, c_who, cbg, acol) in enumerate(cats_s5):
        cy = Inches(1.12 + idx * 0.88)
        add_block(s5, Inches(0.25), cy, Inches(4.0), Inches(0.80), acol, th3["border_col"], 1.5)
        tx1 = s5.shapes.add_textbox(Inches(0.35), cy + Inches(0.08), Inches(3.8), Inches(0.65))
        tx1.text_frame.vertical_anchor = MSO_ANCHOR.MIDDLE
        p1 = tx1.text_frame.paragraphs[0]
        p1.text = c_title
        p1.font.name = FONT_NAME
        p1.font.size = Pt(16)
        p1.font.bold = True
        p1.font.color.rgb = RGBColor(255, 255, 255) if th3["name"] != "風格2：沉浸式暗色科技風" else RGBColor(18, 18, 20)

        add_block(s5, Inches(4.30), cy, Inches(8.783), Inches(0.80), cbg, th3["border_col"], 1.5)
        tx2 = s5.shapes.add_textbox(Inches(4.45), cy + Inches(0.08), Inches(8.5), Inches(0.65))
        tx2.text_frame.word_wrap = True
        tx2.text_frame.vertical_anchor = MSO_ANCHOR.MIDDLE
        p2 = tx2.text_frame.paragraphs[0]
        p2.text = c_who
        p2.font.name = FONT_NAME
        p2.font.size = Pt(14.5)
        p2.font.bold = True
        p2.font.color.rgb = th3["text_main"]

    # ==========================================================================
    # SLIDES 6~37: 16 位學生雙頁滿版大字版
    #   Page A (個案深度解析):
    #     - 左側 45% (視覺焦點): OCR 乾淨大字表（字大滿版）
    #     - 右側 55% (精準吸收): 只保留「人物故事梗概」並放大滿版（粗體鮮明）
    #   Page B (專家急救處方):
    #     - 滿版大色塊卡片化分離、文字精簡粗體鮮明、落實 3 秒原則
    # ==========================================================================
    for st in STUDENTS_DATA:
        sid = st["id"]
        th_case = get_theme("cases")

        # ----------------------------------------------------------------------
        # PAGE A: 左 45% OCR 乾淨大字表滿版 ＋ 右 55% 人物故事梗概放大滿版
        # ----------------------------------------------------------------------
        sa = prs.slides.add_slide(blank_layout)
        add_bg(sa, th_case["slide_bg"])
        add_top_bar(sa, th_case,
                    f"個案深度解析 ｜ 學生 #{sid:02d} 【{st['name']}】",
                    st["title"],
                    badge_text=st["risk"], is_danger=("🚨" in st["risk"]))

        # ==================== 左側 45%：OCR 乾淨大字表（字大滿版） ====================
        # Left 45% Region: left=0.25", width=5.80", top=1.10", height=6.22"
        add_block(sa, Inches(0.25), Inches(1.10), Inches(5.80), Inches(6.22),
                  th_case["card_bg_1"], th_case["border_col"], th_case["border_width"])

        # 左側頂部標題橫條
        lbl_left = add_block(sa, Inches(0.25), Inches(1.10), Inches(5.80), Inches(0.52),
                             th_case["table_hdr_bg"], None, 0)
        ltf = lbl_left.text_frame
        ltf.vertical_anchor = MSO_ANCHOR.MIDDLE
        lp = ltf.paragraphs[0]
        lp.text = f"🔍 原始志願表 OCR 乾淨大字版（#{sid:02d} 填寫順序）"
        lp.font.name = FONT_NAME
        lp.font.size = Pt(16.5)
        lp.font.bold = True
        lp.font.color.rgb = th_case["table_hdr_fg"]
        lp.alignment = PP_ALIGN.CENTER

        # 滿版大字表格 (11 列: 1 表頭 + 10 志願)
        n_rows = len(st["choices"]) + 1
        t_box = sa.shapes.add_table(n_rows, 3, Inches(0.32), Inches(1.66), Inches(5.66), Inches(5.58))
        tbl = t_box.table
        tbl.columns[0].width = Inches(0.76)
        tbl.columns[1].width = Inches(2.50)
        tbl.columns[2].width = Inches(2.40)

        t_headers = ["序位", "選填學校與科別", "門檻 ＆ 演算法下場"]
        for c_i, h in enumerate(t_headers):
            cell = tbl.cell(0, c_i)
            cell.text = h
            cell.fill.solid()
            cell.fill.fore_color.rgb = th_case["header_bg"]
            cell.text_frame.vertical_anchor = MSO_ANCHOR.MIDDLE
            for p in cell.text_frame.paragraphs:
                p.font.name = FONT_NAME
                p.font.size = Pt(13)
                p.font.bold = True
                p.font.color.rgb = th_case["header_title"]
                p.alignment = PP_ALIGN.CENTER

        for r_i, ch in enumerate(st["choices"]):
            row_idx = r_i + 1
            c0 = tbl.cell(row_idx, 0)
            c1 = tbl.cell(row_idx, 1)
            c2 = tbl.cell(row_idx, 2)

            c0.text = f"{r_i+1:02d}"
            c1.text = f"{ch[0]} {ch[1]}"
            c2.text = ch[2]

            is_trap = any(k in ch[2] for k in ["🚨", "❌", "⚠️", "直接成交", "作廢", "攔截", "截殺", "地雷", "浪費"])
            bg_col = th_case["table_row_trap"] if is_trap else (th_case["table_row_even"] if r_i % 2 == 0 else th_case["table_row_odd"])
            fg_col = th_case["table_trap_fg"] if is_trap else th_case["text_main"]

            for c_idx, c_cell in enumerate([c0, c1, c2]):
                c_cell.fill.solid()
                c_cell.fill.fore_color.rgb = bg_col
                c_cell.text_frame.margin_left = Inches(0.05)
                c_cell.text_frame.margin_right = Inches(0.04)
                c_cell.text_frame.margin_top = Inches(0.02)
                c_cell.text_frame.margin_bottom = Inches(0.02)
                c_cell.text_frame.vertical_anchor = MSO_ANCHOR.MIDDLE
                for p in c_cell.text_frame.paragraphs:
                    p.font.name = FONT_NAME
                    p.font.size = Pt(13 if c_idx < 2 else 12)
                    p.font.bold = True
                    p.font.color.rgb = fg_col if c_idx == 2 else th_case["text_main"]
                    if c_idx == 0:
                        p.alignment = PP_ALIGN.CENTER

        # ==================== 右側 55%：人物故事梗概（放大滿版） ====================
        # Right 55% Region: left=6.18", width=6.90", top=1.10", height=6.22"
        # 採用 5 大滿版幾何色塊垂直堆疊，字體全面粗體放大 (17pt~22pt)！
        story_blocks = [
            ("⚡ 會考戰力落點", st["score"], th_case["card_bg_2"], th_case["accent_warn"], 0.96, 20),
            ("🔧 天賦興趣與專長", st["interest"], th_case["card_bg_4"], th_case["accent_success"], 1.28, 16.5),
            ("👨‍👩‍👦 家庭期待與聲音", st["family"], th_case["card_bg_1"], th_case["accent_primary"], 1.18, 16.5),
            ("💥 選填盲點與慘劇", st["doubts"], th_case["card_bg_3"], th_case["accent_danger"], 1.38, 16.5),
            ("🏷️ 專家診斷分類", st["category_tag"], th_case["card_bg_5"], th_case["accent_primary"], 0.94, 16.5)
        ]

        curr_y = 1.10
        for b_label, b_val, b_bg, b_acc, b_h, f_sz in story_blocks:
            # 左側色塊標籤頭 (寬 2.15 吋)
            lbl_box = add_block(sa, Inches(6.18), Inches(curr_y), Inches(2.22), Inches(b_h),
                                b_acc, th_case["border_col"], 1.5)
            ltf = lbl_box.text_frame
            ltf.word_wrap = True
            ltf.vertical_anchor = MSO_ANCHOR.MIDDLE
            lp = ltf.paragraphs[0]
            lp.text = b_label
            lp.font.name = FONT_NAME
            lp.font.size = Pt(16.5)
            lp.font.bold = True
            lp.font.color.rgb = RGBColor(255, 255, 255) if th_case["name"] != "風格2：沉浸式暗色科技風" else RGBColor(18, 18, 20)
            lp.alignment = PP_ALIGN.CENTER

            # 右側故事內容大字塊 (寬 4.68 吋)
            val_box = add_block(sa, Inches(8.42), Inches(curr_y), Inches(4.66), Inches(b_h),
                                b_bg, th_case["border_col"], 1.5)
            vtf = val_box.text_frame
            vtf.word_wrap = True
            vtf.vertical_anchor = MSO_ANCHOR.MIDDLE
            vtf.margin_left = Inches(0.14)
            vtf.margin_right = Inches(0.12)
            vp = vtf.paragraphs[0]
            vp.text = b_val
            vp.font.name = FONT_NAME
            vp.font.size = Pt(f_sz)
            vp.font.bold = True
            vp.font.color.rgb = th_case["text_main"]
            vp.line_spacing = 1.2

            curr_y += b_h + 0.12

        # ----------------------------------------------------------------------
        # PAGE B: 專家急救處方（文字精簡粗體鮮明、滿版色塊卡片化分離、3秒原則）
        # ----------------------------------------------------------------------
        sb = prs.slides.add_slide(blank_layout)
        add_bg(sb, th_case["slide_bg"])
        add_top_bar(sb, th_case,
                    f"專家急救處方 ｜ 學生 #{sid:02d} 【{st['name']}】",
                    f"【接地氣重整戰略】{st['title']}",
                    badge_text=st["tag_p2"], is_danger=False)

        # 上半部左卡片 (48% 寬)：🌟 天賦優勢 ＆ 老師懂你的心 (滿版色塊)
        add_block(sb, Inches(0.25), Inches(1.10), Inches(6.25), Inches(2.75),
                  th_case["card_bg_4"], th_case["accent_success"], 2.5)
        hdr_b1 = add_block(sb, Inches(0.25), Inches(1.10), Inches(6.25), Inches(0.55),
                           th_case["accent_success"], None, 0)
        htf1 = hdr_b1.text_frame
        htf1.vertical_anchor = MSO_ANCHOR.MIDDLE
        hp1 = htf1.paragraphs[0]
        hp1.text = "🌟 你的超能力 ＆ 老師坐在你身邊懂你"
        hp1.font.name = FONT_NAME
        hp1.font.size = Pt(19)
        hp1.font.bold = True
        hp1.font.color.rgb = RGBColor(255, 255, 255) if th_case["name"] != "風格2：沉浸式暗色科技風" else RGBColor(18, 18, 20)
        hp1.alignment = PP_ALIGN.CENTER

        tx_b1 = sb.shapes.add_textbox(Inches(0.42), Inches(1.72), Inches(5.92), Inches(2.05))
        tf_b1 = tx_b1.text_frame
        tf_b1.word_wrap = True
        pb1_1 = tf_b1.paragraphs[0]
        pb1_1.text = f"💪 {st['strength']}"
        pb1_1.font.name = FONT_NAME
        pb1_1.font.size = Pt(16)
        pb1_1.font.bold = True
        pb1_1.font.color.rgb = th_case["text_main"]
        pb1_1.line_spacing = 1.2
        pb1_1.space_after = Pt(8)

        pb1_2 = tf_b1.add_paragraph()
        pb1_2.text = f"🤝 {st['anxiety_empathy']}"
        pb1_2.font.name = FONT_NAME
        pb1_2.font.size = Pt(15.5)
        pb1_2.font.bold = True
        pb1_2.font.color.rgb = th_case["text_sub"]
        pb1_2.line_spacing = 1.2

        # 上半部右卡片 (50% 寬)：🚨 演算法致命地雷（為何會出事？）
        add_block(sb, Inches(6.65), Inches(1.10), Inches(6.433), Inches(2.75),
                  th_case["card_bg_3"], th_case["accent_danger"], 2.5)
        hdr_b2 = add_block(sb, Inches(6.65), Inches(1.10), Inches(6.433), Inches(0.55),
                           th_case["accent_danger"], None, 0)
        htf2 = hdr_b2.text_frame
        htf2.vertical_anchor = MSO_ANCHOR.MIDDLE
        hp2 = htf2.paragraphs[0]
        hp2.text = "🚨 演算法殘酷真相（電腦不講人情！）"
        hp2.font.name = FONT_NAME
        hp2.font.size = Pt(19)
        hp2.font.bold = True
        hp2.font.color.rgb = RGBColor(255, 255, 255)
        hp2.alignment = PP_ALIGN.CENTER

        tx_b2 = sb.shapes.add_textbox(Inches(6.82), Inches(1.72), Inches(6.12), Inches(2.05))
        tf_b2 = tx_b2.text_frame
        tf_b2.word_wrap = True
        pb2_1 = tf_b2.paragraphs[0]
        pb2_1.text = f"❌ {st['traps'][0]}"
        pb2_1.font.name = FONT_NAME
        pb2_1.font.size = Pt(15.5)
        pb2_1.font.bold = True
        pb2_1.font.color.rgb = th_case["table_trap_fg"] if th_case["name"] == "風格2：沉浸式暗色科技風" else th_case["accent_danger"]
        pb2_1.line_spacing = 1.2
        pb2_1.space_after = Pt(8)

        pb2_2 = tf_b2.add_paragraph()
        pb2_2.text = f"❌ {st['traps'][1]}"
        pb2_2.font.name = FONT_NAME
        pb2_2.font.size = Pt(15)
        pb2_2.font.bold = True
        pb2_2.font.color.rgb = th_case["text_main"]
        pb2_2.line_spacing = 1.2

        # 下半部滿版雙欄行動處方卡 (高度 2.55 吋)
        # 左行動卡：🎯 前段衝刺與主戰場
        add_block(sb, Inches(0.25), Inches(3.98), Inches(6.25), Inches(2.45),
                  th_case["card_bg_2"], th_case["accent_warn"], 2.5)
        hdr_s1 = add_block(sb, Inches(0.25), Inches(3.98), Inches(6.25), Inches(0.50),
                           th_case["accent_warn"], None, 0)
        hs1_tf = hdr_s1.text_frame
        hs1_tf.vertical_anchor = MSO_ANCHOR.MIDDLE
        hs1_p = hs1_tf.paragraphs[0]
        hs1_p.text = "🎯 處方 1：前段衝刺 ＆ 核心主戰場重排"
        hs1_p.font.name = FONT_NAME
        hs1_p.font.size = Pt(18)
        hs1_p.font.bold = True
        hs1_p.font.color.rgb = RGBColor(255, 255, 255) if th_case["name"] != "風格2：沉浸式暗色科技風" else RGBColor(18, 18, 20)
        hs1_p.alignment = PP_ALIGN.CENTER

        tx_s1 = sb.shapes.add_textbox(Inches(0.42), Inches(4.54), Inches(5.92), Inches(1.82))
        tf_s1 = tx_s1.text_frame
        tf_s1.word_wrap = True
        tf_s1.vertical_anchor = MSO_ANCHOR.MIDDLE
        ps1 = tf_s1.paragraphs[0]
        ps1.text = f"✅ {st['solutions'][0]}"
        ps1.font.name = FONT_NAME
        ps1.font.size = Pt(16)
        ps1.font.bold = True
        ps1.font.color.rgb = th_case["text_main"]
        ps1.line_spacing = 1.25

        # 右行動卡：🛡️ 後段保底與排雷行動
        add_block(sb, Inches(6.65), Inches(3.98), Inches(6.433), Inches(2.45),
                  th_case["card_bg_1"], th_case["accent_primary"], 2.5)
        hdr_s2 = add_block(sb, Inches(6.65), Inches(3.98), Inches(6.433), Inches(0.50),
                           th_case["accent_primary"], None, 0)
        hs2_tf = hdr_s2.text_frame
        hs2_tf.vertical_anchor = MSO_ANCHOR.MIDDLE
        hs2_p = hs2_tf.paragraphs[0]
        hs2_p.text = "🛡️ 處方 2：穩固保底防線 ＆ 徹底排雷"
        hs2_p.font.name = FONT_NAME
        hs2_p.font.size = Pt(18)
        hs2_p.font.bold = True
        hs2_p.font.color.rgb = RGBColor(255, 255, 255) if th_case["name"] != "風格2：沉浸式暗色科技風" else RGBColor(18, 18, 20)
        hs2_p.alignment = PP_ALIGN.CENTER

        tx_s2 = sb.shapes.add_textbox(Inches(6.82), Inches(4.54), Inches(6.12), Inches(1.82))
        tf_s2 = tx_s2.text_frame
        tf_s2.word_wrap = True
        tf_s2.vertical_anchor = MSO_ANCHOR.MIDDLE
        ps2 = tf_s2.paragraphs[0]
        ps2.text = f"✅ {st['solutions'][1]}"
        ps2.font.name = FONT_NAME
        ps2.font.size = Pt(16)
        ps2.font.bold = True
        ps2.font.color.rgb = th_case["text_main"]
        ps2.line_spacing = 1.25

        # 底部滿版溫暖賦能金句橫幅 (高度 0.78 吋)
        warm_bar = add_block(sb, Inches(0.25), Inches(6.54), Inches(12.833), Inches(0.80),
                             th_case["banner_bg"], th_case["border_col"], 2.0)
        wtf = warm_bar.text_frame
        wtf.word_wrap = True
        wtf.vertical_anchor = MSO_ANCHOR.MIDDLE
        wp = wtf.paragraphs[0]
        wp.text = f"❤️ 專家給你的真心話：「{st['warmth']}」"
        wp.font.name = FONT_NAME
        wp.font.size = Pt(17.5)
        wp.font.bold = True
        wp.font.color.rgb = th_case["banner_fg"]
        wp.alignment = PP_ALIGN.CENTER

    # ==========================================================================
    # SLIDE 38: 親師陪伴對話心法 (Concept -> Style 1 Billboard)
    # ==========================================================================
    s38 = prs.slides.add_slide(blank_layout)
    add_bg(s38, th1["slide_bg"])
    add_top_bar(s38, th1, "親師陪伴與對話指引 ｜ 智慧化解選填焦慮",
                "給正在迷茫同學的三句真心話：走對路，比搶名校更發光！", "❤️ 坐在你身邊陪你選", is_danger=False)

    cards_38 = [
        ("1. 會考只是分組，不是判決書", "紙筆測驗不代表你的全部價值",
         "看到考卷就頭痛也沒關係！像修車阿樂一樣，雙手靈巧、懂機械、愛料理或愛設計，走進技術型高中就是你的王牌主場！",
         th1["card_bg_1"], th1["accent_primary"]),
        ("2. 勇敢爭取公立技職，不自我放棄", "5B 或 2B3C 也有公立高職能拼！",
         "別以為沒考 A 就只能隨便抓私校塞第 1 志願！前 10 志願同拿 30 分，把霧峰農工、草屯商工、豐原高商排前面衝刺，完全零扣分！",
         th1["card_bg_2"], th1["accent_warn"]),
        ("3. 雙贏溝通，不拿志願表賭氣", "找尋科技與興趣的交集點",
         "跟爸媽意見不同時，別像電競熱舞翰把高工塞第 8 志願敷衍！善用『高工圖文傳播科、家商服飾/多媒體』，達成親子雙贏！",
         th1["card_bg_5"], th1["accent_success"])
    ]
    for idx, (ct, cs, cd, cbg, acol) in enumerate(cards_38):
        cx = Inches(0.25 + idx * 4.31)
        add_block(s38, cx, Inches(1.15), Inches(4.20), Inches(6.15), cbg, th1["border_col"], 3.0)
        hdr = add_block(s38, cx, Inches(1.15), Inches(4.20), Inches(0.85), acol, None, 0)
        htf = hdr.text_frame
        htf.vertical_anchor = MSO_ANCHOR.MIDDLE
        hp = htf.paragraphs[0]
        hp.text = ct
        hp.font.name = FONT_NAME
        hp.font.size = Pt(20)
        hp.font.bold = True
        hp.font.color.rgb = RGBColor(255, 255, 255) if th1["name"] != "風格2：沉浸式暗色科技風" else RGBColor(18, 18, 20)
        hp.alignment = PP_ALIGN.CENTER

        tx = s38.shapes.add_textbox(cx + Inches(0.22), Inches(2.20), Inches(3.76), Inches(4.9))
        tf = tx.text_frame
        tf.word_wrap = True
        p1 = tf.paragraphs[0]
        p1.text = cs
        p1.font.name = FONT_NAME
        p1.font.size = Pt(19)
        p1.font.bold = True
        p1.font.color.rgb = acol
        p1.space_after = Pt(14)

        p2 = tf.add_paragraph()
        p2.text = cd
        p2.font.name = FONT_NAME
        p2.font.size = Pt(16.5)
        p2.font.bold = True
        p2.font.color.rgb = th1["text_main"]
        p2.line_spacing = 1.35

    # ==========================================================================
    # SLIDE 39: 正式送出前「五步重整急救包」 (System -> Style 2 Dark Tech)
    # ==========================================================================
    s39 = prs.slides.add_slide(blank_layout)
    add_bg(s39, th2["slide_bg"])
    add_top_bar(s39, th2, "實戰行動 SOP ｜ 正式志願表列印簽名前必做",
                "正式送出前 30 秒保命檢核：「五步重整急救包」", "✅ 5步動作零失誤", is_danger=False)

    steps_39 = [
        ("STEP 1", "揪出「倒填與水壩」", "由上往下檢查積點：絕不讓長億(53點)或二中(70點)擋在文華(85點)、惠文(76點)或東山(72點)前面！", th2["card_bg_3"], th2["accent_danger"]),
        ("STEP 2", "移開「私校攔路門神」", "第 1~5 志願全力留給想讀的公立高中職！私校保底移至第 6~10 志願，絕不提早截殺公立機會！", th2["card_bg_5"], th2["accent_warn"]),
        ("STEP 3", "刪除「湊數與跨縣市地雷」", "技職生刪掉湊數的普通高中；刪掉彰化田中達德、苗栗三義卓蘭、南投水里深山等超遠距學校！", th2["card_bg_2"], th2["accent_primary"]),
        ("STEP 4", "啟動「同校連填不扣分」", "同校連續選填不同科（如家商連填4科、霧農連填4科、慈明連填3科）只算1個志願序，把安全網織到最密！", th2["card_bg_4"], th2["accent_success"]),
        ("STEP 5", "核對「重複格子與科名」", "檢查有無連按兩次同校同科（如竹山廣設、長億普通科），確認無誤再列印正式志願表完成家長雙簽！", th2["card_bg_1"], th2["accent_primary"])
    ]
    for idx, (s_num, s_tit, s_desc, cbg, acol) in enumerate(steps_39):
        sy = Inches(1.12 + idx * 1.24)
        add_block(s39, Inches(0.25), sy, Inches(1.85), Inches(1.12), acol, th2["border_col"], 2.0)
        ntf = s39.shapes.add_textbox(Inches(0.25), sy + Inches(0.2), Inches(1.85), Inches(0.7)).text_frame
        np = ntf.paragraphs[0]
        np.text = s_num
        np.font.name = FONT_NAME
        np.font.size = Pt(24)
        np.font.bold = True
        np.font.color.rgb = RGBColor(18, 18, 20) if th2["name"] == "風格2：沉浸式暗色科技風" else RGBColor(255, 255, 255)
        np.alignment = PP_ALIGN.CENTER

        add_block(s39, Inches(2.15), sy, Inches(10.933), Inches(1.12), cbg, acol, 2.0)
        tx = s39.shapes.add_textbox(Inches(2.35), sy + Inches(0.08), Inches(10.55), Inches(0.96))
        tf = tx.text_frame
        tf.word_wrap = True
        tf.vertical_anchor = MSO_ANCHOR.MIDDLE
        p1 = tf.paragraphs[0]
        p1.text = f"【{s_tit}】 {s_desc}"
        p1.font.name = FONT_NAME
        p1.font.size = Pt(16.5)
        p1.font.bold = True
        p1.font.color.rgb = th2["text_main"]

    # ==========================================================================
    # SLIDE 40: 結語與祝福滿版看板 (Concept -> Style 1 Billboard)
    # ==========================================================================
    s40 = prs.slides.add_slide(blank_layout)
    add_bg(s40, th1["cover_bg"])
    add_block(s40, Inches(0.45), Inches(0.45), Inches(12.433), Inches(6.6),
              th1["cover_bg"], th1["cover_text"], 4.0)

    tx40 = s40.shapes.add_textbox(Inches(0.9), Inches(1.1), Inches(11.5), Inches(5.3))
    tf40 = tx40.text_frame
    tf40.word_wrap = True
    p40 = tf40.paragraphs[0]
    p40.text = "適性揚才 ｜ 走對跑道，你就是第一名！"
    p40.font.name = FONT_NAME
    p40.font.size = Pt(44)
    p40.font.bold = True
    p40.font.color.rgb = th1["cover_text"]

    p40_2 = tf40.add_paragraph()
    p40_2.text = (
        "• 喜歡動手修車、做料理、畫設計 ➔ 技職舞台讓你成為金牌職人！\n"
        "• 喜歡鑽研數理、語文、學術探究 ➔ 普通高中讓你展翅高飛！\n"
        "• 記住：前 10 志願同分 30 分，把最愛且務實的公立目標放最前，\n"
        "  把穩健的保底守門員放第 6~10 志願，迎向屬於你的璀璨三年！"
    )
    p40_2.font.name = FONT_NAME
    p40_2.font.size = Pt(24)
    p40_2.font.bold = True
    p40_2.font.color.rgb = th1["cover_sub"]
    p40_2.space_before = Pt(24)
    p40_2.line_spacing = 1.35

    prs.save(out_pptx)
    print(f"[OK] Saved PPTX ({mode_key}): {out_pptx}")
    return out_pptx


def convert_all_pptx_to_pdf(pptx_list):
    print("Converting all PPTX files to PDF via PowerPoint COM (comtypes)...")
    ppt_app = comtypes.client.CreateObject("PowerPoint.Application")
    ppt_app.Visible = 1
    ppt_app.DisplayAlerts = 1  # ppAlertsNone
    pdf_list = []
    try:
        for pptx_file in pptx_list:
            out_pdf = pptx_file.replace(".pptx", ".pdf")
            abs_pptx = os.path.abspath(pptx_file)
            abs_pdf = os.path.abspath(out_pdf)
            if os.path.exists(abs_pdf):
                try:
                    os.remove(abs_pdf)
                except Exception as e:
                    print(f"Warning removing old {out_pdf}: {e}")
            deck = ppt_app.Presentations.Open(abs_pptx)
            deck.SaveAs(abs_pdf, 32)  # 32 = ppSaveAsPDF
            deck.Close()
            print(f"[OK] Generated PDF: {out_pdf}")
            pdf_list.append(out_pdf)
    finally:
        ppt_app.Quit()
    return pdf_list


if __name__ == "__main__":
    decks_to_build = [
        ("style1_billboard", "中投區國三志願選填_風格1_高對比極簡看板風.pptx"),
        ("style2_darktech", "中投區國三志願選填_風格2_沉浸式暗色科技風.pptx"),
        ("style3_colorblock", "中投區國三志願選填_風格3_大色塊幾何雜誌風.pptx"),
        ("hybrid_master", "中投區國三升學志願選填分析_專家深度診斷與策略指南_116版.pptx")
    ]
    built_pptx = []
    for m_key, fname in decks_to_build:
        built_pptx.append(build_deck_for_mode(m_key, fname))

    built_pdfs = convert_all_pptx_to_pdf(built_pptx)
    print("\n================ ALL 4 DECKS & PDFS GENERATED SUCCESSFULLY ================")
    for p, d in zip(built_pptx, built_pdfs):
        print(f" - PPTX: {p}  |  PDF: {d}")
