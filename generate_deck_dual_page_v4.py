# -*- coding: utf-8 -*-
"""
generate_deck_dual_page_v4.py
Generates the 40-slide presentation matching the user's latest requirements:
- Large fonts, reduced word count, high contrast for 70-inch classroom TV.
- Each student has 2 pages:
  - Page 1: Left half has enlarged original student form image; Right half has OCR converted clean bold table text.
  - Page 2: Expert admission counseling analysis & recommendations, concise, positive, empathetic, and encouraging.
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
import pymupdf

# Colors - Bright, comfortable, modern
C_NAVY_DARK    = RGBColor(15, 23, 42)     # Slate 900
C_NAVY_LIGHT   = RGBColor(30, 41, 59)     # Slate 800
C_BG_LIGHT     = RGBColor(248, 250, 252)  # Slate 50
C_WHITE        = RGBColor(255, 255, 255)
C_BORDER       = RGBColor(203, 213, 225)  # Slate 300
C_TEXT_MAIN    = RGBColor(15, 23, 42)     # Deep dark for readability
C_TEXT_MUTED   = RGBColor(71, 85, 105)    # Slate 600

C_BLUE_ACCENT  = RGBColor(2, 132, 199)    # Sky 600
C_BLUE_LIGHT   = RGBColor(240, 249, 255)  # Sky 50
C_BLUE_BORDER  = RGBColor(186, 230, 253)  # Sky 200

C_ROSE_ACCENT  = RGBColor(225, 29, 72)    # Rose 600
C_ROSE_LIGHT   = RGBColor(255, 241, 242)  # Rose 50
C_ROSE_BORDER  = RGBColor(251, 113, 133)  # Rose 400

C_EMERALD_ACC  = RGBColor(5, 150, 105)    # Emerald 600
C_EMERALD_LIGHT= RGBColor(236, 253, 245)  # Emerald 50
C_EMERALD_BOR  = RGBColor(52, 211, 153)   # Emerald 400

C_AMBER_ACCENT = RGBColor(217, 119, 6)    # Amber 600
C_AMBER_LIGHT  = RGBColor(254, 252, 232)  # Amber 50
C_AMBER_BORDER = RGBColor(251, 191, 36)   # Amber 400

C_ROW_ALT      = RGBColor(241, 245, 249)  # Slate 100

FONT_NAME = "Microsoft JhengHei"

def add_bg(slide, color=C_BG_LIGHT):
    bg = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, Inches(13.333), Inches(7.5))
    bg.fill.solid()
    bg.fill.fore_color.rgb = color
    bg.line.fill.background()
    return bg

def create_card(slide, left, top, width, height, bg_color=C_WHITE, border_color=C_BORDER, line_width=1.5):
    card = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left, top, width, height)
    card.fill.solid()
    card.fill.fore_color.rgb = bg_color
    if border_color:
        card.line.color.rgb = border_color
        card.line.width = Pt(line_width)
    else:
        card.line.fill.background()
    return card

def add_header(slide, tag_text, title_text, badge_text=None, badge_color=C_ROSE_ACCENT, badge_bg=C_ROSE_LIGHT):
    tx_tag = slide.shapes.add_textbox(Inches(0.6), Inches(0.20), Inches(8.5), Inches(0.32))
    tf_tag = tx_tag.text_frame
    tf_tag.word_wrap = True
    tf_tag.margin_left = tf_tag.margin_top = tf_tag.margin_right = tf_tag.margin_bottom = 0
    p_tag = tf_tag.paragraphs[0]
    p_tag.text = tag_text
    p_tag.font.name = FONT_NAME
    p_tag.font.size = Pt(13)
    p_tag.font.bold = True
    p_tag.font.color.rgb = C_BLUE_ACCENT

    tx_title = slide.shapes.add_textbox(Inches(0.6), Inches(0.50), Inches(8.8), Inches(0.58))
    tf_title = tx_title.text_frame
    tf_title.word_wrap = True
    tf_title.margin_left = tf_title.margin_top = tf_title.margin_right = tf_title.margin_bottom = 0
    p_title = tf_title.paragraphs[0]
    p_title.text = title_text
    p_title.font.name = FONT_NAME
    p_title.font.size = Pt(19.5)
    p_title.font.bold = True
    p_title.font.color.rgb = C_NAVY_DARK

    if badge_text:
        badge = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(9.6), Inches(0.32), Inches(3.15), Inches(0.60))
        badge.fill.solid()
        badge.fill.fore_color.rgb = badge_bg
        badge.line.color.rgb = badge_color
        badge.line.width = Pt(1.5)
        btf = badge.text_frame
        btf.vertical_anchor = MSO_ANCHOR.MIDDLE
        bp = btf.paragraphs[0]
        bp.alignment = PP_ALIGN.CENTER
        bp.text = badge_text
        bp.font.name = FONT_NAME
        bp.font.size = Pt(13)
        bp.font.bold = True
        bp.font.color.rgb = badge_color

# Rich Student Data (All 16 students, positive & empathetic counseling)
STUDENTS_DATA = [
    {
        "id": 1,
        "title": "大拍賣式選填：修車電商樣樣來，最後放一中壓驚？",
        "risk": "🚨 致命毀滅性倒填",
        "badge_color": C_ROSE_ACCENT,
        "badge_bg": C_ROSE_LIGHT,
        "tag_p2": "💡 專家處方：大反轉重排 ＋ 職群聚焦",
        "strength": "對實用技能充滿熱情，跨領域嘗試的直覺很敏銳！別因模擬考分數不高而自卑，你擁有很強的動手實作潛能！",
        "anxiety_empathy": "害怕沒學校讀，急著抓私立學校求心安；填到最後想起父母期望，順手填一中壓驚。",
        "traps": [
            "電腦由上往下審查，第 1 志願慈明(約20點)直接定案，後面的公立高中連被通電讀取的機會都沒有！",
            "考 5A 滿分也只能去讀慈明，後方的忠明、惠文、台中一中 100% 全部變成廢紙！"
        ],
        "solutions": [
            "【公立名校勇敢排第 1~3 志願】前 10 志願同分滿分，大膽衝刺一中或忠明零扣分！",
            "【私立優質職校移至第 8~10 志願】作為強效保底守門員，100% 穩健安全不落榜！"
        ],
        "warmth": "會考只是一時的測驗！選一個你真心喜歡的專業練到精湛，你在技職一樣是全台灣最閃亮的大師！",
        "choices": [
            ("私立慈明高中", "電子商務科", "私立高職 門檻約20點"),
            ("私立青年高中", "電子科", "私立高職 動手實作"),
            ("私立青年高中", "汽車科", "私立高職 動力機械"),
            ("私立慈明高中", "資訊科", "私立高職 軟硬體實務"),
            ("私立青年高中", "廣告設計科", "私立高職 視覺設計"),
            ("市立忠明高中", "普通科", "公立名校 71-76點 倒填作廢"),
            ("市立西苑高中", "普通科", "公立名校 70-74點 倒填作廢"),
            ("市立東山高中", "普通科", "公立名校 70-74點 倒填作廢"),
            ("市立惠文高中", "普通科", "前段名校 76-77點 倒填作廢"),
            ("台中一中", "普通科", "頂尖首選 95-97點 殿後作廢")
        ]
    },
    {
        "id": 2,
        "title": "倒吃甘蔗變倒撞南牆！長億填第一，一中文華放水流",
        "risk": "🚨 致命毀滅性倒填",
        "badge_color": C_ROSE_ACCENT,
        "badge_bg": C_ROSE_LIGHT,
        "tag_p2": "💡 專家處方：滑道順序倒帶 ＋ 長億退後保底",
        "strength": "全班最用功負責的孩子！查滿 30 所公立高中代碼，這份認真與韌性就是你未來成功最大的本錢！",
        "anxiety_empathy": "想把每一步都走穩，誤以為像爬樓梯一樣先填長億當基底最保險，一階一階往上爬到一中。",
        "traps": [
            "電腦分發是溜滑梯不是爬樓梯！長億門檻53-64點，第 1 志願達標立刻敲定成交！",
            "考 5A++ (111點) 滿點也只能去讀長億，後面 29 所名校集體在系統外淚崩！"
        ],
        "solutions": [
            "【完全倒帶重排】一中(01)➔附中(02)➔中科(03)➔文華(04)➔惠文(05)➔二中(06)！",
            "【長億退到第 11~15 志願】保留填滿 30 個志願的好習慣，建立全班最強無敵安全網！"
        ],
        "warmth": "你的認真全班第一！把最棒的夢想放在滑道起點，溜下去才能抓住三年苦讀最漂亮的成果！",
        "choices": [
            ("市立長億高中", "普通科", "屯區指標 53-64點 致命首位"),
            ("市立大里高中", "普通科", "優質公立 61-74點 倒填遭截"),
            ("國立興大附中", "普通科", "頂尖名校 86-94點 遭長億截"),
            ("市立惠文高中", "普通科", "前段名校 76-77點 遭長億截"),
            ("市立文華高中", "普通科", "前段名校 82-88點 遭長億截"),
            ("台中一中", "普通科", "頂尖首選 95-97點 遭長億截"),
            ("國立中科實中", "普通科", "頂尖名校 86-91點 遭長億截"),
            ("台中二中", "普通科", "前段名校 69-74點 遭長億截"),
            ("市立忠明高中", "普通科", "優質公立 71-76點 遭長億截"),
            ("市立西苑高中", "普通科", "社區名校 70-74點 遭長億截")
        ]
    },
    {
        "id": 3,
        "title": "新民日文成鐵壁門神！把東山高中與台中家商擋在門外",
        "risk": "🚨 高危私校門神阻斷",
        "badge_color": C_ROSE_ACCENT,
        "badge_bg": C_ROSE_LIGHT,
        "tag_p2": "💡 專家處方：公立名校前移 ＋ 新民日文後盾",
        "strength": "對日文與多媒體設計有敏銳的審美直覺，兼具文創與語言天賦，未來走在潮流前端！",
        "anxiety_empathy": "喜歡日文和設計，又擔心沒學校念，聽講解範例時順手把長億填在第 1 個求心安。",
        "traps": [
            "新民日文(約35點)放第 3 志願變鐵壁門神，長億沒上就直接被新民接走！",
            "東山高中(70-74點)與台中家商排在第 8~10 志願，錄取門檻遠高於新民卻全被封死！"
        ],
        "solutions": [
            "【公立名校大膽前移】東山高中(01)➔市大里(02)➔興大附農(03)➔台中家商(04)！",
            "【新民日文退至 08~10 志願】穩健保底，既能衝刺頂尖公立，又能牢牢守住日文夢！"
        ],
        "warmth": "喜歡日文和文創超有眼光！別讓私校當了程咬金，把公立排前面、新民當後盾，兩邊都顧得到！",
        "choices": [
            ("市立長億高中", "普通科", "屯區指標 53-64點 抄自範例"),
            ("市立大里高中", "普通科", "優質公立 61-74點 倒填遭截"),
            ("私立新民高中", "應用日語科", "私立名科 35-48點 門神路障"),
            ("市立新社高中", "普通科", "公立保底 31-36點 安全底線"),
            ("私立新民高中", "普通科", "私立綜合高中部"),
            ("私立新民高中", "多媒體設計科", "私立文創熱門科"),
            ("私立新民高中", "觀光事業科", "私立餐旅服務類科"),
            ("市立東山高中", "普通科", "社區名校 70-74點 倒填作廢"),
            ("台中家商", "流行服飾科", "公立頂尖 55-65點 遭截作廢"),
            ("國立興大附農", "觀光事業科", "國立名校 60-68點 遭截作廢")
        ]
    },
    {
        "id": 4,
        "title": "環台美業大冒險！為了一把剪刀，剪到田中與苗栗卓蘭？",
        "risk": "⚠️ 荒謬地理通勤迷航",
        "badge_color": C_AMBER_ACCENT,
        "badge_bg": C_AMBER_LIGHT,
        "tag_p2": "💡 專家處方：鎖定市區校車名校 ＋ 家商服飾衝刺",
        "strength": "專業純度 100%！全心全意熱愛美容造型，目標純粹專一，具備成為頂級造型總監的專注特質！",
        "anxiety_empathy": "看選填代碼只認「美容」兩個字，完全沒查地圖，不知道跨縣市每天通車超過 2 小時會有多累。",
        "traps": [
            "達德在彰化田中、卓蘭在苗栗山城！每天早上 5:30 出門通車，還沒練刀工就先累癱！",
            "漏填中投區家政群公立第一霸主「台中家商流行服飾科」(55-65點)！"
        ],
        "solutions": [
            "【增列公立夢幻首選】第 1 志願大膽填「台中家商流行服飾科」免費衝刺！",
            "【斬斷跨縣市奔波】刪除彰化達德、苗栗卓蘭！留市區校車直達名校(明德/僑泰/青年)！"
        ],
        "warmth": "你的專注讓老師非常佩服！學設計最需要體力，留在市區好學校吹冷氣做造型，天天睡飽飽！",
        "choices": [
            ("私立僑泰高中", "美容科", "市區名校 校車四通八達"),
            ("私立慈明高中", "美容科", "台中太平 專業時尚教室"),
            ("私立青年高中", "美容科", "台中大里 美容美髮老牌"),
            ("私立玉山高中", "美容科", "東勢深山 地理交通不便"),
            ("私立達德商工", "美容科", "彰化田中 跨縣市遠距"),
            ("國立卓蘭高中", "美髮技術科", "苗栗卓蘭 偏遠山城"),
            ("私立明台高中", "美容科", "台中霧峰 古蹟校園"),
            ("私立致用高中", "美容科", "台中大甲 海線遠距"),
            ("私立明德高中", "美容科", "台中南區 家政老字號"),
            ("私立明德高中", "美容科", "同校同科 重複選填無效")
        ]
    },
    {
        "id": 5,
        "title": "高空鋼索無防護網！非5A不讀，掉一題摔入重考班",
        "risk": "🚨 極高危零保底鋼索",
        "badge_color": C_ROSE_ACCENT,
        "badge_bg": C_ROSE_LIGHT,
        "tag_p2": "💡 專家處方：理順前段排序 ＋ 強制佈署安全網",
        "strength": "實力頂尖的學霸！平時模擬考 4A~5A，具有強大企圖心與自律精神，勇於挑戰巔峰！",
        "anxiety_empathy": "自尊心強，覺得填長億或清水太委屈，堅信會考一定能正常發揮，但缺乏應對失常的備案。",
        "traps": [
            "萬一會考當天發燒感冒考出 1A4B (約60點)，全填70點以上名校將直接「落榜」無校可讀！",
            "文華排二中後面、惠文排東山後面，有微幅倒填卡死風險！"
        ],
        "solutions": [
            "【理順排位】一中(01)➔附中(02)➔中科(03)➔文華(04)➔惠文(05)➔二中(06)！",
            "【強制佈署安全網】第 10 志願後排入市大里、長億、清水，完成 100% 不落榜防護！"
        ],
        "warmth": "有雄心壯志想上一中太棒了！但大將軍出征一定帶降落傘，織好防護網才能毫無顧忌全力衝刺！",
        "choices": [
            ("台中一中", "普通科", "頂尖首選 95-97點 滿分衝刺"),
            ("國立興大附中", "普通科", "頂尖名校 86-94點 男女兼收"),
            ("國立中科實中", "普通科", "頂尖名校 86-91點 數理見長"),
            ("台中二中", "普通科", "指標高中 69-74點 老牌公立"),
            ("市立文華高中", "普通科", "前段名校 82-88點 倒填遭二中截"),
            ("私立衛道中學", "普通科", "私立頂尖升學班"),
            ("私立明道中學", "普通科", "私立綜合名校高中部"),
            ("市立東山高中", "普通科", "社區名校 70-74點 捷運便利"),
            ("市立惠文高中", "普通科", "前段名校 76-77點 倒填遭東山截"),
            ("市立西苑高中", "普通科", "社區名校 70-74點 西屯首選")
        ]
    },
    {
        "id": 6,
        "title": "商管界未來CEO！階梯很漂亮，但夾了一顆生機工科地雷",
        "risk": "⚠️ 隱藏職群地雷炸裂",
        "badge_color": C_AMBER_ACCENT,
        "badge_bg": C_AMBER_LIGHT,
        "tag_p2": "💡 專家處方：果斷排除工科地雷 ＋ 商管王牌串聯",
        "strength": "商管直覺敏銳、思路條理分明！填寫 19 個志願層次清晰，善用同校多科同序原則，全班模範！",
        "anxiety_empathy": "做足功課，看到興大附農是國立名校，以為生物產業機電是好玩的農業生物，順手勾選。",
        "traps": [
            "生機科是硬核機械工科！每天要碰車床、配電路、拆引擎換油污零件，與商管背道而馳！",
            "萬一家商差 1 點沒上掉進生機科，三年會讀得極度痛苦！"
        ],
        "solutions": [
            "【立刻排雷】果斷刪除生機科，換成附農「觀光事業」或家商「應用英語」！",
            "【商管王牌一氣呵成】家商4科 ➔ 豐商4科 ➔ 長億 ➔ 草商，完美商管大範本！"
        ],
        "warmth": "你的志願表排得超有智慧！避開工科地雷，在商管專業發光發熱，三年後頂尖科大在等你！",
        "choices": [
            ("台中二中", "普通科", "前段名校 69-74點 衝刺首選"),
            ("市立東山高中", "普通科", "優質公立 70-74點 實力落點"),
            ("台中家商", "商業經營科", "全台商管第一 68-76點"),
            ("台中家商", "國際貿易科", "全台商管榜首 68-76點"),
            ("台中家商", "資料處理科", "資訊商管熱門 55-65點"),
            ("台中家商", "會計事務科", "財務核心基石 55-65點"),
            ("國立興大附農", "生物產業機電科", "硬核工科！車床實習地雷"),
            ("國立豐原高商", "商經/國貿/資處/服飾", "山線老牌商職 4科連填"),
            ("市立長億高中", "普通科", "屯區指標 53-64點 安全保底"),
            ("國立草屯商工", "商業經營科", "公立商職 40-52點 依序至19")
        ]
    },
    {
        "id": 7,
        "title": "慈明百貨全包VIP！買下航電電商餐飲，明道電子瑟瑟發抖",
        "risk": "⚠️ 私校大包牌與倒填",
        "badge_color": C_AMBER_ACCENT,
        "badge_bg": C_AMBER_LIGHT,
        "tag_p2": "💡 專家處方：科技修護向前衝 ＋ 慈明退後守門",
        "strength": "對航空電子與新興科技充满好奇與探索慾，視野前衛，敢於接觸高科技新領域！",
        "anxiety_empathy": "覺得慈明校舍新穎設備好，全勾進去；第 10 志願隨手填了東山高中。",
        "traps": [
            "慈明門檻約 20 點，第 1 志願直接成交定案，後面的明道與弘文根本沒機會審核！",
            "航電、電商、資訊、餐飲目標太分散，未來統測跨考會大幅分散精力！"
        ],
        "solutions": [
            "【科技修護往前排】明道電子(01)➔僑泰/高工➔慈明航電/資訊(05保底)！",
            "【若有 2A 實力】東山高中放第 1 志願衝刺，弘文/明道放第 2~3 志願！"
        ],
        "warmth": "高科技航空與電子很有前途！選最想讀的一兩所放前面，慈明在後守門，這樣才吃得飽又吃得好！",
        "choices": [
            ("私立慈明高中", "航空電子科", "新興科技 門檻約20點"),
            ("私立慈明高中", "電子商務科", "電商網拍實務"),
            ("私立慈明高中", "資訊科", "軟硬體實習"),
            ("私立慈明高中", "餐飲管理科", "跳躍跨群 餐飲廚藝"),
            ("私立明道中學", "電子科", "升學私校 35-45點 遭慈明截"),
            ("私立弘文中學", "普通科", "潭子私立名校高中部"),
            ("私立立人高中", "普通科", "大里私立名校高中部"),
            ("私立青年高中", "電子科", "私立高職工科實作"),
            ("私立嘉陽高中", "飛機修護科", "清水海線修護科系"),
            ("市立東山高中", "普通科", "公立名校 70-74點 殿後作廢")
        ]
    },
    {
        "id": 8,
        "title": "複製貼上迷航記！竹山廣設連填兩次，家商國貿倒填作廢",
        "risk": "⚠️ 複製錯誤與嚴重倒填",
        "badge_color": C_AMBER_ACCENT,
        "badge_bg": C_AMBER_LIGHT,
        "tag_p2": "💡 專家處方：剔除重複與跨縣市 ＋ 家商國貿前移",
        "strength": "兼具商管與設計美感天賦，跨領域整合潛力巨大，心思細膩！",
        "anxiety_empathy": "想兼顧設計與餐飲，心情慌亂手滑連按了兩次代碼，又以為苗栗三義離家很近。",
        "traps": [
            "重複填寫浪費寶貴志願額度；苗栗三義每天通車單趟逾 80 分鐘太辛苦！",
            "台中家商國貿(68-76點)與幼保錄取門檻超高，排在竹山後面 100% 被作廢！"
        ],
        "solutions": [
            "【刪除重複竹山與苗栗三義】把通車時間留給自己好好生活與學習！",
            "【家商前移】家商國貿(01)➔幼保(02)➔霧農(03-04)➔竹山(05)！明德第 8~10 保底！"
        ],
        "warmth": "手滑打錯代碼別緊張！模擬選填就是幫你除錯，省下的通車時間拿來畫畫做甜點，高中更幸福！",
        "choices": [
            ("國立霧峰農工", "餐飲管理科", "公立技職 48-56點"),
            ("國立竹山高中", "廣告設計科", "公立美工 35-42點"),
            ("國立竹山高中", "廣告設計科", "手滑事故！完全重複連填"),
            ("台中家商", "國際貿易科", "商管頂尖 68-76點 倒填作廢"),
            ("縣立三義高中", "餐飲管理科", "苗栗山城 每日通勤遠距"),
            ("台中家商", "幼兒保育科", "公立名科 55-65點 倒填作廢"),
            ("國立霧峰農工", "電子科", "公立電機電子群 50-55點"),
            ("國立卓蘭高中", "餐飲管理科", "苗栗偏遠山區 每日奔波"),
            ("國立南投高中", "美工科", "公立美工設計科 35-42點"),
            ("私立明德高中", "幼兒保育科", "私立幼保 倒扣至11志願")
        ]
    },
    {
        "id": 9,
        "title": "斜槓青年的極限修煉！修車美容照服一人全包，電腦累壞",
        "risk": "⚠️ 五大職群天女散花",
        "badge_color": C_AMBER_ACCENT,
        "badge_bg": C_AMBER_LIGHT,
        "tag_p2": "💡 專家處方：鎖定優勢主軸 ＋ 僑泰衝刺慈明保底",
        "strength": "個性隨和、適應力極強，對世界抱持開放心態，能在不同環境中找到自己的安身處！",
        "anxiety_empathy": "覺得「有學校念就好」，對未來方向缺乏篤定感，不知道自己最擅長什麼。",
        "traps": [
            "5 大職群專業科目完全不相通，未來統測升科大跨考會非常吃力！",
            "第 1 志願慈明餐飲直接成交，後面填再多也是陪襯，無法審核較高分的僑泰應英。"
        ],
        "solutions": [
            "【深度性向三選一】最想做「造型美容」、「餐飲甜點」還是「外語實務」？",
            "【校風排位】僑泰放前面衝刺，慈明放後面保底；剔除不相干的汽修與長照！"
        ],
        "warmth": "你的隨和是很大優點，但一技之長要學一輩子！選一兩個你最喜歡的做到頂尖，路才會越走越寬！",
        "choices": [
            ("私立慈明高中", "餐飲管理科", "門檻約20點 首位成交"),
            ("私立慈明高中", "美容科", "時尚造型專業"),
            ("私立僑泰高中", "應用英語科", "外語實務 倒填遭截殺"),
            ("私立僑泰高中", "美容科", "市區知名美容科"),
            ("私立僑泰高中", "餐飲管理科", "知名餐飲實習教室"),
            ("私立青年高中", "餐飲管理科", "大里私立高職"),
            ("私立青年高中", "美容科", "時尚造型類科"),
            ("私立僑泰高中", "機械科", "硬核動力機械群"),
            ("私立僑泰高中", "汽車科", "車輛修護實習"),
            ("私立慈明高中", "照顧服務科", "長照醫護護理群")
        ]
    },
    {
        "id": 10,
        "title": "長億雙胞胎驚魂！連填兩次長億，東山排後面委屈大哭",
        "risk": "⚠️ 重複志願與嚴重倒填",
        "badge_color": C_AMBER_ACCENT,
        "badge_bg": C_AMBER_LIGHT,
        "tag_p2": "💡 專家處方：刪除長億雙胞胎 ＋ 東山前移至第1志願",
        "strength": "外語商管狂熱愛好者！前半段階梯排序行雲流水，教科書級的專業選填！",
        "anxiety_empathy": "前面超有章法，後面慌亂手滑連貼兩次長億，隨手補了東山高中。",
        "traps": [
            "長億連填兩次浪費名額；東山高中門檻 2A3B (70-74點)！",
            "東山排在長億(53-64點)後面，是標準倒填，永遠不可能被錄取！"
        ],
        "solutions": [
            "【刪除重複長億】將東山高中移至第 1 志願或緊隨家商之後！",
            "【完美範本】家商(01-02) ➔ 東山(03) ➔ 豐商(04-05) ➔ 投商(06) ➔ 長億(07保底)！"
        ],
        "warmth": "前面的商管階梯排得太漂亮了！只要移開搗蛋的長億雙胞胎，你的志願表就是滿分範本！",
        "choices": [
            ("台中家商", "應用英語科", "商職榜首 76-84點 頂尖首選"),
            ("台中家商", "國際貿易科", "公立名科 68-76點 商管王牌"),
            ("國立豐原高商", "應用英語科", "外語名科 52-62點"),
            ("國立豐原高商", "國際貿易科", "山線商管 50-58點"),
            ("國立南投高商", "應用英語科", "公立商職 40-50點"),
            ("國立南投高商", "國際貿易科", "公立商職 36-46點"),
            ("市立長億高中", "普通科", "屯區指標 53-64點 截流"),
            ("市立長億高中", "普通科", "手滑事故！完全重複連填"),
            ("市立東山高中", "普通科", "公立名校 70-74點 倒填作廢"),
            ("國立南投高中", "應用英語科", "公立高中職 35-45點 遠距")
        ]
    },
    {
        "id": 11,
        "title": "神裝丟在新手村後面的玩家！台中高工資電排第8/9志願",
        "risk": "🚨 極端錯置與私校阻斷",
        "badge_color": C_ROSE_ACCENT,
        "badge_bg": C_ROSE_LIGHT,
        "tag_p2": "💡 專家處方：家庭坦誠對話 ＋ 高工資電排第一",
        "strength": "兼具藝術表演的感性與理工科技的理性潛能，舞台魅力與思維兼備！",
        "anxiety_empathy": "自己熱愛表演跳舞，但長輩希望他讀高工資電賺錢，孩子夾在中間無奈交差。",
        "traps": [
            "台中高工資電需 4A1B~5A，排在門檻20點的新民後面，100% 不可能被審核！",
            "考 5A 滿點也只能去讀新民表藝，高工資電直接報廢！"
        ],
        "solutions": [
            "【家庭會議確認方向】若走科技，台中高工資電勇敢放第 01~02 志願！",
            "【若堅持表演】新民放前面，拿掉高工，改填家商幼保兼顧退路！"
        ],
        "warmth": "志願表不能做表面功夫！勇敢和爸媽坐下來聊，無論選哪條路，正正當當排第一才是真本事！",
        "choices": [
            ("私立新民高中", "表演藝術科", "私立表藝 約20點"),
            ("私立青年高中", "舞蹈科", "私立舞蹈專班"),
            ("私立青年高中", "電影電視科", "影視傳播實作"),
            ("台中家商", "幼兒保育科", "公立名科 55-65點 遭私校截"),
            ("私立明德高中", "幼兒保育科", "私立幼保專業"),
            ("市立東勢高工", "資訊科", "公立高工 45-52點"),
            ("市立東勢高工", "電子科", "公立高工 45-52點"),
            ("台中高工", "資訊科", "頂尖神科 71-92點 倒填作廢"),
            ("台中高工", "電子科", "頂尖神科 71-98點 倒填作廢"),
            ("私立明德高中", "普通科", "私立綜合高中部")
        ]
    },
    {
        "id": 12,
        "title": "大明與新民的鐵壁防線！第5/6志願卡大佛，家商6科全滅",
        "risk": "🚨 致命私校門神阻斷",
        "badge_color": C_ROSE_ACCENT,
        "badge_bg": C_ROSE_LIGHT,
        "tag_p2": "💡 專家處方：拆除私校水壩 ＋ 家商6科全數前移",
        "strength": "積極認真，填寫 15 個志願展現出強烈的求學意願與廣泛的探索熱忱！",
        "anxiety_empathy": "想讀公立也想讀名氣響亮的家商，但排志願憑直覺隨便插隊，把家商卡死。",
        "traps": [
            "大明新民門檻30-45點，長億沒上就直接被錄取！",
            "後面的台中家商(55-76點)，就算考了 3A2B 也只能含淚去讀大明！"
        ],
        "solutions": [
            "【拆除私校水壩】立刻將台中家商 6 大科系移至第 03~08 志願！",
            "【安全梯次】東山(01)➔市大里(02)➔家商(03-08)➔長億(09)；私校退到第11志願後！"
        ],
        "warmth": "你很有眼光挑的都是好學校！把家商請到前面，大明新民在後頭護駕，這樣才是最高明的排法！",
        "choices": [
            ("市立東山高中", "普通科", "社區名校 70-74點 衝刺"),
            ("市立長億高中", "普通科", "屯區指標 53-64點 落點"),
            ("市立新社高中", "普通科", "公立保底 31-36點 倒填"),
            ("市立大里高中", "普通科", "優質公立 61-74點 倒填遭截"),
            ("私立大明高中", "普通科", "門神路障 30-45點 截殺公立"),
            ("私立新民高中", "普通科", "門神路障 30-45點 截殺公立"),
            ("台中家商", "商經/國貿/資處", "公立頂尖 遭私校攔截作廢"),
            ("台中家商", "會計/服飾/幼保", "公立頂尖 遭私校攔截作廢"),
            ("私立僑泰高中", "普通科", "私立升學班 (至15個)"),
            ("私立明道中學", "普通科", "私立升學班 (至15個)")
        ]
    },
    {
        "id": 13,
        "title": "長億高中成最偉大攔截王！一腳踢飛文華，5A 看到會哭",
        "risk": "🚨 致命長億水壩攔截",
        "badge_color": C_ROSE_ACCENT,
        "badge_bg": C_ROSE_LIGHT,
        "tag_p2": "💡 專家處方：拔除長億水壩 ＋ 文華女中正面迎戰",
        "strength": "實力出類拔萃的女學霸！目標高遠，平時成績極度優秀，具備衝刺頂尖名校的雄厚實力！",
        "anxiety_empathy": "實力出眾但太害怕落榜，順手在第 4 志願塞了長億當護身符，變成攔路虎。",
        "traps": [
            "文華女生門檻 4A1B (82-85點)，竟排在長億(53-64點)後面！",
            "若考了 84 點，二中附中沒上，電腦第 4 志願直接錄取長億，文華瞬間蒸發！"
        ],
        "solutions": [
            "【拔除長億水壩】文華高中立刻移至第 3 志願！",
            "【梯次順暢】女中(01)➔附中(02)➔文華(03)➔二中(04)➔東山(05)➔長億(08保底)！"
        ],
        "warmth": "你的實力這麼棒，千萬別被恐懼綁架！長億是港灣不是路障，把文華放回第一線，放膽去飛！",
        "choices": [
            ("台中女中", "普通科", "頂尖首選 94-95點 夢幻衝刺"),
            ("國立興大附中", "普通科", "頂尖名校 86-94點 穩健落點"),
            ("台中二中", "普通科", "前段名校 69-74點 堅實防守"),
            ("市立長億高中", "普通科", "屯區指標 53-64點 致命第4攔截"),
            ("市立文華高中", "普通科", "前段名校 82-85點 遭長億踢飛"),
            ("市立東山高中", "普通科", "社區名校 70-74點 遭長億踢飛"),
            ("市立大里高中", "普通科", "優質公立 61-74點 遭長億踢飛"),
            ("國立豐原高中", "普通科", "山線名校 53-70點 遭長億踢飛"),
            ("市立后綜高中", "普通科", "社區高中 45-52點 遭長億踢飛"),
            ("市立新社高中", "普通科", "公立保底 31-36點 陸續至16")
        ]
    },
    {
        "id": 14,
        "title": "私校包圍的公立夢！附農家商塞在第13/14志願早被截胡",
        "risk": "🚨 公私嚴重顛倒倒填",
        "badge_color": C_ROSE_ACCENT,
        "badge_bg": C_ROSE_LIGHT,
        "tag_p2": "💡 專家處方：翻轉自卑心態 ＋ 國立名校昂首排第一",
        "strength": "對文創設計充滿嚮往，心思細膩純真，內心深處渴望考上好學校證明自己！",
        "anxiety_empathy": "自認考不上公立，心態自卑怯弱，把國立學校當許願池放最後，不知前 10 志願同分滿分。",
        "traps": [
            "前 10 志願都是 30 分滿分，填國立名校完全不扣分！",
            "明台青年放第 2~5 志願，就算超常發揮考了 2A 也只能含淚去讀明台！"
        ],
        "solutions": [
            "【翻轉自卑】附農觀光(01)➔附農餐飲(02)➔家商服飾(03)➔長億(04)！",
            "【滿分區保底】私立職校移至第 05~08 志願，在前 10 志願內安心保底！"
        ],
        "warmth": "別小看自己！制度給了你前10志願免費築夢的特權，昂首挺胸排第一，考上人生大不同！",
        "choices": [
            ("市立長億高中", "普通科", "屯區指標 53-64點 抄自範例"),
            ("私立明台高中", "室內設計科", "室內設計實務 約20點"),
            ("私立青年高中", "廣告設計科", "視覺設計實習"),
            ("私立僑泰高中", "廣告設計科", "知名設計類科"),
            ("私立慈明高中", "廣告設計科", "多媒體美工實習"),
            ("其他私立科系", "各類職業科", "陸續填至第12志願全是私校"),
            ("國立興大附農", "觀光事業科", "國立名校 60-68點 第13扣分作廢"),
            ("國立興大附農", "餐飲管理科", "國立名校 60-68點 第14扣分作廢"),
            ("台中家商", "流行服飾科", "公立頂尖 55-65點 第15扣分作廢")
        ]
    },
    {
        "id": 15,
        "title": "環球餐飲小當家！葳格高學費與水里深山公車受得了嗎？",
        "risk": "⚠️ 財務與交通極限挑戰",
        "badge_color": C_AMBER_ACCENT,
        "badge_bg": C_AMBER_LIGHT,
        "tag_p2": "💡 專家處方：確認家庭財務預算 ＋ 果斷刪除偏遠深山",
        "strength": "專業純度 100% 全填餐飲！對料理美食充滿赤誠與熱情，未來的金牌米其林大廚！",
        "anxiety_empathy": "全心全意熱愛做菜，但選校純看科系名字，完全沒看學雜費與通車距離。",
        "traps": [
            "葳格高中實習雜費高昂，若未與家長確認可能造成家庭經濟負擔！",
            "水里商工在南投深山濁水溪旁，每天搭客運翻山越嶺太折騰，睡眠嚴重不足！"
        ],
        "solutions": [
            "【確認家庭預算】有負擔則換明德/僑泰！",
            "【果斷刪除水里與玉山】附農(01)➔明德(02)➔僑泰(03)➔明道(04)➔青年(05)！"
        ],
        "warmth": "大廚需要充足睡眠練刀工，天天搭客運去深山手都晃了！留在市區好學校，爸媽放心你更精進！",
        "choices": [
            ("國立興大附農", "餐飲管理科", "國立頂尖 60-68點 首選衝刺"),
            ("私立葳格高中", "餐飲管理科", "雙語私校 高昂學費實習費"),
            ("私立僑泰高中", "餐飲管理科", "大里名校 設備新穎"),
            ("私立明道中學", "餐飲管理科", "私立名校 餐飲專班"),
            ("私立玉山高中", "餐飲管理科", "東勢深山 交通偏遠"),
            ("私立大明高中", "餐飲管理科", "大里私立高職"),
            ("私立明德高中", "餐飲管理科", "南區名校 餐旅老字號"),
            ("國立水里商工", "餐飲管理科", "南投深山濁水溪畔 遠距"),
            ("私立嘉陽高中", "餐飲管理科", "台中清水海線"),
            ("市立新社高中", "餐飲管理科", "公立高職 共14個全餐飲")
        ]
    },
    {
        "id": 16,
        "title": "長億攔截秀續集！第3志願放長億，文華惠文深情目送",
        "risk": "🚨 致命長億水壩攔截",
        "badge_color": C_ROSE_ACCENT,
        "badge_bg": C_ROSE_LIGHT,
        "tag_p2": "💡 專家處方：長億移至第10志願 ＋ 明星高中大步前進",
        "strength": "頂尖資優女生！心儀女中、文華、惠文等名校，具有出色的學習天賦與深厚底蘊！",
        "anxiety_empathy": "既想衝一級名校，又怕落榜怕得要命，慌亂中把長億放第 3 志願當急救避難所。",
        "traps": [
            "長億(53-64點)排第 3，後方的文華、二中、忠明、惠文全比長億高 10~25 點！",
            "考 3A2B (75點) 會在第 3 志願直接被長億鎖死，7 所明星高中集體報廢！"
        ],
        "solutions": [
            "【拔除第 3 志願長億】文華、惠文、二中、忠明全部往前推進！",
            "【黃金順序】女中(01)➔附中(02)➔文華(03)➔惠文(04)➔二中(05)➔忠明(06)➔長億放第10！"
        ],
        "warmth": "保底是用來墊背的，不是用來擋路的！把長億放最後守門，前面的名校隨你挑，相信三年努力！",
        "choices": [
            ("台中女中", "普通科", "頂尖首選 94-95點 夢幻首位"),
            ("國立興大附中", "普通科", "頂尖名校 86-94點 男女兼收"),
            ("市立長億高中", "普通科", "屯區指標 53-64點 致命第3水壩"),
            ("市立文華高中", "普通科", "前段名校 82-85點 遭長億攔截"),
            ("台中二中", "普通科", "前段名校 69-74點 遭長億攔截"),
            ("市立忠明高中", "普通科", "優質公立 71-76點 遭長億攔截"),
            ("市立惠文高中", "普通科", "前段名校 76-77點 遭長億攔截"),
            ("市立東山高中", "普通科", "社區名校 70-74點 遭長億攔截"),
            ("市立大里高中", "普通科", "優質公立 61-74點 遭長億攔截"),
            ("市立西苑高中", "普通科", "社區名校 70-74點 遭長億攔截")
        ]
    }
]

def add_common_slides_before(prs, blank_layout):
    # SLIDE 1: COVER
    s1 = prs.slides.add_slide(blank_layout)
    add_bg(s1, C_NAVY_DARK)

    bar = s1.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(1.2), Inches(1.4), Inches(0.18), Inches(4.5))
    bar.fill.solid()
    bar.fill.fore_color.rgb = C_BLUE_ACCENT
    bar.line.fill.background()

    tx1 = s1.shapes.add_textbox(Inches(1.6), Inches(1.3), Inches(10.5), Inches(4.8))
    tf1 = tx1.text_frame
    tf1.word_wrap = True

    p1 = tf1.paragraphs[0]
    p1.text = "中投區國三志願選填 專家深度診斷與戰略指南"
    p1.font.name = FONT_NAME
    p1.font.size = Pt(32)
    p1.font.bold = True
    p1.font.color.rgb = C_WHITE

    p1_sub = tf1.add_paragraph()
    p1_sub.text = "16 位國三生選填盲點拆解 ｜ 原件真跡放大 ＆ OCR 美化大字對照 ＆ 專家正面賦能"
    p1_sub.font.name = FONT_NAME
    p1_sub.font.size = Pt(19)
    p1_sub.font.color.rgb = RGBColor(148, 163, 184)
    p1_sub.space_before = Pt(14)

    p1_desc = tf1.add_paragraph()
    p1_desc.text = "「成績只是一時的標籤，選對策略才是致勝關鍵！」\n雙頁對照設計 ｜ 原始單據高清放大 ｜ OCR 粗大文字 ｜ 溫暖同理陪伴每位孩子"
    p1_desc.font.name = FONT_NAME
    p1_desc.font.size = Pt(15)
    p1_desc.font.color.rgb = RGBColor(226, 232, 240)
    p1_desc.space_before = Pt(18)

    tx1_ft = s1.shapes.add_textbox(Inches(1.6), Inches(6.2), Inches(10), Inches(0.5))
    tf1_ft = tx1_ft.text_frame
    p1_ft = tf1_ft.paragraphs[0]
    p1_ft.text = "規格：教室 70 吋大電視最佳化 ｜ 字體極大、字數精簡、色彩明亮舒服 ｜ PowerPoint & PDF 雙格式"
    p1_ft.font.name = FONT_NAME
    p1_ft.font.size = Pt(12.5)
    p1_ft.font.color.rgb = RGBColor(148, 163, 184)

    # SLIDE 2: CORE PHILOSOPHY
    s2 = prs.slides.add_slide(blank_layout)
    add_bg(s2, C_BG_LIGHT)
    add_header(s2, "專家核心觀念 ｜ 中投區免試入學分發機制", "為什麼「選填策略」比會考多拿兩題更具關鍵影響力？")

    cards_data = [
        ("⚖️ 電腦分發不可逆", "一旦符合立即成交，永不回頭！",
         "分發程式依照個人序位「由上而下逐一核對」。只要某一志願達標立刻敲定，後面的志願哪怕是一中、女中，電腦連看都不看！",
         C_ROSE_ACCENT, C_ROSE_LIGHT),
        ("🎯 多對兩題不如防一倒填", "考場多拿2題，不敵一個手滑！",
         "多對兩題頂多增加 1~2 點積點；但若發生「倒填」（低分放高分前面），能上名校卻被第 1 志願攔截，三年的苦讀瞬間成空！",
         C_AMBER_ACCENT, C_AMBER_LIGHT),
        ("🛡️ 充分利用 30 分滿分區", "前 10 志願同分，大膽衝刺零扣分！",
         "中投區前 10 志願皆得滿分 30 分！善用前 2~3 個名額大膽填寫夢幻名校，完全零風險、無倒扣懲罰，把分數價值發揮到極致！",
         C_EMERALD_ACC, C_EMERALD_LIGHT)
    ]

    for idx, (ct, cs, cd, ccol, cbg) in enumerate(cards_data):
        cleft = Inches(0.6 + idx * 4.1)
        card = create_card(s2, cleft, Inches(1.3), Inches(3.9), Inches(5.65), C_WHITE, ccol)
        
        tb = s2.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, cleft + Inches(0.2), Inches(1.5), Inches(3.5), Inches(0.9))
        tb.fill.solid()
        tb.fill.fore_color.rgb = cbg
        tb.line.color.rgb = ccol
        ttb = tb.text_frame
        ttb.vertical_anchor = MSO_ANCHOR.MIDDLE
        tp = ttb.paragraphs[0]
        tp.text = ct
        tp.font.name = FONT_NAME
        tp.font.size = Pt(16)
        tp.font.bold = True
        tp.font.color.rgb = ccol
        tp.alignment = PP_ALIGN.CENTER
        
        tx = s2.shapes.add_textbox(cleft + Inches(0.25), Inches(2.6), Inches(3.4), Inches(4.0))
        tf = tx.text_frame
        tf.word_wrap = True
        
        p_sub = tf.paragraphs[0]
        p_sub.text = f"【核心關鍵】\n{cs}"
        p_sub.font.name = FONT_NAME
        p_sub.font.size = Pt(13.5)
        p_sub.font.bold = True
        p_sub.font.color.rgb = C_NAVY_DARK
        p_sub.space_after = Pt(12)
        
        p_desc = tf.add_paragraph()
        p_desc.text = cd
        p_desc.font.name = FONT_NAME
        p_desc.font.size = Pt(13)
        p_desc.font.color.rgb = C_TEXT_MAIN

    # SLIDE 3: 3-YEAR OFFICIAL CUTOFF DATA TABLE
    s3 = prs.slides.add_slide(blank_layout)
    add_bg(s3, C_BG_LIGHT)
    add_header(s3, "客觀數據實證庫 ｜ 111~113 學年度中投區免試入學", "近三年會考各校錄取門檻、ABC等第與積點對照總表")

    t_shape = s3.shapes.add_table(6, 6, Inches(0.6), Inches(1.25), Inches(12.133), Inches(5.7))
    t = t_shape.table
    t.columns[0].width = Inches(1.6)
    t.columns[1].width = Inches(2.6)
    t.columns[2].width = Inches(1.7)
    t.columns[3].width = Inches(1.3)
    t.columns[4].width = Inches(1.3)
    t.columns[5].width = Inches(3.633)

    headers = ["梯隊類別", "代表學校 / 熱門科系", "會考積分 (等第)", "111年", "112~113年", "大螢幕專家選填重點摘要"]
    for c_idx, h_text in enumerate(headers):
        cell = t.cell(0, c_idx)
        cell.fill.solid()
        cell.fill.fore_color.rgb = C_NAVY_LIGHT
        cell.margin_left = cell.margin_right = Inches(0.06)
        p = cell.text_frame.paragraphs[0]
        p.alignment = PP_ALIGN.CENTER
        p.text = h_text
        p.font.name = FONT_NAME
        p.font.size = Pt(13)
        p.font.bold = True
        p.font.color.rgb = C_WHITE

    table_rows = [
        ("第一志願群\n(頂尖明星)", "台中一中 / 台中女中\n興大附中 / 中科實中\n台中高工(電機/資訊)", "30分 (5A)\n28~30分 (4A1B~5A)\n24~30分 (2A~5A)", "97 / 94點\n86 點\n70~80點", "95~97 點\n86~94 點\n71~98 點", "一中女中穩定需5A(4~5+)。高工電機資訊競爭極激烈，達4A1B~5A水準。"),
        ("第二志願群\n(前段名校)", "文華高中(男/女)\n惠文高中 / 台中二中\n台中家商(應英/國貿)", "28分 (4A1B)\n26~28分 (3A~4A)\n24~26分 (2A~3A)", "88 / 85點\n77 / 74點\n74~78點", "88 / 82點\n76 / 69點\n76~84 點", "文華男生88點/女生82~85點。惠文76點超車二中。家商應英需3A，商職榜首。"),
        ("第三志願群\n(優質公立)", "忠明高中 / 東山高中\n西苑高中 / 附農(觀餐)\n台中高工(機械/冷凍)", "26分 (3A2B)\n24分 (2A3B)\n20~24分 (5B+~2A)", "71 點\n71 / 70點\n56~62點", "76 點\n74 / 74 點\n60~71 點", "忠明攀升至76點。東山、西苑穩定在2A3B(70~74點)。附農觀光熱門需5B+~1A。"),
        ("第四志願群\n(社區與技術)", "市立大里 / 清水 / 長億\n豐原高中 / 豐原高商\n台中家商(資處/幼保)", "22~24分 (1A~2A)\n20~22分 (5B+~1A)\n20~22分 (5B+~1A)", "63 / 58點\n53 / 50點\n52~60點", "64~74 點\n58~70 點\n55~65 點", "市大里近年達74點。長億53~64點(5B+至1A)，常被誤放第1志願攔截名校！"),
        ("第五志願群\n(安全保底)", "后綜高中 / 新社高中\n草屯商工 / 南投高商\n各大私校 (普通/職校)", "20分 (5B)\n18~20分 (4B1C~5B)\n12~24分 (3C~2A)", "48 / 31點\n36~48點\n30~65點", "52 / 36 點\n40~52 點\n30~65 點", "新社為公立安全底線(31~36點)。私校菁英班需1A~3A獎學金，多數職校填報即錄取。")
    ]

    for r_idx, r_val in enumerate(table_rows):
        bg_r = C_ROW_ALT if r_idx % 2 == 1 else C_WHITE
        for c_idx in range(6):
            cell = t.cell(r_idx + 1, c_idx)
            cell.fill.solid()
            cell.fill.fore_color.rgb = bg_r
            cell.margin_left = cell.margin_right = Inches(0.06)
            p = cell.text_frame.paragraphs[0]
            p.text = r_val[c_idx]
            p.font.name = FONT_NAME
            p.font.size = Pt(12)
            p.font.bold = (c_idx in [0, 2])
            p.font.color.rgb = C_BLUE_ACCENT if c_idx == 0 else C_TEXT_MAIN
            if c_idx in [0, 2, 3, 4]:
                p.alignment = PP_ALIGN.CENTER

    # SLIDE 4: RULES & GOLDEN RATIO
    s4 = prs.slides.add_slide(blank_layout)
    add_bg(s4, C_BG_LIGHT)
    add_header(s4, "免試入學實戰心法 ｜ 攻守兼備架構", "志願選填三大鐵則 ＆ 前 10 志願「2 : 5 : 3」黃金比例")

    c_left = create_card(s4, Inches(0.6), Inches(1.25), Inches(5.9), Inches(5.75), C_WHITE, C_BORDER)
    tx_l = s4.shapes.add_textbox(Inches(0.85), Inches(1.4), Inches(5.4), Inches(5.4))
    tf_l = tx_l.text_frame
    tf_l.word_wrap = True

    p_lh = tf_l.paragraphs[0]
    p_lh.text = "📌 志願選填三大鐵則（違者必重罰！）"
    p_lh.font.name = FONT_NAME
    p_lh.font.size = Pt(16)
    p_lh.font.bold = True
    p_lh.font.color.rgb = C_NAVY_DARK
    p_lh.space_after = Pt(14)

    rules_list = [
        ("鐵則一：高分排前，低分排後（嚴禁倒填！）", "電腦第一輪達標立刻鎖死！先填長億保底，後面的一中女中文華全數作廢。"),
        ("鐵則二：公立放前面，私立退後段（杜絕私校門神）", "私立學校放第 3~5 志願會變攔截網，後面的公立高中職全部看得到吃不到。"),
        ("鐵則三：同校多科同序原則，專業職群集中深耕", "同一學校連續填不同科算同一個志願序！切忌跨 5 大不相干職群天女散花。")
    ]

    for rt, rd in rules_list:
        p_t = tf_l.add_paragraph()
        p_t.text = f"▶ {rt}"
        p_t.font.name = FONT_NAME
        p_t.font.size = Pt(14)
        p_t.font.bold = True
        p_t.font.color.rgb = C_BLUE_ACCENT
        p_t.space_before = Pt(8)
        
        p_d = tf_l.add_paragraph()
        p_d.text = rd
        p_d.font.name = FONT_NAME
        p_d.font.size = Pt(12.5)
        p_d.font.color.rgb = C_TEXT_MAIN

    c_right = create_card(s4, Inches(6.8), Inches(1.25), Inches(5.9), Inches(5.75), C_WHITE, C_BORDER)
    tx_r = s4.shapes.add_textbox(Inches(7.05), Inches(1.4), Inches(5.4), Inches(5.4))
    tf_r = tx_r.text_frame
    tf_r.word_wrap = True

    p_rh = tf_r.paragraphs[0]
    p_rh.text = "📊 前 10 志願滿分保護傘「2 : 5 : 3」配置"
    p_rh.font.name = FONT_NAME
    p_rh.font.size = Pt(16)
    p_rh.font.bold = True
    p_rh.font.color.rgb = C_NAVY_DARK
    p_rh.space_after = Pt(14)

    tiers_list = [
        ("🚀 夢幻衝刺區（志願 01 ~ 02，佔 20%）", "比平時落點高 1~2 點之理想名校。前 10 志願同分滿分，大膽衝刺零成本！"),
        ("🎯 實力落點區（志願 03 ~ 07，佔 50%）", "與會考序位或模擬考實力完全相符的學校。這是分發命中率最高的核心靶心！"),
        ("🛡️ 穩健保底區（志願 08 ~ 10，佔 30%）", "低於落點 3~5 點，100% 會上且願意去讀的學校。在前 10 志願內守住不被倒扣！"),
        ("🔒 安全防護網（志願 11 ~ 30+，填好填滿）", "最多可填 50 個志願！絕不只填 10 個，後段填滿才能徹底杜絕意外落榜。")
    ]

    for tt, td in tiers_list:
        p_t = tf_r.add_paragraph()
        p_t.text = tt
        p_t.font.name = FONT_NAME
        p_t.font.size = Pt(14)
        p_t.font.bold = True
        p_t.font.color.rgb = C_EMERALD_ACC
        p_t.space_before = Pt(8)
        
        p_d = tf_r.add_paragraph()
        p_d.text = td
        p_d.font.name = FONT_NAME
        p_d.font.size = Pt(12.5)
        p_d.font.color.rgb = C_TEXT_MAIN

    # SLIDE 5: BIG DATA DIAGNOSIS
    s5 = prs.slides.add_slide(blank_layout)
    add_bg(s5, C_BG_LIGHT)
    add_header(s5, "全班模擬卷大數據檢驗 ｜ 六大集體盲區", "全班普遍罹患的六大選填迷思：你是不是也不小心踩雷了？")

    big_pitfalls = [
        ("🚨 範例抄襲盲信症", "多達 8 位同學受害！", "看到考卷範例寫長億高中(064336)，就照抄進志願表，長億無辜變成全班最大攔截王！", C_ROSE_ACCENT, C_ROSE_LIGHT, C_ROSE_BORDER),
        ("⚠️ 倒吃甘蔗變成倒撞牆", "7 位同學排序完全顛倒！", "把低分學校排前面、一中文華排最後。電腦第 1 志願直接成交，後面的名校全部放水流！", C_ROSE_ACCENT, C_ROSE_LIGHT, C_ROSE_BORDER),
        ("🛑 私校門神一夫當關", "6 位同學公私順序顛倒！", "把大明、新民放第 3~5 志願，後方排了台中家商或東山。私校門神直接把公立名校全封死！", C_AMBER_ACCENT, C_AMBER_LIGHT, C_AMBER_BORDER),
        ("🌀 斜槓青年天女散花", "5 位同學跨 5 大職群！", "一張表同時填修車、美容、幼保、外語、電機！目標太發散，未來升學統測跨考會非常痛苦。", C_AMBER_ACCENT, C_AMBER_LIGHT, C_AMBER_BORDER),
        ("📉 高空鋼索零防護網", "11 位同學只填 10 個！", "明明可填 50 個卻只填 10 個，甚至全填 5A 明星高中。會考稍有失常，直接無校可讀落榜！", C_BLUE_ACCENT, C_BLUE_LIGHT, C_BLUE_BORDER),
        ("🚌 環台通勤大冒險", "3 位同學跨 3 縣市！", "為讀美容或餐飲，勾到彰化田中、苗栗卓蘭或南投水里。單趟車程 2 小時，天天累癱在車上！", C_BLUE_ACCENT, C_BLUE_LIGHT, C_BLUE_BORDER)
    ]

    for idx, (pt, ps, pd, pcol, pbg, pbor) in enumerate(big_pitfalls):
        col = idx % 3
        row = idx // 3
        cx = Inches(0.6 + col * 4.1)
        cy = Inches(1.35 + row * 2.85)
        
        card = create_card(s5, cx, cy, Inches(3.9), Inches(2.65), C_WHITE, pbor)
        
        hb = s5.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, cx + Inches(0.2), cy + Inches(0.2), Inches(3.5), Inches(0.75))
        hb.fill.solid()
        hb.fill.fore_color.rgb = pbg
        hb.line.color.rgb = pbor
        hbtf = hb.text_frame
        hbtf.vertical_anchor = MSO_ANCHOR.MIDDLE
        hp1 = hbtf.paragraphs[0]
        hp1.text = pt
        hp1.font.name = FONT_NAME
        hp1.font.size = Pt(14)
        hp1.font.bold = True
        hp1.font.color.rgb = pcol
        hp1.alignment = PP_ALIGN.CENTER
        
        tx = s5.shapes.add_textbox(cx + Inches(0.2), cy + Inches(1.05), Inches(3.5), Inches(1.5))
        tf = tx.text_frame
        tf.word_wrap = True
        
        p2 = tf.paragraphs[0]
        p2.text = f"【災情】{ps}"
        p2.font.name = FONT_NAME
        p2.font.size = Pt(12)
        p2.font.bold = True
        p2.font.color.rgb = C_NAVY_DARK
        p2.space_after = Pt(4)
        
        p3 = tf.add_paragraph()
        p3.text = pd
        p3.font.name = FONT_NAME
        p3.font.size = Pt(11.5)
        p3.font.color.rgb = C_TEXT_MAIN

def add_common_slides_after(prs, blank_layout):
    # SLIDE 38: PARENT-TEACHER WISDOM & EMPATHY
    s38 = prs.slides.add_slide(blank_layout)
    add_bg(s38, C_BG_LIGHT)
    add_header(s38, "親師陪伴與對話指引 ｜ 智慧化解選填焦慮", "給正在迷茫與缺乏自信同學的真心話：走對路比搶名校更發光！")

    pt_cards = [
        ("❤️ 接納自我，不以分數定終身", "會考只是一場分組遊戲，不是人生判決書",
         "成績落後並不代表你不如人，只是這場紙筆測驗不適合你的天賦。未來的社會需要各行各業的專家，在適合自己的跑道上深耕，你一樣能跑出第一名！",
         C_BLUE_ACCENT, C_BLUE_LIGHT, C_BLUE_BORDER),
        ("🔍 探索真實天賦，適性揚才", "雙手動得比筆快，技職是發光大舞台",
         "不要把文科孩子硬塞進電機工科，也不要把熱愛實作的孩子關在普通高中死背。108 課綱技職升學管道極暢通，學得一技之長，眼神有光、生活更有底氣！",
         C_EMERALD_ACC, C_EMERALD_LIGHT, C_EMERALD_BOR),
        ("🤝 勇敢為自己的選擇負責", "把最終選擇權與承擔的勇氣交給孩子",
         "志願表上的學校，未來三年是孩子每天要背著書包去讀的。父母與導師負責提供客觀數據與保底防護，讓孩子為自己的夢想掌舵，更能激發主動學習動能！",
         C_AMBER_ACCENT, C_AMBER_LIGHT, C_AMBER_BORDER)
    ]

    for idx, (ct, cs, cb, ccol, cbg, cbor) in enumerate(pt_cards):
        cleft = Inches(0.6 + idx * 4.1)
        card = create_card(s38, cleft, Inches(1.3), Inches(3.9), Inches(5.65), C_WHITE, cbor)
        
        tb = s38.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, cleft + Inches(0.2), Inches(1.5), Inches(3.5), Inches(0.9))
        tb.fill.solid()
        tb.fill.fore_color.rgb = cbg
        tb.line.color.rgb = ccol
        ttb = tb.text_frame
        ttb.vertical_anchor = MSO_ANCHOR.MIDDLE
        tp = ttb.paragraphs[0]
        tp.text = ct
        tp.font.name = FONT_NAME
        tp.font.size = Pt(15.5)
        tp.font.bold = True
        tp.font.color.rgb = ccol
        tp.alignment = PP_ALIGN.CENTER
        
        tx = s38.shapes.add_textbox(cleft + Inches(0.25), Inches(2.6), Inches(3.4), Inches(4.0))
        tf = tx.text_frame
        tf.word_wrap = True
        
        p_sub = tf.paragraphs[0]
        p_sub.text = f"【輔導心法】\n{cs}"
        p_sub.font.name = FONT_NAME
        p_sub.font.size = Pt(13.5)
        p_sub.font.bold = True
        p_sub.font.color.rgb = C_NAVY_DARK
        p_sub.space_after = Pt(12)
        
        p_desc = tf.add_paragraph()
        p_desc.text = cb
        p_desc.font.name = FONT_NAME
        p_desc.font.size = Pt(12.5)
        p_desc.font.color.rgb = C_TEXT_MAIN

    # SLIDE 39: SCORECARD & 5-STEP ACTION
    s39 = prs.slides.add_slide(blank_layout)
    add_bg(s39, C_BG_LIGHT)
    add_header(s39, "策略健檢總體檢 ｜ 全班急診分流", "16 位學生選填風險等級評分 ＆ 正式選填「五步重整急救包」")

    c_left = create_card(s39, Inches(0.6), Inches(1.2), Inches(6.0), Inches(5.8), C_WHITE, C_BORDER)
    tx_t = s39.shapes.add_textbox(Inches(0.8), Inches(1.35), Inches(5.6), Inches(5.4))
    tf_t = tx_t.text_frame
    tf_t.word_wrap = True
    p_th = tf_t.paragraphs[0]
    p_th.text = "📋 16 位同學選填風險分流榜"
    p_th.font.name = FONT_NAME
    p_th.font.size = Pt(15.5)
    p_th.font.bold = True
    p_th.font.color.rgb = C_NAVY_DARK
    p_th.space_after = Pt(10)

    score_groups = [
        ("🚨 紅燈急診區（致命倒填/私校門神/零保底）",
         "學生 #01, #02, #03, #05, #11, #12, #13, #14, #16\n➔ 現有志願表將導致名校全作廢或意外落榜，必須立刻反轉重排！",
         C_ROSE_ACCENT),
        ("⚠️ 黃燈觀察區（職群地雷/通勤過遠/重複連填）",
         "學生 #04, #06, #07, #08, #09, #10, #15\n➔ 方向大致正確但夾雜工科地雷或手滑重複，微調 2~3 個即可成完美榜樣！",
         C_AMBER_ACCENT)
    ]

    for sgt, sgd, sgcol in score_groups:
        p_gt = tf_t.add_paragraph()
        p_gt.text = sgt
        p_gt.font.name = FONT_NAME
        p_gt.font.size = Pt(13)
        p_gt.font.bold = True
        p_gt.font.color.rgb = sgcol
        p_gt.space_before = Pt(8)
        
        p_gd = tf_t.add_paragraph()
        p_gd.text = sgd
        p_gd.font.name = FONT_NAME
        p_gd.font.size = Pt(12)
        p_gd.font.color.rgb = C_TEXT_MAIN

    c_right = create_card(s39, Inches(6.9), Inches(1.2), Inches(5.8), Inches(5.8), C_WHITE, C_BORDER)
    tx_act = s39.shapes.add_textbox(Inches(7.15), Inches(1.35), Inches(5.3), Inches(5.4))
    tf_act = tx_act.text_frame
    tf_act.word_wrap = True

    p_ah = tf_act.paragraphs[0]
    p_ah.text = "🛠️ 正式選填「五步重整急救包」"
    p_ah.font.name = FONT_NAME
    p_ah.font.size = Pt(15.5)
    p_ah.font.bold = True
    p_ah.font.color.rgb = C_NAVY_DARK
    p_ah.space_after = Pt(10)

    steps = [
        ("第一步：核對三年官方錄取門檻", "依照 111~113 年積點由高到低排列，嚴禁任何倒填！"),
        ("第二步：拔除長億與私校門神", "長億放第 8~10 志願，私立學校一律置於公立之後保底。"),
        ("第三步：落實「2:5:3」黃金比例", "2個衝刺 ＋ 5個實力落點 ＋ 3個安全保底，兼顧夢想與防護。"),
        ("第四步：善用同校多科同序原則", "技職學生連續填寫同校不同科系，只計 1 個志願序積分。"),
        ("第五步：填滿志願徹底杜絕落榜", "中投區可填 50 個志願，至少填 20~30 個以上才交卷！")
    ]

    for st_title, st_desc in steps:
        p_st = tf_act.add_paragraph()
        p_st.text = f"▶ {st_title}"
        p_st.font.name = FONT_NAME
        p_st.font.size = Pt(13)
        p_st.font.bold = True
        p_st.font.color.rgb = C_EMERALD_ACC
        p_st.space_before = Pt(6)
        
        p_sd = tf_act.add_paragraph()
        p_sd.text = st_desc
        p_sd.font.name = FONT_NAME
        p_sd.font.size = Pt(12)
        p_sd.font.color.rgb = C_TEXT_MAIN

    # SLIDE 40: CONCLUSION
    s40 = prs.slides.add_slide(blank_layout)
    add_bg(s40, C_NAVY_DARK)

    bar40 = s40.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(1.2), Inches(1.5), Inches(0.18), Inches(4.5))
    bar40.fill.solid()
    bar40.fill.fore_color.rgb = C_EMERALD_ACC
    bar40.line.fill.background()

    tx40 = s40.shapes.add_textbox(Inches(1.6), Inches(1.4), Inches(10.5), Inches(4.8))
    tf40 = tx40.text_frame
    tf40.word_wrap = True

    p40_1 = tf40.paragraphs[0]
    p40_1.text = "勇於築夢，智慧選填，做自己人生的第一名！"
    p40_1.font.name = FONT_NAME
    p40_1.font.size = Pt(30)
    p40_1.font.bold = True
    p40_1.font.color.rgb = C_WHITE

    p40_2 = tf40.add_paragraph()
    p40_2.text = "會考成績不是人生的終點線，而是發掘自我潛能的起跑點。"
    p40_2.font.name = FONT_NAME
    p40_2.font.size = Pt(18)
    p40_2.font.color.rgb = RGBColor(148, 163, 184)
    p40_2.space_before = Pt(16)

    p40_3 = tf40.add_paragraph()
    p40_3.text = "「選你所愛，愛你所選。」\n當你用智慧排好了安全的防護網，就請昂首挺胸、大步向前！\n無論最終落腳哪一所學校，只要保持熱情與專注，你在那裡都會是最閃耀的星星！"
    p40_3.font.name = FONT_NAME
    p40_3.font.size = Pt(15.5)
    p40_3.font.color.rgb = RGBColor(226, 232, 240)
    p40_3.space_before = Pt(20)

    tx40_ft = s40.shapes.add_textbox(Inches(1.6), Inches(6.1), Inches(10), Inches(0.5))
    tf40_ft = tx40_ft.text_frame
    p40_ft = tf40_ft.paragraphs[0]
    p40_ft.text = "祝福所有國三畢業生：金榜題名 ｜ 揚帆起航 ｜ 前程似錦！"
    p40_ft.font.name = FONT_NAME
    p40_ft.font.size = Pt(13)
    p40_ft.font.bold = True
    p40_ft.font.color.rgb = C_AMBER_ACCENT

# =============================================================
# BUILD COMPLETE 40-SLIDE PRESENTATION
# =============================================================
def build_presentation(pptx_path, img_dir):
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    blank_layout = prs.slide_layouts[6]

    # Slides 1 to 5
    add_common_slides_before(prs, blank_layout)

    # 16 Students: 2 slides each (Total 32 slides)
    for st in STUDENTS_DATA:
        s_id = st["id"]

        # =============================================================
        # PAGE 1: 原始真跡放大 (左) ＋ OCR 美化粗大表格 (右)
        # =============================================================
        s_p1 = prs.slides.add_slide(blank_layout)
        add_bg(s_p1, C_BG_LIGHT)

        add_header(s_p1,
                   f"學生個案原始選填單對照 ｜ 學生編號 #{s_id:02d}",
                   f"【原始志願 vs OCR清晰大字對照】{st['title']}",
                   st["risk"], st["badge_color"], st["badge_bg"])

        # LEFT HALF: Enlarge Original Form Image (Width 5.5", Height 6.05")
        card_img = create_card(s_p1, Inches(0.6), Inches(1.15), Inches(5.5), Inches(6.05), C_WHITE, C_BORDER)
        
        # Header banner inside left card
        tx_ih = s_p1.shapes.add_textbox(Inches(0.7), Inches(1.22), Inches(5.3), Inches(0.30))
        tf_ih = tx_ih.text_frame
        tf_ih.margin_left = tf_ih.margin_top = tf_ih.margin_right = tf_ih.margin_bottom = 0
        p_ih = tf_ih.paragraphs[0]
        p_ih.text = "📋 原始模擬選填單真跡（原件放大檢視）"
        p_ih.font.name = FONT_NAME
        p_ih.font.size = Pt(12.5)
        p_ih.font.bold = True
        p_ih.font.color.rgb = C_NAVY_DARK

        img_file = os.path.join(img_dir, f"page_{s_id:02d}.png")
        if os.path.exists(img_file):
            from PIL import Image
            with Image.open(img_file) as im:
                im_w, im_h = im.size
            aspect = im_w / float(im_h)
            
            # Max bounding box inside left card: width 5.2", height 5.4"
            max_w, max_h = 5.2, 5.4
            if max_w / max_h > aspect: # taller
                target_h = max_h
                target_w = target_h * aspect
            else: # wider
                target_w = max_w
                target_h = target_w / aspect

            pos_x = 0.6 + (5.5 - target_w) / 2.0
            pos_y = 1.58 + (5.4 - target_h) / 2.0
            s_p1.shapes.add_picture(img_file, Inches(pos_x), Inches(pos_y), width=Inches(target_w), height=Inches(target_h))

        # RIGHT HALF: OCR Converted Clean Bold Table (Width 6.4", Height 6.05")
        card_tb = create_card(s_p1, Inches(6.35), Inches(1.15), Inches(6.38), Inches(6.05), C_WHITE, C_BORDER)

        tx_th = s_p1.shapes.add_textbox(Inches(6.55), Inches(1.22), Inches(6.0), Inches(0.32))
        tf_th = tx_th.text_frame
        tf_th.margin_left = tf_th.margin_top = tf_th.margin_right = tf_th.margin_bottom = 0
        p_th = tf_th.paragraphs[0]
        p_th.text = "🔍 原始志願序 OCR 辨識對照表（大字清晰版）"
        p_th.font.name = FONT_NAME
        p_th.font.size = Pt(13.5)
        p_th.font.bold = True
        p_th.font.color.rgb = C_BLUE_ACCENT

        num_choices = min(len(st["choices"]), 10)
        t_row_h = Inches(0.36) # larger row height for big bold text
        t_shape = s_p1.shapes.add_table(num_choices + 1, 3, Inches(6.5), Inches(1.58), Inches(6.08), t_row_h * (num_choices + 1))
        table = t_shape.table
        table.columns[0].width = Inches(0.75) # 志願序號
        table.columns[1].width = Inches(2.35) # 學校名稱
        table.columns[2].width = Inches(2.98) # 科別與落點門檻狀態

        headers = ["序號", "學校名稱", "填報科別 ＆ 門檻狀態"]
        for c_idx, h_text in enumerate(headers):
            cell = table.cell(0, c_idx)
            cell.fill.solid()
            cell.fill.fore_color.rgb = C_NAVY_LIGHT
            cell.margin_left = cell.margin_right = Inches(0.06)
            p = cell.text_frame.paragraphs[0]
            p.alignment = PP_ALIGN.CENTER
            p.text = h_text
            p.font.name = FONT_NAME
            p.font.size = Pt(12)
            p.font.bold = True
            p.font.color.rgb = C_WHITE

        for r_idx in range(num_choices):
            r_data = st["choices"][r_idx]
            cell_0 = table.cell(r_idx + 1, 0)
            cell_1 = table.cell(r_idx + 1, 1)
            cell_2 = table.cell(r_idx + 1, 2)

            bg_row = C_ROW_ALT if r_idx % 2 == 1 else C_WHITE
            for c in [cell_0, cell_1, cell_2]:
                c.fill.solid()
                c.fill.fore_color.rgb = bg_row
                c.margin_left = c.margin_right = Inches(0.06)

            p0 = cell_0.text_frame.paragraphs[0]
            p0.alignment = PP_ALIGN.CENTER
            p0.text = f"{r_idx+1:02d}"
            p0.font.name = FONT_NAME
            p0.font.size = Pt(13)
            p0.font.bold = True
            p0.font.color.rgb = C_BLUE_ACCENT

            p1 = cell_1.text_frame.paragraphs[0]
            p1.text = r_data[0]
            p1.font.name = FONT_NAME
            p1.font.size = Pt(13)
            p1.font.bold = True
            p1.font.color.rgb = C_TEXT_MAIN

            p2 = cell_2.text_frame.paragraphs[0]
            p2.text = f"{r_data[1]}  {r_data[2]}"
            p2.font.name = FONT_NAME
            p2.font.size = Pt(11.5)
            p2.font.bold = True
            if any(w in r_data[2] for w in ["倒填", "作廢", "遭", "截", "事故", "地雷", "水壩", "首位"]):
                p2.font.color.rgb = C_ROSE_ACCENT
            else:
                p2.font.color.rgb = C_TEXT_MUTED

        # Bottom Quick Summary Box inside Right Card
        box_stats_top = Inches(1.58) + t_row_h * (num_choices + 1) + Inches(0.12)
        box_stats_h = Inches(7.05) - box_stats_top
        if box_stats_h > Inches(0.5):
            box_stats = s_p1.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(6.5), box_stats_top, Inches(6.08), box_stats_h)
            box_stats.fill.solid()
            box_stats.fill.fore_color.rgb = C_BLUE_LIGHT
            box_stats.line.color.rgb = C_BLUE_BORDER
            stf = box_stats.text_frame
            stf.word_wrap = True
            stf.margin_left = stf.margin_right = Inches(0.12)
            stf.margin_top = stf.margin_bottom = Inches(0.06)

            sp = stf.paragraphs[0]
            sp.text = f"⚡ 原始選填特徵初診："
            sp.font.name = FONT_NAME
            sp.font.size = Pt(12)
            sp.font.bold = True
            sp.font.color.rgb = C_BLUE_ACCENT

            kp = stf.add_paragraph()
            kp.text = f"• {st['traps'][0]}"
            kp.font.name = FONT_NAME
            kp.font.size = Pt(11.5)
            kp.font.bold = True
            kp.font.color.rgb = C_TEXT_MAIN
            kp.space_before = Pt(2)

        # =============================================================
        # PAGE 2: 升學輔導專家深度診斷與正面賦能處方 (下一頁)
        # =============================================================
        s_p2 = prs.slides.add_slide(blank_layout)
        add_bg(s_p2, C_BG_LIGHT)

        add_header(s_p2,
                   f"升學輔導專家深度診斷與重整處方 ｜ 學生編號 #{s_id:02d}",
                   f"【專家正面賦能處方】{st['title']}",
                   st["tag_p2"], C_EMERALD_ACC, C_EMERALD_LIGHT)

        # 3 Spacious, Bright, Comfortable Cards across the full width (Width 12.133")
        # CARD 1: 🌟 老師看見你的優勢與特質 ＆ 深刻同理 (Height 1.70")
        c1 = create_card(s_p2, Inches(0.6), Inches(1.20), Inches(12.133), Inches(1.70), C_BLUE_LIGHT, C_BLUE_BORDER)
        tx1 = s_p2.shapes.add_textbox(Inches(0.85), Inches(1.30), Inches(11.6), Inches(1.50))
        tf1 = tx1.text_frame
        tf1.word_wrap = True

        p1_h = tf1.paragraphs[0]
        p1_h.text = "🌟 老師看見你的亮點 ＆ 深刻同理孩子的心聲"
        p1_h.font.name = FONT_NAME
        p1_h.font.size = Pt(15.5)
        p1_h.font.bold = True
        p1_h.font.color.rgb = C_BLUE_ACCENT
        p1_h.space_after = Pt(4)

        p1_str = tf1.add_paragraph()
        p1_str.text = f"【你的優勢潛能】{st['strength']}"
        p1_str.font.name = FONT_NAME
        p1_str.font.size = Pt(13.5)
        p1_str.font.bold = True
        p1_str.font.color.rgb = C_NAVY_DARK
        p1_str.space_after = Pt(3)

        p1_anx = tf1.add_paragraph()
        p1_anx.text = f"【懂你的處境】{st['anxiety_empathy']}"
        p1_anx.font.name = FONT_NAME
        p1_anx.font.size = Pt(12.5)
        p1_anx.font.color.rgb = C_TEXT_MUTED

        # CARD 2: ⚠️ 演算法地雷避坑指南 (Height 1.85")
        c2 = create_card(s_p2, Inches(0.6), Inches(3.05), Inches(12.133), Inches(1.85), C_ROSE_LIGHT, C_ROSE_BORDER)
        tx2 = s_p2.shapes.add_textbox(Inches(0.85), Inches(3.15), Inches(11.6), Inches(1.65))
        tf2 = tx2.text_frame
        tf2.word_wrap = True

        p2_h = tf2.paragraphs[0]
        p2_h.text = "⚠️ 演算法地雷警示：別讓這張表沒收了你三年的努力！"
        p2_h.font.name = FONT_NAME
        p2_h.font.size = Pt(15.5)
        p2_h.font.bold = True
        p2_h.font.color.rgb = C_ROSE_ACCENT
        p2_h.space_after = Pt(4)

        for tp in st["traps"]:
            p_tp = tf2.add_paragraph()
            p_tp.text = f"❌ {tp}"
            p_tp.font.name = FONT_NAME
            p_tp.font.size = Pt(13)
            p_tp.font.bold = True
            p_tp.font.color.rgb = C_TEXT_MAIN
            p_tp.space_after = Pt(3)

        # CARD 3: 🚀 專家量身打造：最強黃金重整步法 ＆ 愛的叮嚀 (Height 2.15")
        c3 = create_card(s_p2, Inches(0.6), Inches(5.05), Inches(12.133), Inches(2.15), C_EMERALD_LIGHT, C_EMERALD_BOR)
        tx3 = s_p2.shapes.add_textbox(Inches(0.85), Inches(5.15), Inches(11.6), Inches(1.95))
        tf3 = tx3.text_frame
        tf3.word_wrap = True

        p3_h = tf3.paragraphs[0]
        p3_h.text = "🚀 專家量身打造：最強黃金重整對策 ＆ 老師愛的叮嚀"
        p3_h.font.name = FONT_NAME
        p3_h.font.size = Pt(15.5)
        p3_h.font.bold = True
        p3_h.font.color.rgb = C_EMERALD_ACC
        p3_h.space_after = Pt(4)

        for sp in st["solutions"]:
            p_sp = tf3.add_paragraph()
            p_sp.text = f"✅ {sp}"
            p_sp.font.name = FONT_NAME
            p_sp.font.size = Pt(13.5)
            p_sp.font.bold = True
            p_sp.font.color.rgb = C_NAVY_DARK
            p_sp.space_after = Pt(3)

        p3_w = tf3.add_paragraph()
        p3_w.text = f"❤️ 老師愛的叮嚀：「{st['warmth']}」"
        p3_w.font.name = FONT_NAME
        p3_w.font.size = Pt(13)
        p3_w.font.bold = True
        p3_w.font.color.rgb = RGBColor(180, 83, 9)
        p3_w.space_before = Pt(3)

    # Summary Slides 38 to 40
    add_common_slides_after(prs, blank_layout)

    prs.save(pptx_path)
    print(f"Successfully generated 40-slide presentation at {pptx_path}")

# =============================================================
# EXPORT TO PDF VIA POWERPOINT COM
# =============================================================
def convert_pptx_to_pdf_com(pptx_path, pdf_path):
    print(f"Converting {pptx_path} to PDF via PowerPoint COM...")
    abs_pptx = os.path.abspath(pptx_path)
    abs_pdf = os.path.abspath(pdf_path)

    ppt = comtypes.client.CreateObject("PowerPoint.Application")
    ppt.Visible = 1
    pres = ppt.Presentations.Open(abs_pptx)
    pres.SaveAs(abs_pdf, 32) # 32 = ppSaveAsPDF
    pres.Close()
    ppt.Quit()
    print(f"Successfully exported {abs_pdf} via PowerPoint COM!")

if __name__ == "__main__":
    img_dir = "G:/我的雲端硬碟/antigravity2026 test/志願選填分析/rotated_pages"
    dest_dir = "G:/我的雲端硬碟/antigravity2026 test/志願選填分析"

    # Main output file
    target_pptx = os.path.join(dest_dir, "國三升學志願選填_專家深度診斷與策略分析.pptx")
    target_pdf  = os.path.join(dest_dir, "國三升學志願選填_專家深度診斷與策略分析.pdf")

    build_presentation(target_pptx, img_dir)
    convert_pptx_to_pdf_com(target_pptx, target_pdf)
    print("ALL DONE PERFECTLY!")
