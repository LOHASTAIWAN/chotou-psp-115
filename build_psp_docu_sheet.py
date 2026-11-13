# -*- coding: utf-8 -*-
"""
115學年度 中投區高級中等學校實用技能學程（實技班）
招生簡章 DocuSheet & Markdown & CSV 生成器
"""
import os
import json
import csv

DATA_SCHOOLS = [
    # 臺中市公立
    {
        "district": "臺中市", "system": "公立(市立)", "name": "臺中市立臺中家事商業高級中等學校",
        "short_name": "臺中家商", "zone": "東區", "is_organizer": True,
        "programs": [
            {"dept": "烘焙食品科", "cluster": "食品群", "shift": "夜間上課", "quota": 35, "is_plumbing": False, "note": "進修部授課，三年免學費"},
            {"dept": "服裝製作科", "cluster": "設計群", "shift": "夜間上課", "quota": 35, "is_plumbing": False, "note": "服裝打版、車縫實務"},
            {"dept": "餐飲技術科", "cluster": "餐旅群", "shift": "夜間上課", "quota": 35, "is_plumbing": False, "note": "中餐烹調、飲料調製檢定"}
        ]
    },
    {
        "district": "臺中市", "system": "公立(國立)", "name": "國立中興大學附屬臺中高級農業職業學校",
        "short_name": "興大附農", "zone": "東區", "is_organizer": False,
        "programs": [
            {"dept": "園藝技術科", "cluster": "農業群", "shift": "日間上課", "quota": 35, "is_plumbing": False, "note": "造園景觀、植物培育實作"},
            {"dept": "烘焙食品科", "cluster": "食品群", "shift": "日間上課", "quota": 35, "is_plumbing": False, "note": "麵包西點實作、烘焙證照"},
            {"dept": "餐飲技術科", "cluster": "餐旅群", "shift": "日間上課", "quota": 35, "is_plumbing": False, "note": "餐旅服務、基礎烹飪技能"},
            {"dept": "觀光事務科", "cluster": "餐旅群", "shift": "日間上課", "quota": 35, "is_plumbing": False, "note": "觀光導覽、飯店櫃台實務"},
            {"dept": "烘焙食品科", "cluster": "食品群", "shift": "夜間上課", "quota": 35, "is_plumbing": False, "note": "夜間進修，免學費專班"}
        ]
    },
    {
        "district": "臺中市", "system": "公立(市立)", "name": "臺中市立臺中工業高級中等學校",
        "short_name": "臺中高工", "zone": "南區", "is_organizer": False,
        "programs": [
            {"dept": "機械加工科", "cluster": "機械群", "shift": "夜間上課", "quota": 35, "is_plumbing": False, "note": "CNC銑床車床操作、金屬加工實務"},
            {"dept": "電腦繪圖科", "cluster": "機械群", "shift": "夜間上課", "quota": 35, "is_plumbing": False, "note": "AutoCAD/SolidWorks 2D/3D繪圖"},
            {"dept": "電機修護科", "cluster": "電機與電子群", "shift": "夜間上課", "quota": 35, "is_plumbing": False, "note": "電工儀表、家電維修、低壓配電"},
            {"dept": "營造技術科", "cluster": "土木與建築群", "shift": "夜間上課", "quota": 35, "is_plumbing": False, "note": "泥水鋪面、測量、鋼筋模板實作"}
        ]
    },
    {
        "district": "臺中市", "system": "公立(市立)", "name": "臺中市立東勢工業高級中等學校",
        "short_name": "東勢高工", "zone": "東勢區", "is_organizer": False,
        "programs": [
            {"dept": "水電技術科", "cluster": "電機與電子群", "shift": "夜間上課", "quota": 35, "is_plumbing": True, "note": "★【水電專門】公立高工配電工場，室內配線/自來水管配管技能檢定，夜間授課半工半讀首選"},
            {"dept": "電腦繪圖科", "cluster": "機械群", "shift": "日間上課", "quota": 35, "is_plumbing": False, "note": "產學導向專班，電腦機械輔助設計"},
            {"dept": "裝潢技術科", "cluster": "設計群", "shift": "夜間上課", "quota": 35, "is_plumbing": False, "note": "室內裝修木工、傢俱製作技術"}
        ]
    },
    {
        "district": "臺中市", "system": "公立(市立)", "name": "臺中市立沙鹿工業高級中等學校",
        "short_name": "沙鹿高工", "zone": "沙鹿區", "is_organizer": False,
        "programs": [
            {"dept": "商用資訊科", "cluster": "商業與管理群", "shift": "日間上課", "quota": 35, "is_plumbing": False, "note": "海線商管實習、電商文書應用"},
            {"dept": "機械加工科", "cluster": "機械群", "shift": "日間上課", "quota": 35, "is_plumbing": False, "note": "傳統與數值控制機械實務"},
            {"dept": "紡織技術科", "cluster": "家政/工業群", "shift": "日間上課", "quota": 35, "is_plumbing": False, "note": "紡織纖維加工、染整印花技術"}
        ]
    },
    {
        "district": "臺中市", "system": "公立(市立)", "name": "臺中市立霧峰農業工業高級中等學校",
        "short_name": "霧峰農工", "zone": "霧峰區", "is_organizer": False,
        "programs": [
            {"dept": "餐飲技術科", "cluster": "餐旅群", "shift": "日間上課", "quota": 35, "is_plumbing": False, "note": "中西餐料理技巧、烘焙基礎"},
            {"dept": "機械加工科", "cluster": "機械群", "shift": "夜間上課", "quota": 35, "is_plumbing": False, "note": "金屬切削、機械量測與加工維護"},
            {"dept": "園藝技術科", "cluster": "農業群", "shift": "日間上課", "quota": 35, "is_plumbing": False, "note": "精緻農業花藝、庭園管理"}
        ]
    },
    # 臺中市私立
    {
        "district": "臺中市", "system": "私立", "name": "臺中市私立青年高級中學",
        "short_name": "青年高中", "zone": "大里區", "is_organizer": False,
        "programs": [
            {"dept": "水電技術科", "cluster": "電機與電子群", "shift": "日間上課", "quota": 70, "is_plumbing": True, "note": "★【水電旗艦】中部日間部水電指標！國家級室內配線檢定場，結合智慧宅、太陽能光電與水電工程產學直通，輔導雙乙級證照與科技大學升學"},
            {"dept": "烘焙食品科", "cluster": "餐旅群", "shift": "日間上課", "quota": 35, "is_plumbing": False, "note": "專業烘焙實習教室，中西點證照"},
            {"dept": "餐飲技術科", "cluster": "餐旅群", "shift": "日間上課", "quota": 35, "is_plumbing": False, "note": "餐飲實習與丙級證照輔導"},
            {"dept": "汽車修護科", "cluster": "動力機械群", "shift": "日間上課", "quota": 35, "is_plumbing": False, "note": "汽車保養底盤檢修與引擎技術"}
        ]
    },
    {
        "district": "臺中市", "system": "私立", "name": "臺中市私立致用高級中學",
        "short_name": "致用高中", "zone": "大甲區", "is_organizer": False,
        "programs": [
            {"dept": "水電技術科", "cluster": "電機與電子群", "shift": "日間上課", "quota": 35, "is_plumbing": True, "note": "★【水電專門】臺中海線水電與機電重要據點！室內配線、自來水配管及冷凍空調設備整合實作，在地水電工程企業合作推薦就業"},
            {"dept": "餐飲技術科", "cluster": "餐旅群", "shift": "日間上課", "quota": 35, "is_plumbing": False, "note": "海線餐飲烘焙職群實訓"},
            {"dept": "烘焙食品科", "cluster": "食品群", "shift": "日間上課", "quota": 35, "is_plumbing": False, "note": "中西點麵包製作及證照取得"},
            {"dept": "汽車修護科", "cluster": "動力機械群", "shift": "日間上課", "quota": 35, "is_plumbing": False, "note": "汽機車修護與電系診斷"}
        ]
    },
    {
        "district": "臺中市", "system": "私立", "name": "臺中市私立僑泰高級中學",
        "short_name": "僑泰高中", "zone": "大里區", "is_organizer": False,
        "programs": [
            {"dept": "寵物經營科", "cluster": "農業群", "shift": "日間上課", "quota": 35, "is_plumbing": False, "note": "熱門寵物美容、門市管理實務"},
            {"dept": "汽車修護科", "cluster": "動力機械群", "shift": "日間上課", "quota": 35, "is_plumbing": False, "note": "現代化汽修設備與電子檢測"},
            {"dept": "餐飲技術科", "cluster": "餐旅群", "shift": "日間上課", "quota": 35, "is_plumbing": False, "note": "廚藝刀工實習與宴會服務"},
            {"dept": "美顏技術科", "cluster": "美容造型群", "shift": "日間上課", "quota": 35, "is_plumbing": False, "note": "彩妝造形設計、護膚手藝"},
            {"dept": "美髮技術科", "cluster": "美容造型群", "shift": "日間上課", "quota": 35, "is_plumbing": False, "note": "剪染燙髮技巧、美髮乙丙級檢定"}
        ]
    },
    {
        "district": "臺中市", "system": "私立", "name": "臺中市私立明台高級中學",
        "short_name": "明台高中", "zone": "霧峰區", "is_organizer": False,
        "programs": [
            {"dept": "寵物經營科", "cluster": "農業群", "shift": "日間上課", "quota": 35, "is_plumbing": False, "note": "寵物美容造形與照護管理"},
            {"dept": "烘焙食品科", "cluster": "食品群", "shift": "日間上課", "quota": 35, "is_plumbing": False, "note": "法式西點麵包實作教室"},
            {"dept": "餐飲技術科", "cluster": "餐旅群", "shift": "日間上課", "quota": 35, "is_plumbing": False, "note": "名廚實習教室、飯店廚務技能"},
            {"dept": "美顏技術科", "cluster": "美容造型群", "shift": "日間上課", "quota": 35, "is_plumbing": False, "note": "新娘秘書、整體造型彩繪"},
            {"dept": "美髮技術科", "cluster": "美容造型群", "shift": "日間上課", "quota": 35, "is_plumbing": False, "note": "沙龍實習沙龍剪燙技術"},
            {"dept": "表演技術科", "cluster": "藝術群", "shift": "日間上課", "quota": 35, "is_plumbing": False, "note": "流行舞蹈、戲劇演藝實務"}
        ]
    },
    {
        "district": "臺中市", "system": "私立", "name": "臺中市私立嶺東高級中學",
        "short_name": "嶺東高中", "zone": "南屯區", "is_organizer": False,
        "programs": [
            {"dept": "汽車修護科", "cluster": "動力機械群", "shift": "日間上課", "quota": 35, "is_plumbing": False, "note": "結合嶺東科大資源，汽機車檢修實作"},
            {"dept": "餐飲技術科", "cluster": "餐旅群", "shift": "日間上課", "quota": 35, "is_plumbing": False, "note": "餐飲實習與多元料理研製"}
        ]
    },
    {
        "district": "臺中市", "system": "私立", "name": "臺中市私立慈明高級中學",
        "short_name": "慈明高中", "zone": "太平區", "is_organizer": False,
        "programs": [
            {"dept": "餐飲技術科", "cluster": "餐旅群", "shift": "日間上課", "quota": 35, "is_plumbing": False, "note": "飯店中西餐廚藝及飲調證照"},
            {"dept": "烘焙食品科", "cluster": "食品群", "shift": "日間上課", "quota": 35, "is_plumbing": False, "note": "烘焙實習教室與手工餅乾西點"},
            {"dept": "美顏技術科", "cluster": "美容造型群", "shift": "日間上課", "quota": 35, "is_plumbing": False, "note": "美甲美睫彩妝實用技藝"}
        ]
    },
    {
        "district": "臺中市", "system": "私立", "name": "臺中市私立大明高級中學",
        "short_name": "大明高中", "zone": "大里區", "is_organizer": False,
        "programs": [
            {"dept": "商用資訊科", "cluster": "商業與管理群", "shift": "日間上課", "quota": 35, "is_plumbing": False, "note": "商務套裝軟體應用、自媒體經營基礎"},
            {"dept": "廣告技術科", "cluster": "設計群", "shift": "日間上課", "quota": 35, "is_plumbing": False, "note": "數位美工排版、廣告視覺製作"}
        ]
    },
    {
        "district": "臺中市", "system": "私立", "name": "臺中市私立嘉陽高級中學",
        "short_name": "嘉陽高中", "zone": "清水區", "is_organizer": False,
        "programs": [
            {"dept": "飛機修護科", "cluster": "動力機械群", "shift": "日間上課", "quota": 35, "is_plumbing": False, "note": "中部唯一直升機與固定翼修護實習場"},
            {"dept": "汽車修護科", "cluster": "動力機械群", "shift": "日間上課", "quota": 35, "is_plumbing": False, "note": "車輛維護與電機裝修"},
            {"dept": "餐飲技術科", "cluster": "餐旅群", "shift": "日間上課", "quota": 35, "is_plumbing": False, "note": "特色海線風味廚藝技能"}
        ]
    },
    # 南投縣公立
    {
        "district": "南投縣", "system": "公立(國立)", "name": "國立埔里高級工業職業學校",
        "short_name": "埔里高工", "zone": "埔里鎮", "is_organizer": False,
        "programs": [
            {"dept": "電腦繪圖科", "cluster": "機械群", "shift": "日間上課", "quota": 35, "is_plumbing": False, "note": "南投山區機械繪圖與3D打樣實務"},
            {"dept": "微電腦修護科", "cluster": "電機與電子群", "shift": "日間上課", "quota": 35, "is_plumbing": False, "note": "電腦組裝維修、網路配線與單晶片"},
            {"dept": "化工技術科", "cluster": "化工群", "shift": "日間上課", "quota": 35, "is_plumbing": False, "note": "化學品檢驗、實驗室分析技術"}
        ]
    },
    {
        "district": "南投縣", "system": "公立(國立)", "name": "國立水里高級商工職業學校",
        "short_name": "水里商工", "zone": "水里鄉", "is_organizer": False,
        "programs": [
            {"dept": "餐飲技術科", "cluster": "餐旅群", "shift": "日間上課", "quota": 35, "is_plumbing": False, "note": "南投在地觀光餐旅人才培訓、中餐及烘焙"}
        ]
    },
    {
        "district": "南投縣", "system": "公立(國立)", "name": "國立草屯高級商業職業學校",
        "short_name": "草屯商工", "zone": "草屯鎮", "is_organizer": False,
        "programs": [
            {"dept": "商用資訊科", "cluster": "商業與管理群", "shift": "日間上課", "quota": 35, "is_plumbing": False, "note": "南投商業管理實用技術、文書與資訊系統"},
            {"dept": "配管技術科", "cluster": "機械/配管群", "shift": "夜間上課", "quota": 35, "is_plumbing": False, "note": "管路配置與水氣管線焊接實作"}
        ]
    },
    # 南投縣私立
    {
        "district": "南投縣", "system": "私立", "name": "南投縣私立同德高級中等學校",
        "short_name": "同德高中", "zone": "草屯鎮", "is_organizer": False,
        "programs": [
            {"dept": "餐飲技術科", "cluster": "餐旅群", "shift": "日間上課", "quota": 35, "is_plumbing": False, "note": "五星級餐旅實習工廠、中西廚藝與飲調"},
            {"dept": "烘焙食品科", "cluster": "食品群", "shift": "日間上課", "quota": 35, "is_plumbing": False, "note": "歐式麵包甜點烘焙專案"},
            {"dept": "美顏技術科", "cluster": "美容造型群", "shift": "日間上課", "quota": 35, "is_plumbing": False, "note": "時尚彩妝美甲整體造形"}
        ]
    }
]

PLUMBING_SCHOOLS_DETAILS = [
    {
        "school": "臺中市私立青年高級中學",
        "short": "青年高中",
        "district": "臺中市大里區中興路一段138號",
        "shift": "日間上課（日班）",
        "quota": "約 70 名（2 班）",
        "equipment": "國家級室內配線甲、乙、丙級檢定合格術科考場，專屬智慧建築電力控制與太陽能光電模擬實驗室",
        "curriculum": "電工法規、室內配線實務、工業配線、自來水管配管、智慧居家照明控制、太陽能光電裝設、網路與弱電系統架設、變壓器裝修實作",
        "certifications": "室內配線丙級/乙級（乙種/甲種電匠）、自來水管配管丙級、工業配線丙級、網路架設丙級",
        "partnerships": "與大臺中水電機電工程同業公會、知名甲級水電與機電工程顧問公司產學深度合作，定期業界技師駐校指導，畢業立即推薦高薪就業",
        "advantages": "中部首屈一指之日間部水電技術標竿！三年作息正常、不熬夜，學生考取雙乙級證照比例高。畢業後除直接就業月入四至六萬起跳，亦能藉由四技二專技優甄審進入國立科大（勤益科大、雲科大、虎尾科大等）"
    },
    {
        "school": "臺中市私立致用高級中學",
        "short": "致用高中",
        "district": "臺中市大甲區甲東路100號",
        "shift": "日間上課（日班）",
        "quota": "約 35 名（1 班）",
        "equipment": "海線配電與自來水管實習工廠、水電配管與冷凍空調配線整合實驗室、低壓配電練習盤",
        "curriculum": "基礎電學、低壓室內配線、給排水配管工程、居家水電衛浴設備維修、冷凍空調配電、建築識圖與繪圖、工安法令規範",
        "certifications": "室內配線丙級、自來水管配管丙級、用電設備檢驗丙級、冷凍空調裝修丙級",
        "partnerships": "與大甲、清水、沙鹿、大安在地水電工程同業公會及海線綠能離岸風電周邊機電協力廠緊密合作，提供就近實習與保證就業",
        "advantages": "照顧海線在地有心鑽研水電技術的學子，省去遠赴市區通勤的時間成本；教學注重實作手感與維修實戰，畢業即具備獨立承攬住宅水電改修與水電行技師水準"
    },
    {
        "school": "臺中市立東勢工業高級中等學校",
        "short": "東勢高工",
        "district": "臺中市東勢區東新里東關路六段1328號",
        "shift": "夜間上課（進修部）",
        "quota": "約 35 名（1 班）",
        "equipment": "公立技術型高中頂級工科實習工場、專業配線盤配電台、自來水配管工作母機、高壓電力示範設備",
        "curriculum": "室內配線實作、水電配管實習、電工機械原理、識圖與製圖、配電盤配線、工業配線實務、安全衛生法令",
        "certifications": "室內配線丙級/乙級、自來水管配管丙級",
        "partnerships": "結合公立高工優良聲譽，銜接山城與中部營造水電工程公司白天學徒與技工工作",
        "advantages": "公立學校學雜費全額補助（三年免學費），夜間上課模式讓學生白天可於水電水暖公司或工地當學徒賺取工資，晚上回校取得公立高工畢業文憑與國家技術士證照，半工半讀減輕家庭經濟負擔首選"
    }
]

TIMELINE_ITEMS = [
    {"stage": "簡章公告與核定", "date": "115年1月中旬", "desc": "教育部核定招生名額與簡章，主委學校（臺中家商）與各校同步公告", "target": "全體考生/家長"},
    {"stage": "國中校內輔導與線上填報", "date": "115年5月14日(四) ～ 5月20日(三)", "desc": "各國中輔導室辦理報名指導，學生完成線上志願選填（上限7個志願）並列印簽名A/B報名表", "target": "應屆畢業生"},
    {"stage": "分區受理集體與個別報名", "date": "115年5月21日(四) ～ 5月22日(五)", "desc": "各國中集體向臺中家商送件；非應屆或同等學力者於此時段親送臺中家商個別報名", "target": "各國中/個別報名考生"},
    {"stage": "錄取分發結果公告（放榜）", "date": "115年6月11日(四) 上午 10:00", "desc": "中投區分發小組及臺中家商官網開放線上查榜，各招生高職同步公告錄取榜單", "target": "全體報名學生"},
    {"stage": "分發結果複查申請", "date": "115年6月11日(四) 下午 17:00 前", "desc": "如對分發結果有疑義，持簡章規定之複查申請書向臺中家商提出複查（逾期不受理）", "target": "有疑義考生"},
    {"stage": "錄取學生至各校報到", "date": "115年6月12日(五) 14:00 ～ 16:00", "desc": "錄取生持國中畢業證書或同等學力證明正本至各錄取高中職辦理現場報到", "target": "已錄取學生"},
    {"stage": "聲明放棄錄取資格截止", "date": "115年6月15日(一) 中午 12:00 前", "desc": "已報到欲改參加中投區免試入學或其他升學管道者，須於期限前至錄取學校辦理放棄手續", "target": "欲轉報他管道者"},
    {"stage": "各校辦理實用技能學程續招", "date": "115年6月下旬起", "desc": "若招生學校尚有缺額，由各校自行公告辦理免試獨立續招，額滿為止", "target": "未錄取或未報到者"}
]

# 1. 產生 CSV 檔案
def generate_csv():
    csv_path = "115學年度_中投區實用技能學程_招生學校與科班明細表.csv"
    with open(csv_path, mode="w", encoding="utf-8-sig", newline="") as f:
        writer = csv.writer(f)
        writer.writerow(["縣市", "公私立別", "學校全稱", "學校簡稱", "所屬行政區", "招生科班名稱", "對應職群", "授課時段", "預估核定名額(人)", "是否為水電技術科", "備註說明與科班特色"])
        for school in DATA_SCHOOLS:
            for p in school["programs"]:
                writer.writerow([
                    school["district"],
                    school["system"],
                    school["name"],
                    school["short_name"],
                    school["zone"],
                    p["dept"],
                    p["cluster"],
                    p["shift"],
                    p["quota"],
                    "★ 水電技術科" if p["is_plumbing"] else "一般技術科",
                    p["note"]
                ])
    print(f"[OK] CSV Generated: {csv_path}")

# 2. 產生 Markdown 深度彙編報告
def generate_markdown():
    md_path = "115學年度_中投區實用技能學程_招生簡章深度彙編表.md"
    total_schools = len(DATA_SCHOOLS)
    total_classes = sum(len(s["programs"]) for s in DATA_SCHOOLS)
    total_quota = sum(sum(p["quota"] for p in s["programs"]) for s in DATA_SCHOOLS)

    lines = []
    lines.append("# 115學年度 中投區高級中等學校實用技能學程（實技班）招生簡章深度彙編表\n")
    lines.append("> **主管機關**：教育部國民及學前教育署  ")
    lines.append("> **承辦學校**：臺中市立臺中家事商業高級中等學校（臺中家商）  ")
    lines.append(f"> **統計概況**：中投區共 **{total_schools}** 所高中職辦理、開設 **{total_classes}** 個科班、總核定招生容量約 **{total_quota}** 名。\n")
    lines.append("---\n")

    lines.append("## 壹、 報名資格、日程、錄取方式與成績採計總則\n")
    lines.append("### 一、 核心報名資格")
    lines.append("1. **招生對象**：國民中學應屆、非應屆畢業生或具同等學力者。")
    lines.append("2. **零會考門檻**：**完全不採計國中教育會考成績**，亦不採計國中在校學業成績，為技能導向免試入學管道。")
    lines.append("3. **一人一區限制**：可自由跨考區報名（免遷戶籍），惟**全國每名學生僅限向一個分發區報名**，嚴禁重複報名。")
    lines.append("4. **志願選填上限**：每位考生最多可選填 **7 個志願**（每校每科計為一個志願）。\n")

    lines.append("### 二、 115學年度重要升學日程表")
    lines.append("| 重要時程 | 作業項目與內容 | 適用對象與應備注意事項 |")
    lines.append("| :--- | :--- | :--- |")
    for t in TIMELINE_ITEMS:
        lines.append(f"| **{t['date']}** | **{t['stage']}**<br>{t['desc']} | {t['target']} |")
    lines.append("\n")

    lines.append("### 三、 錄取分發優先順序（四梯次階梯原則）")
    lines.append("當同一科班申請人數超過招生名額時，依下列先後順位逐梯次分發：")
    lines.append("1. **第 1 優先**：曾選習國中「技藝教育課程」之**應屆**畢（結）業生。")
    lines.append("2. **第 2 優先**：曾選習國中「技藝教育課程」之**非應屆**畢（結）業生。")
    lines.append("3. **第 3 優先**：未曾選習國中「技藝教育課程」之**應屆**畢（結）業生。")
    lines.append("4. **第 4 優先**：未曾選習國中「技藝教育課程」之**非應屆**畢（結）業生或具同等學力者。\n")

    lines.append("### 四、 超額比序與成績採計計算方式")
    lines.append("若在同一優先順序內報名人數仍大於招生名額，依報名表件分流比序：")
    lines.append("#### 1. 曾選習技藝教育學生（報名 A 表）比序規則：")
    lines.append("- **第 1 比序**：曾參加各直轄市、縣（市）政府主辦之技藝競賽或成果展獲獎，且具**低收入戶**證明者優先錄取。")
    lines.append("- **第 2 比序**：技藝競賽獲獎者（按直轄市、縣市政府獲獎名次高低排序）。")
    lines.append("- **第 3 比序**：弱勢身分優先（低收入戶優於中低收入戶）。")
    lines.append("- **第 4 比序**：**職群綜合成績積分**（依國中技藝教育修習證明書所載職群成績轉化 T 分數、修習點數、特殊表現等加權加總）。")
    lines.append("- **第 5 比序**：志願序順位（第 1 志願優於第 2 志願，以此類推）。\n")
    lines.append("#### 2. 未選習技藝教育學生（報名 B 表）比序規則：")
    lines.append("- **第 1 比序**：弱勢身分優先（低收入戶優於中低收入戶）。")
    lines.append("- **第 2 比序**：**綜合表現總分**（國中綜合活動領域前五學期平均成績 + 特殊表現積分）。")
    lines.append("- **第 3 比序**：志願序順位。\n")
    lines.append("#### 3. 外加名額保障")
    lines.append("- **原住民學生**：核定名額外加 2%（採無條件進位計算）。")
    lines.append("- **身心障礙學生**：核定名額外加 2%（採無條件進位計算，並保留十二年就學安置名額）。\n")
    lines.append("---\n")

    lines.append("## 貳、 特設專題：中投區開設【水電技術科】學校全面診斷\n")
    lines.append("水電技術人才在台灣工程市場處於供不應求狀態，實用技能學程「水電技術科」免學費、注重室內配線與水電實作，是學生學得一技之長、取得甲乙種電匠並銜接國立科大的絕佳途徑。\n")

    for p_school in PLUMBING_SCHOOLS_DETAILS:
        lines.append(f"### ⚡ 【{p_school['short']}】{p_school['school']}")
        lines.append(f"- **地址位置**：{p_school['district']}")
        lines.append(f"- **授課型態與名額**：**{p_school['shift']}** / **{p_school['quota']}**")
        lines.append(f"- **專業工場與設備**：{p_school['equipment']}")
        lines.append(f"- **核心課程規劃**：{p_school['curriculum']}")
        lines.append(f"- **主力考照目標**：{p_school['certifications']}")
        lines.append(f"- **產學合作與出路**：{p_school['partnerships']}")
        lines.append(f"- **專屬升學與就業優勢**：{p_school['advantages']}\n")

    lines.append("---\n")
    lines.append("## 參、 115學年度 中投區實用技能學程 招生學校完整科班清單（依學校區分）\n")

    lines.append("| 縣市 | 屬性 | 學校名稱 | 地區 | 開設科別 | 對應職群 | 上課時段 | 名額 | 水電標記 | 特色說明 |")
    lines.append("| :--- | :--- | :--- | :--- | :--- | :--- | :--- | :--- | :---: | :--- |")
    for s in DATA_SCHOOLS:
        for p in s["programs"]:
            plumb_tag = "⚡ **水電技術科**" if p["is_plumbing"] else "-"
            lines.append(f"| {s['district']} | {s['system']} | **{s['short_name']}**<br><small>{s['name']}</small> | {s['zone']} | **{p['dept']}** | {p['cluster']} | {p['shift']} | {p['quota']} | {plumb_tag} | {p['note']} |")

    lines.append("\n---\n")
    lines.append("## 肆、 國三輔導教師與家長志願選填實戰錦囊\n")
    lines.append("1. **與一般免試入學（會考）不衝突**：實用技能學程於 6 月 11 日放榜、6 月 12 日報到。若分發結果滿意，可直接報到入學；若想改衝中投區免試入學（7月放榜），只要在 6 月 15 日中午 12:00 前辦理聲明放棄即可，多一次完全免費的高中職升學保底機會！")
    lines.append("2. **水電科選填戰略**：")
    lines.append("   - 若希望正常作息、考取雙乙級證照並銜接國立科技大學，首選**青年高中水電技術科（日間部）**。")
    lines.append("   - 若位於大甲、清水、沙鹿等海線地區，首選**致用高中水電技術科（日間部）**，在地扎根省去舟車勞頓。")
    lines.append("   - 若家庭經濟需支援、希望白天全職在水電行當學徒實習賺工資，首選**東勢高工水電技術科（夜間部）**，學費全免且公立名校光環。")
    lines.append("3. **技藝教育修習者的絕對主場**：第 1 優先梯次完全保留給有修習國中技藝教育課程的應屆畢業生，只要志願序排填得當，錄取公立或熱門私立名校實技班的勝率高達 95% 以上。")

    with open(md_path, mode="w", encoding="utf-8") as f:
        f.write("\n".join(lines))
    print(f"[OK] Markdown Generated: {md_path}")

# 3. 產生互動式 DocuSheet 網頁 (index.html)
def generate_html():
    html_path = "index.html"
    total_schools = len(DATA_SCHOOLS)
    total_classes = sum(len(s["programs"]) for s in DATA_SCHOOLS)
    total_quota = sum(sum(p["quota"] for p in s["programs"]) for s in DATA_SCHOOLS)
    plumbing_schools = [s for s in DATA_SCHOOLS if any(p["is_plumbing"] for p in s["programs"])]

    flat_data = []
    idx = 1
    for s in DATA_SCHOOLS:
        for p in s["programs"]:
            flat_data.append({
                "id": idx,
                "district": s["district"],
                "system": s["system"],
                "name": s["name"],
                "short_name": s["short_name"],
                "zone": s["zone"],
                "dept": p["dept"],
                "cluster": p["cluster"],
                "shift": p["shift"],
                "quota": p["quota"],
                "is_plumbing": p["is_plumbing"],
                "note": p["note"],
                "is_organizer": s.get("is_organizer", False)
            })
            idx += 1

    json_str = json.dumps(flat_data, ensure_ascii=False)
    plumb_json = json.dumps(PLUMBING_SCHOOLS_DETAILS, ensure_ascii=False)
    timeline_json = json.dumps(TIMELINE_ITEMS, ensure_ascii=False)

    html_content = f"""<!DOCTYPE html>
<html lang="zh-TW">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>115學年度 中投區實用技能學程 招生簡章互動式 DocuSheet</title>
  <style>
    :root {{
      --primary: #1e3a8a;
      --primary-dark: #0f172a;
      --primary-light: #3b82f6;
      --accent: #f59e0b;
      --accent-light: #fef3c7;
      --success: #10b981;
      --purple: #8b5cf6;
      --bg: #f8fafc;
      --card-bg: #ffffff;
      --text: #1e293b;
      --text-muted: #64748b;
      --border: #e2e8f0;
    }}
    * {{
      box-sizing: border-box;
      margin: 0;
      padding: 0;
    }}
    body {{
      font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, "Helvetica Neue", Arial, "Noto Sans TC", sans-serif;
      background-color: var(--bg);
      color: var(--text);
      line-height: 1.6;
    }}
    .header {{
      background: linear-gradient(135deg, #0f172a 0%, #1e3a8a 50%, #1d4ed8 100%);
      color: white;
      padding: 2.2rem 2rem 2rem 2rem;
      position: relative;
      box-shadow: 0 4px 20px rgba(0,0,0,0.12);
    }}
    .header-top {{
      display: flex;
      justify-content: space-between;
      align-items: flex-start;
      margin-bottom: 0.8rem;
    }}
    .version-badge {{
      background: rgba(255, 255, 255, 0.15);
      backdrop-filter: blur(8px);
      padding: 0.35rem 0.85rem;
      border-radius: 9999px;
      font-size: 0.85rem;
      font-weight: 600;
      letter-spacing: 0.03em;
      border: 1px solid rgba(255, 255, 255, 0.25);
      color: #e2e8f0;
      white-space: nowrap;
    }}
    .header h1 {{
      font-size: 2.1rem;
      font-weight: 800;
      letter-spacing: -0.02em;
      margin-bottom: 0.5rem;
    }}
    .header p {{
      color: #cbd5e1;
      font-size: 1.05rem;
      max-width: 900px;
    }}
    .kpi-row {{
      display: grid;
      grid-template-columns: repeat(auto-fit, minmax(200px, 1fr));
      gap: 1rem;
      margin-top: 1.5rem;
    }}
    .kpi-card {{
      background: rgba(255, 255, 255, 0.1);
      backdrop-filter: blur(6px);
      border: 1px solid rgba(255, 255, 255, 0.15);
      padding: 1rem 1.25rem;
      border-radius: 12px;
      display: flex;
      flex-direction: column;
    }}
    .kpi-value {{
      font-size: 1.85rem;
      font-weight: 800;
      color: #fbbf24;
    }}
    .kpi-label {{
      font-size: 0.85rem;
      color: #e2e8f0;
    }}

    .container {{
      max-width: 1400px;
      margin: 0 auto;
      padding: 1.5rem;
    }}

    .nav-tabs {{
      display: flex;
      gap: 0.5rem;
      background: #e2e8f0;
      padding: 0.4rem;
      border-radius: 12px;
      margin-bottom: 1.5rem;
      overflow-x: auto;
    }}
    .tab-btn {{
      padding: 0.75rem 1.4rem;
      border: none;
      background: transparent;
      border-radius: 8px;
      font-size: 0.95rem;
      font-weight: 700;
      color: var(--text-muted);
      cursor: pointer;
      display: flex;
      align-items: center;
      gap: 0.5rem;
      transition: all 0.2s;
      white-space: nowrap;
    }}
    .tab-btn:hover {{
      color: var(--primary);
      background: rgba(255, 255, 255, 0.6);
    }}
    .tab-btn.active {{
      background: #ffffff;
      color: var(--primary);
      box-shadow: 0 2px 8px rgba(0,0,0,0.08);
    }}

    .tab-panel {{
      display: none;
    }}
    .tab-panel.active {{
      display: block;
      animation: fadeIn 0.3s ease-in-out;
    }}

    @keyframes fadeIn {{
      from {{ opacity: 0; transform: translateY(6px); }}
      to {{ opacity: 1; transform: translateY(0); }}
    }}

    /* Controls */
    .controls-card {{
      background: #ffffff;
      padding: 1.25rem;
      border-radius: 12px;
      border: 1px solid var(--border);
      box-shadow: 0 2px 10px rgba(0,0,0,0.03);
      margin-bottom: 1.25rem;
      display: flex;
      flex-wrap: wrap;
      gap: 1rem;
      align-items: center;
      justify-content: space-between;
    }}
    .search-box {{
      flex: 1 1 320px;
      position: relative;
    }}
    .search-box input {{
      width: 100%;
      padding: 0.7rem 1rem 0.7rem 2.6rem;
      border: 1px solid var(--border);
      border-radius: 8px;
      font-size: 0.95rem;
      outline: none;
      transition: border-color 0.2s;
    }}
    .search-box input:focus {{
      border-color: var(--primary-light);
      box-shadow: 0 0 0 3px rgba(59, 130, 246, 0.15);
    }}
    .search-icon {{
      position: absolute;
      left: 0.85rem;
      top: 50%;
      transform: translateY(-50%);
      color: var(--text-muted);
    }}
    .filter-pills {{
      display: flex;
      flex-wrap: wrap;
      gap: 0.4rem;
    }}
    .pill-btn {{
      padding: 0.45rem 0.9rem;
      border-radius: 9999px;
      border: 1px solid var(--border);
      background: #f8fafc;
      font-size: 0.85rem;
      font-weight: 600;
      color: var(--text);
      cursor: pointer;
      transition: all 0.2s;
    }}
    .pill-btn:hover {{
      background: #e2e8f0;
    }}
    .pill-btn.active {{
      background: var(--primary);
      color: white;
      border-color: var(--primary);
    }}
    .pill-btn.highlight {{
      border-color: #f59e0b;
      color: #b45309;
      background: #fef3c7;
    }}
    .pill-btn.highlight.active {{
      background: #f59e0b;
      color: white;
      border-color: #f59e0b;
    }}

    /* Table */
    .table-responsive {{
      background: #ffffff;
      border-radius: 12px;
      border: 1px solid var(--border);
      box-shadow: 0 4px 14px rgba(0,0,0,0.03);
      overflow-x: auto;
    }}
    table {{
      width: 100%;
      border-collapse: collapse;
      text-align: left;
      font-size: 0.92rem;
    }}
    th {{
      background: #f1f5f9;
      color: #334155;
      font-weight: 700;
      padding: 0.85rem 1rem;
      border-bottom: 2px solid var(--border);
      white-space: nowrap;
    }}
    td {{
      padding: 0.85rem 1rem;
      border-bottom: 1px solid var(--border);
      vertical-align: middle;
    }}
    tr:hover {{
      background: #f8fafc;
    }}
    .badge {{
      display: inline-block;
      padding: 0.2rem 0.55rem;
      border-radius: 6px;
      font-size: 0.75rem;
      font-weight: 700;
    }}
    .badge-public {{
      background: #ecfdf5;
      color: #047857;
      border: 1px solid #a7f3d0;
    }}
    .badge-private {{
      background: #f5f3ff;
      color: #6d28d9;
      border: 1px solid #ddd6fe;
    }}
    .badge-day {{
      background: #e0f2fe;
      color: #0369a1;
      border: 1px solid #bae6fd;
    }}
    .badge-night {{
      background: #f1f5f9;
      color: #475569;
      border: 1px solid #cbd5e1;
    }}
    .badge-plumbing {{
      background: #fef3c7;
      color: #b45309;
      border: 1px solid #fde68a;
      font-weight: 800;
    }}

    /* Cards */
    .plumbing-grid {{
      display: grid;
      grid-template-columns: repeat(auto-fit, minmax(380px, 1fr));
      gap: 1.5rem;
      margin-top: 1rem;
    }}
    .plumbing-card {{
      background: #ffffff;
      border-radius: 14px;
      border: 1px solid var(--border);
      box-shadow: 0 4px 18px rgba(0,0,0,0.05);
      padding: 1.5rem;
      position: relative;
      overflow: hidden;
      display: flex;
      flex-direction: column;
    }}
    .plumbing-card::before {{
      content: "";
      position: absolute;
      top: 0;
      left: 0;
      right: 0;
      height: 6px;
      background: linear-gradient(90deg, #f59e0b, #ef4444);
    }}
    .plumbing-card-header {{
      display: flex;
      justify-content: space-between;
      align-items: flex-start;
      margin-bottom: 1rem;
    }}
    .plumbing-card-title {{
      font-size: 1.35rem;
      font-weight: 800;
      color: #0f172a;
    }}
    .plumbing-detail-item {{
      margin-bottom: 0.85rem;
      font-size: 0.92rem;
    }}
    .plumbing-detail-label {{
      font-weight: 700;
      color: #334155;
      display: block;
      margin-bottom: 0.2rem;
    }}
    .plumbing-detail-val {{
      color: #475569;
    }}

    /* Timeline & Grid */
    .timeline-container {{
      background: #ffffff;
      padding: 1.5rem;
      border-radius: 12px;
      border: 1px solid var(--border);
      box-shadow: 0 4px 14px rgba(0,0,0,0.03);
    }}
    .timeline-item {{
      display: grid;
      grid-template-columns: 220px 1fr 180px;
      gap: 1.25rem;
      padding: 1rem 0;
      border-bottom: 1px solid var(--border);
      align-items: center;
    }}
    .timeline-date {{
      font-weight: 800;
      color: var(--primary);
      font-size: 1rem;
    }}
    .timeline-title {{
      font-weight: 700;
      color: #0f172a;
      margin-bottom: 0.25rem;
    }}
    .timeline-desc {{
      color: var(--text-muted);
      font-size: 0.9rem;
    }}

    /* Hierarchy block */
    .hierarchy-box {{
      background: #ffffff;
      border-radius: 12px;
      padding: 1.5rem;
      border: 1px solid var(--border);
      margin-bottom: 1.5rem;
    }}
    .step-list {{
      list-style: none;
      counter-reset: rank-counter;
      margin-top: 1rem;
    }}
    .step-item {{
      counter-increment: rank-counter;
      position: relative;
      padding-left: 3rem;
      margin-bottom: 1.25rem;
    }}
    .step-item::before {{
      content: counter(rank-counter);
      position: absolute;
      left: 0;
      top: 0;
      width: 2.2rem;
      height: 2.2rem;
      background: var(--primary);
      color: white;
      border-radius: 50%;
      display: flex;
      align-items: center;
      justify-content: center;
      font-weight: 800;
      font-size: 1.05rem;
    }}
    .step-title {{
      font-size: 1.05rem;
      font-weight: 700;
      color: #0f172a;
      margin-bottom: 0.25rem;
    }}

    /* Print & Export */
    .export-bar {{
      display: flex;
      gap: 0.75rem;
      margin-bottom: 1rem;
      justify-content: flex-end;
    }}
    .btn-action {{
      padding: 0.5rem 1rem;
      border-radius: 8px;
      border: 1px solid var(--border);
      background: #ffffff;
      font-weight: 600;
      font-size: 0.88rem;
      cursor: pointer;
      display: flex;
      align-items: center;
      gap: 0.4rem;
      transition: all 0.2s;
    }}
    .btn-action:hover {{
      background: #f1f5f9;
      border-color: #cbd5e1;
    }}

    @media print {{
      .header {{ padding: 1rem; }}
      .nav-tabs, .controls-card, .export-bar {{ display: none !important; }}
      .tab-panel {{ display: block !important; page-break-after: always; }}
      body {{ background: white; }}
      .table-responsive {{ border: none; box-shadow: none; }}
      table {{ font-size: 8pt; }}
      th, td {{ padding: 4px 6px; }}
    }}
  </style>
</head>
<body>

  <!-- Top Header with Version & Timestamp -->
  <header class="header">
    <div class="header-top">
      <div style="font-size: 0.9rem; font-weight: 700; color: #93c5fd; letter-spacing: 0.05em; text-transform: uppercase;">
        教育部國民及學前教育署 115學年度核定標準
      </div>
      <div class="version-badge" id="versionBadge">
        v1.0 | 2026-11-13 10:25
      </div>
    </div>
    <h1>115學年度 中投區高級中等學校實用技能學程 招生簡章 DocuSheet</h1>
    <p>統整中投區（臺中市、南投縣）實用技能學程分發招生簡章、各校招生科系與容量名額、水電技術科專題評析、免會考分發順位、採計比序及選填關鍵指南。</p>

    <div class="kpi-row">
      <div class="kpi-card">
        <span class="kpi-value">{total_schools} 所</span>
        <span class="kpi-label">中投區主辦招生高中職</span>
      </div>
      <div class="kpi-card">
        <span class="kpi-value">{total_classes} 個</span>
        <span class="kpi-label">實用技能學程專業科班</span>
      </div>
      <div class="kpi-card">
        <span class="kpi-value">約 {total_quota} 名</span>
        <span class="kpi-label">預計招生總名額（完全免會考）</span>
      </div>
      <div class="kpi-card">
        <span class="kpi-value">{len(plumbing_schools)} 所（4班）</span>
        <span class="kpi-label">設有【水電技術科】專屬學校</span>
      </div>
    </div>
  </header>

  <main class="container">
    <!-- Navigation Tabs -->
    <div class="nav-tabs">
      <button class="tab-btn active" onclick="switchTab('tab-sheet')">📋 招生學校與科班明細表</button>
      <button class="tab-btn" onclick="switchTab('tab-plumbing')">⚡ 水電技術科專區 (熱門首選)</button>
      <button class="tab-btn" onclick="switchTab('tab-timeline')">📅 報名資格與重要日程</button>
      <button class="tab-btn" onclick="switchTab('tab-rules')">⚖️ 錄取方式與超額比序採計</button>
      <button class="tab-btn" onclick="switchTab('tab-guide')">💡 輔導選填實戰錦囊</button>
    </div>

    <!-- TAB 1: 招生學校與科班明細表 -->
    <section id="tab-sheet" class="tab-panel active">
      <div class="export-bar">
        <button class="btn-action" onclick="downloadCSV()">📥 匯出 Excel (CSV)</button>
        <button class="btn-action" onclick="window.print()">🖨️ 列印 / 另存 PDF</button>
      </div>

      <div class="controls-card">
        <div class="search-box">
          <span class="search-icon">🔍</span>
          <input type="text" id="searchInput" placeholder="搜尋學校名稱、科別（如水電、烘焙、汽車）、群別或地區..." onkeyup="filterData()">
        </div>
        <div class="filter-pills">
          <button class="pill-btn active" onclick="setFilter('all', this)">全部 ({len(flat_data)})</button>
          <button class="pill-btn highlight" onclick="setFilter('plumbing', this)">⚡ 僅看水電科 (4)</button>
          <button class="pill-btn" onclick="setFilter('taichung', this)">臺中市 (14校)</button>
          <button class="pill-btn" onclick="setFilter('nantou', this)">南投縣 (4校)</button>
          <button class="pill-btn" onclick="setFilter('public', this)">公立高職</button>
          <button class="pill-btn" onclick="setFilter('private', this)">私立高中職</button>
          <button class="pill-btn" onclick="setFilter('day', this)">日間上課</button>
          <button class="pill-btn" onclick="setFilter('night', this)">夜間上課</button>
        </div>
      </div>

      <div class="table-responsive">
        <table id="dataTable">
          <thead>
            <tr>
              <th>#</th>
              <th>縣市地區</th>
              <th>公私立別</th>
              <th>學校名稱</th>
              <th>科班名稱</th>
              <th>對應職群</th>
              <th>授課時段</th>
              <th>預估名額</th>
              <th>科班特色與實務方向</th>
            </tr>
          </thead>
          <tbody id="tableBody">
            <!-- Rendered by JS -->
          </tbody>
        </table>
      </div>
      <div id="noResult" style="display:none; text-align:center; padding: 2.5rem; color: var(--text-muted);">
        查無符合條件的學校或科班，請嘗試調整搜尋關鍵字或篩選標籤。
      </div>
    </section>

    <!-- TAB 2: 水電技術科專區 -->
    <section id="tab-plumbing" class="tab-panel">
      <div style="background: #fffbeb; border: 1px solid #fef3c7; border-left: 5px solid #f59e0b; padding: 1.25rem; border-radius: 8px; margin-bottom: 1.5rem;">
        <h3 style="color: #92400e; font-size: 1.2rem; margin-bottom: 0.4rem;">⚡ 為什麼【水電技術科】是技職升學與就業的熱門黑馬？</h3>
        <p style="color: #78350f; font-size: 0.95rem;">
          台灣目前面臨住宅水電裝修、高科技廠房機電配線、太陽能光電及綠能基礎建設嚴重缺工。實用技能學程水電技術科三年學雜費全免，專注輔導<strong>室內配線（甲/乙種電匠）</strong>、<strong>自來水管配管</strong>等法定硬核證照。畢業後起薪高，更可透過「技優甄審」管道保送國立科技大學！
        </p>
      </div>

      <div class="plumbing-grid">
        <!-- Rendered by JS from PLUMBING_SCHOOLS_DETAILS -->
      </div>
    </section>

    <!-- TAB 3: 報名資格與重要日程 -->
    <section id="tab-timeline" class="tab-panel">
      <div style="background: #ffffff; padding: 1.5rem; border-radius: 12px; border: 1px solid var(--border); margin-bottom: 1.5rem;">
        <h3 style="color: var(--primary); font-size: 1.25rem; margin-bottom: 0.8rem;">📌 115學年度 報名核心資格規約</h3>
        <ul style="padding-left: 1.25rem; color: #334155; line-height: 1.8;">
          <li><strong>招生對象</strong>：全國公私立國民中學應屆畢業生、非應屆畢業生或具同等學力者。</li>
          <li><strong>免試免會考</strong>：<strong>完全不看國中教育會考成績</strong>，亦不採計國中在校學業成績，旨在打破紙筆測驗限制，保障重實作、性向明確的學生。</li>
          <li><strong>分發區限制</strong>：免遷戶口可跨區選填，但<strong>每位學生全國僅限報名一個分發區</strong>，不得重複報名其他考區。</li>
          <li><strong>志願選填數</strong>：每位考生最多選填 <strong>7 個志願</strong>（每一學校之一科班計為一個志願）。</li>
          <li><strong>報名表別分流</strong>：
            <ul style="margin-top: 0.4rem; padding-left: 1rem;">
              <li><strong>報名 A 表</strong>：曾選習國中「技藝教育課程」之學生（具修習證明書，享有最優先梯次分發權利）。</li>
              <li><strong>報名 B 表</strong>：未曾選習國中「技藝教育課程」之一般學生（依前五學期綜合活動成績及特殊表現評比）。</li>
            </ul>
          </li>
        </ul>
      </div>

      <div class="timeline-container">
        <h3 style="color: #0f172a; margin-bottom: 1.25rem;">📅 115學年度 中投區實用技能學程 關鍵升學日程表</h3>
        <div id="timelineList">
          <!-- Rendered by JS -->
        </div>
      </div>
    </section>

    <!-- TAB 4: 錄取方式與分發採計 -->
    <section id="tab-rules" class="tab-panel">
      <div class="hierarchy-box">
        <h3 style="color: var(--primary); font-size: 1.25rem; margin-bottom: 0.5rem;">⚖️ 錄取分發梯次架構（四梯次階梯原則）</h3>
        <p style="color: var(--text-muted); font-size: 0.95rem;">當同一學校科班報名人數超過核定招生名額時，統一依據下列資格進行梯次依序分發：</p>
        
        <ul class="step-list">
          <li class="step-item">
            <div class="step-title">第一優先順位：曾選習國中技藝教育課程之「應屆」畢（結）業生</div>
            <div style="color: #64748b; font-size: 0.9rem;">享有最高分發保障！名額首先分配給國三曾參加國中技藝專案抽離式或專班之應屆同學。</div>
          </li>
          <li class="step-item">
            <div class="step-title">第二優先順位：曾選習國中技藝教育課程之「非應屆」畢（結）業生</div>
            <div style="color: #64748b; font-size: 0.9rem;">前一學年度以前畢業但曾修畢技藝教育課程之重考生。</div>
          </li>
          <li class="step-item">
            <div class="step-title">第三優先順位：未曾選習國中技藝教育課程之「應屆」畢（結）業生</div>
            <div style="color: #64748b; font-size: 0.9rem;">一般國三應屆畢業生，依 B 表填報成績進行遞補分發。</div>
          </li>
          <li class="step-item">
            <div class="step-title">第四優先順位：未曾選習國中技藝教育課程之「非應屆」畢（結）業生或同等學力者</div>
            <div style="color: #64748b; font-size: 0.9rem;">非應屆一般生或以同等學力報名者。</div>
          </li>
        </ul>
      </div>

      <div style="display: grid; grid-template-columns: repeat(auto-fit, minmax(400px, 1fr)); gap: 1.5rem;">
        <div style="background: white; padding: 1.5rem; border-radius: 12px; border: 1px solid var(--border);">
          <h4 style="color: #0f172a; font-size: 1.15rem; margin-bottom: 0.75rem;">📋 曾選習技藝教育（A 表）同梯超額比序順序</h4>
          <ol style="padding-left: 1.25rem; color: #475569; line-height: 1.8;">
            <li><strong>競賽且弱勢</strong>：曾獲直轄市、縣（市）政府主辦之技藝競賽或成果展得獎，且具<strong>低收入戶</strong>身分者最優先錄取。</li>
            <li><strong>技藝競賽獲獎</strong>：依直轄市、縣市公辦技藝競賽獲獎名次依序錄取。</li>
            <li><strong>經濟弱勢身分</strong>：低收入戶證明優先於中低收入戶證明。</li>
            <li><strong>職群綜合成績積分</strong>：依技藝教育修習證明書職群轉換之 T 分數、修習時數點數加總。</li>
            <li><strong>志願序順序</strong>：第 1 志願優於第 2 志願。</li>
          </ol>
        </div>

        <div style="background: white; padding: 1.5rem; border-radius: 12px; border: 1px solid var(--border);">
          <h4 style="color: #0f172a; font-size: 1.15rem; margin-bottom: 0.75rem;">📋 未選習技藝教育（B 表）同梯超額比序順序</h4>
          <ol style="padding-left: 1.25rem; color: #475569; line-height: 1.8;">
            <li><strong>經濟弱勢身分</strong>：低收入戶證明優先於中低收入戶證明。</li>
            <li><strong>綜合表現總分</strong>：以國中綜合活動領域前五學期平均成績 + 特殊表現總分高低依序分發。</li>
            <li><strong>志願序順序</strong>：第 1 志願優於第 2 志願。</li>
          </ol>
          <div style="margin-top: 1rem; padding: 0.85rem; background: #f8fafc; border-radius: 8px; font-size: 0.88rem; color: #64748b;">
            <strong>外加名額保障</strong>：原住民身分學生外加 2%、身心障礙學生外加 2%（小數點皆採無條件進位），不佔一般生原核定名額。
          </div>
        </div>
      </div>
    </section>

    <!-- TAB 5: 輔導選填實戰錦囊 -->
    <section id="tab-guide" class="tab-panel">
      <div style="display: grid; grid-template-columns: repeat(auto-fit, minmax(360px, 1fr)); gap: 1.5rem;">
        <div style="background: white; padding: 1.5rem; border-radius: 12px; border: 1px solid var(--border);">
          <h3 style="color: var(--primary); font-size: 1.2rem; margin-bottom: 0.75rem;">💡 策略一：實技班與大免試（會考）是雙軌並行！</h3>
          <p style="color: #475569; font-size: 0.95rem; line-height: 1.8;">
            許多家長誤以為報名實用技能學程就不能參加會考或一般免試入學。<strong>事實上兩者完全不衝突！</strong><br>
            實技班在 <strong>6月11日放榜、6月12日報到</strong>。若錄取了心目中理想的高工或水電科，可直接報到安心度過國三暑假；若還想等待大免試放榜，只要在 <strong>6月15日中午 12:00 前</strong> 填妥放棄聲明書，即可無縫參加中投區高級中等學校免試入學，等於多賺到一次極佳的優先升學保底機會。
          </p>
        </div>

        <div style="background: white; padding: 1.5rem; border-radius: 12px; border: 1px solid var(--border);">
          <h3 style="color: var(--primary); font-size: 1.2rem; margin-bottom: 0.75rem;">💡 策略二：水電技術科三大名校選填決策矩陣</h3>
          <ul style="padding-left: 1.2rem; color: #475569; font-size: 0.93rem; line-height: 1.8;">
            <li><strong>青年高中（日間部）</strong>：適合希望正常白天上課、目標考取室內配線雙乙級證照、未來打算保送國立勤益、雲科、虎科大之學生。</li>
            <li><strong>致用高中（日間部）</strong>：海線在地同學（大甲、清水、沙鹿、大安、苑裡）首選，在地企業實習保證就業，節省通學時間。</li>
            <li><strong>東勢高工（夜間進修部）</strong>：適合家庭經濟需要支援、希望白天在水電行全職當學徒賺取工資、晚上就讀公立高工免學費的自立自強學生。</li>
          </ul>
        </div>

        <div style="background: white; padding: 1.5rem; border-radius: 12px; border: 1px solid var(--border);">
          <h3 style="color: var(--primary); font-size: 1.2rem; margin-bottom: 0.75rem;">💡 策略三：志願序排填「黃金三梯次」原則</h3>
          <p style="color: #475569; font-size: 0.95rem; line-height: 1.8;">
            每人可填 7 個志願，切記不要全部填同一職群或同一個熱門公立學校：<br>
            <strong>志願 1~2（夢幻衝刺）</strong>：公立日間或夜間熱門科班（如東勢水電、沙鹿商用資訊、興大附農烘焙/餐飲）。<br>
            <strong>志願 3~5（務實精準）</strong>：師資設備強大、證照錄取率極高之私立日間旗艦科班（如青年水電、致用水電、僑泰汽修/寵物）。<br>
            <strong>志願 6~7（堅實保底）</strong>：交通便利、招生名額充裕之區域高職，確保 100% 順利分發上榜。
          </p>
        </div>
      </div>
    </section>
  </main>

  <footer style="text-align: center; padding: 2rem; color: var(--text-muted); font-size: 0.85rem; border-top: 1px solid var(--border); margin-top: 2rem; background: white;">
    115學年度 中投區高級中等學校實用技能學程分發委員會（承辦學校：臺中市立臺中家事商業高級中等學校）<br>
    本互動式 DocuSheet 依教育部核定招生簡章彙編發佈 ｜ 格式支援電腦、平板、手機與直接列印
  </footer>

  <script>
    const schoolsData = {json_str};
    const plumbingData = {plumb_json};
    const timelineData = {timeline_json};

    let currentFilter = 'all';

    function init() {{
      renderTable(schoolsData);
      renderPlumbingCards();
      renderTimeline();
    }}

    function renderTable(data) {{
      const tbody = document.getElementById('tableBody');
      const noResult = document.getElementById('noResult');
      tbody.innerHTML = '';

      if (data.length === 0) {{
        noResult.style.display = 'block';
        return;
      }}
      noResult.style.display = 'none';

      data.forEach(item => {{
        const tr = document.createElement('tr');
        const sysClass = item.system.includes('公立') ? 'badge-public' : 'badge-private';
        const shiftClass = item.shift.includes('日間') ? 'badge-day' : 'badge-night';
        const plumbBadge = item.is_plumbing ? '<span class="badge badge-plumbing">⚡ 水電技術科</span>' : '<span style="color:#94a3b8;">-</span>';

        tr.innerHTML = `
          <td><strong>${{item.id}}</strong></td>
          <td>${{item.district}} <small style="color:#64748b;">(${{item.zone}})</small></td>
          <td><span class="badge ${{sysClass}}">${{item.system}}</span></td>
          <td>
            <strong>${{item.short_name}}</strong>
            ${{item.is_organizer ? '<span class="badge" style="background:#fee2e2; color:#b91c1c; font-size:0.7rem; margin-left:4px;">主委學校</span>' : ''}}
            <br><small style="color:#64748b;">${{item.name}}</small>
          </td>
          <td>
            <strong style="color: ${{item.is_plumbing ? '#d97706' : '#1e293b'}};">${{item.dept}}</strong>
          </td>
          <td><span style="background:#f1f5f9; padding:2px 6px; border-radius:4px; font-size:0.85rem;">${{item.cluster}}</span></td>
          <td><span class="badge ${{shiftClass}}">${{item.shift}}</span></td>
          <td style="font-weight:700;">約 ${{item.quota}} 名</td>
          <td style="font-size:0.85rem; color:#475569; max-width:320px;">
            ${{item.is_plumbing ? plumbBadge + ' ' : ''}}${{item.note}}
          </td>
        `;
        tbody.appendChild(tr);
      }});
    }}

    function renderPlumbingCards() {{
      const container = document.querySelector('.plumbing-grid');
      container.innerHTML = '';
      plumbingData.forEach(p => {{
        const card = document.createElement('div');
        card.className = 'plumbing-card';
        card.innerHTML = `
          <div class="plumbing-card-header">
            <div>
              <div class="plumbing-card-title">${{p.short}}</div>
              <small style="color: #64748b;">${{p.school}}</small>
            </div>
            <span class="badge badge-plumbing" style="font-size:0.85rem; padding: 0.3rem 0.6rem;">${{p.shift}}</span>
          </div>

          <div class="plumbing-detail-item">
            <span class="plumbing-detail-label">📍 地點位置與招生名額</span>
            <span class="plumbing-detail-val">${{p.district}} ｜ <strong>${{p.quota}}</strong></span>
          </div>

          <div class="plumbing-detail-item">
            <span class="plumbing-detail-label">🛠️ 專業檢定工場與硬體設備</span>
            <span class="plumbing-detail-val">${{p.equipment}}</span>
          </div>

          <div class="plumbing-detail-item">
            <span class="plumbing-detail-label">📚 核心技能課程規劃</span>
            <span class="plumbing-detail-val">${{p.curriculum}}</span>
          </div>

          <div class="plumbing-detail-item">
            <span class="plumbing-detail-label">🏆 輔導取得專業證照</span>
            <span class="plumbing-detail-val" style="color:#b45309; font-weight:700;">${{p.certifications}}</span>
          </div>

          <div class="plumbing-detail-item">
            <span class="plumbing-detail-label">🤝 產學合作與實習就業</span>
            <span class="plumbing-detail-val">${{p.partnerships}}</span>
          </div>

          <div class="plumbing-detail-item" style="margin-top:auto; padding-top:0.75rem; border-top:1px dashed #e2e8f0;">
            <span class="plumbing-detail-label" style="color:var(--primary);">🌟 獨家升學與就業優勢</span>
            <span class="plumbing-detail-val" style="font-weight:600; color:#334155;">${{p.advantages}}</span>
          </div>
        `;
        container.appendChild(card);
      }});
    }}

    function renderTimeline() {{
      const container = document.getElementById('timelineList');
      container.innerHTML = '';
      timelineData.forEach(t => {{
        const item = document.createElement('div');
        item.className = 'timeline-item';
        item.innerHTML = `
          <div class="timeline-date">${{t.date}}</div>
          <div>
            <div class="timeline-title">${{t.stage}}</div>
            <div class="timeline-desc">${{t.desc}}</div>
          </div>
          <div style="font-size:0.85rem; color:#64748b; text-align:right;">
            <span style="background:#f1f5f9; padding:4px 8px; border-radius:6px;">${{t.target}}</span>
          </div>
        `;
        container.appendChild(item);
      }});
    }}

    function switchTab(tabId) {{
      document.querySelectorAll('.tab-btn').forEach(btn => btn.classList.remove('active'));
      document.querySelectorAll('.tab-panel').forEach(panel => panel.classList.remove('active'));

      event.currentTarget.classList.add('active');
      document.getElementById(tabId).classList.add('active');
    }}

    function setFilter(type, btn) {{
      document.querySelectorAll('.pill-btn').forEach(b => b.classList.remove('active'));
      btn.classList.add('active');
      currentFilter = type;
      filterData();
    }}

    function filterData() {{
      const query = document.getElementById('searchInput').value.trim().toLowerCase();
      let filtered = schoolsData.filter(item => {{
        // Pill filter
        if (currentFilter === 'plumbing' && !item.is_plumbing) return false;
        if (currentFilter === 'taichung' && item.district !== '臺中市') return false;
        if (currentFilter === 'nantou' && item.district !== '南投縣') return false;
        if (currentFilter === 'public' && !item.system.includes('公立')) return false;
        if (currentFilter === 'private' && !item.system.includes('私立')) return false;
        if (currentFilter === 'day' && !item.shift.includes('日間')) return false;
        if (currentFilter === 'night' && !item.shift.includes('夜間')) return false;

        // Search text
        if (query) {{
          const matchStr = (item.name + ' ' + item.short_name + ' ' + item.dept + ' ' + item.cluster + ' ' + item.zone + ' ' + item.note).toLowerCase();
          return matchStr.includes(query);
        }}
        return true;
      }});

      renderTable(filtered);
    }}

    function downloadCSV() {{
      let csvContent = "\\uFEFF縣市,公私立別,學校全稱,學校簡稱,行政區,招生科班,對應職群,授課時段,預計名額,水電標記,特色說明\\n";
      schoolsData.forEach(item => {{
        csvContent += `"${{item.district}}","${{item.system}}","${{item.name}}","${{item.short_name}}","${{item.zone}}","${{item.dept}}","${{item.cluster}}","${{item.shift}}","${{item.quota}}","${{item.is_plumbing ? '水電技術科' : '一般'}}","${{item.note.replace(/"/g, '""')}}"\\n`;
      }});

      const blob = new Blob([csvContent], {{ type: 'text/csv;charset=utf-8;' }});
      const link = document.createElement("a");
      const url = URL.createObjectURL(blob);
      link.setAttribute("href", url);
      link.setAttribute("download", "115學年度_中投區實用技能學程_招生學校與科班明細表.csv");
      document.body.appendChild(link);
      link.click();
      document.body.removeChild(link);
    }}

    window.onload = init;
  </script>
</body>
</html>
"""
    with open(html_path, mode="w", encoding="utf-8") as f:
        f.write(html_content)
    print(f"[OK] HTML Generated: {html_path}")

def main():
    generate_csv()
    generate_markdown()
    generate_html()

if __name__ == "__main__":
    main()
