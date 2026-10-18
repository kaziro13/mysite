#!/usr/bin/env python3
"""日本の化学企業 規模順リスト Excel生成スクリプト"""

from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill, NamedStyle
from openpyxl.utils import get_column_letter
from openpyxl.formatting.rule import FormulaRule
from openpyxl.worksheet.datavalidation import DataValidation

wb = Workbook()

# Styles
thin = Border(
    left=Side(style="thin", color="B0B0B0"),
    right=Side(style="thin", color="B0B0B0"),
    top=Side(style="thin", color="B0B0B0"),
    bottom=Side(style="thin", color="B0B0B0"),
)
header_fill = PatternFill("solid", fgColor="1F4E79")
header_font = Font(name="Yu Gothic", bold=True, color="FFFFFF", size=11)
title_font = Font(name="Yu Gothic", bold=True, size=16, color="1F4E79")
subtitle_font = Font(name="Yu Gothic", bold=True, size=12, color="2E75B6")
body_font = Font(name="Yu Gothic", size=10)
wrap = Alignment(wrap_text=True, vertical="top")
center = Alignment(wrap_text=True, vertical="center", horizontal="center")
alt_fill = PatternFill("solid", fgColor="F2F2F2")
note_fill = PatternFill("solid", fgColor="FFF2CC")
criteria_fill = PatternFill("solid", fgColor="D6EAF8")
primary_fill = PatternFill("solid", fgColor="C6EFCE")


def style_header(ws, row, cols):
    for col in range(1, cols + 1):
        cell = ws.cell(row=row, column=col)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(wrap_text=True, vertical="center", horizontal="center")
        cell.border = thin


def style_body(ws, start_row, end_row, cols):
    for r in range(start_row, end_row + 1):
        for c in range(1, cols + 1):
            cell = ws.cell(row=r, column=c)
            cell.font = body_font
            cell.alignment = wrap
            cell.border = thin
            if r % 2 == 0:
                cell.fill = alt_fill


def set_widths(ws, widths):
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w


# ============================================================
# Sheet 1: 規模基準（規模を考える上で基準になる要素）
# ============================================================
ws1 = wb.active
ws1.title = "1_規模基準"

ws1["A1"] = "日本の化学企業｜規模を考える上で基準になる要素（全要素）"
ws1["A1"].font = title_font
ws1.merge_cells("A1:F1")

ws1["A2"] = (
    "本表の企業順位は「連結売上高」を主キーに降順ソートしています。"
    "ただし規模の多面評価には下表の全要素を併用するのが実務上妥当です。"
)
ws1["A2"].font = body_font
ws1["A2"].alignment = wrap
ws1.merge_cells("A2:F2")
ws1.row_dimensions[2].height = 36

headers1 = [
    "カテゴリ",
    "基準要素",
    "定義・見方",
    "規模判断での役割",
    "本表での使用",
    "留意点",
]
for i, h in enumerate(headers1, 1):
    ws1.cell(row=4, column=i, value=h)
style_header(ws1, 4, 6)

criteria = [
    ("財務規模", "連結売上高（Revenue）", "直近会計年度の連結売上収益/売上高", "事業規模の最基本指標。業界横断比較の主キー", "主キー（順位決定）", "決算期・会計基準差、事業構成（消費財含む）に注意"),
    ("財務規模", "営業利益 / 営業利益率", "本業の稼ぐ力と収益性", "「大きいが薄い」か「高収益ニッチ」かを判別", "補助列", "石化市況・半導体サイクルで大きく変動"),
    ("財務規模", "当期純利益", "最終的な利益規模", "株主還元余力・財務健全性の代理", "補助（参考）", "特別損益・減損で一時的に歪む"),
    ("財務規模", "時価総額（Market Cap）", "株価×発行済株式数", "市場が評価する企業価値規模", "補助列（概算）", "日々変動。非上場は算出不可"),
    ("財務規模", "総資産 / 純資産", "バランスシート規模", "装置産業の資本集約度を示す", "補助（参考）", "持分法・リース会計の影響あり"),
    ("財務規模", "EBITDA / EV/EBITDA", "キャッシュ創出と企業価値倍率", "買収・投資判断での規模・割安感", "補助（参考）", "定義差（特別項目除外）に注意"),
    ("財務規模", "研究開発費（R&D費）", "年間R&D支出額・売上高比率", "技術投資規模＝将来競争力の代理", "補助列（概算）", "開示粒度が企業により異なる"),
    ("財務規模", "設備投資額（Capex）", "有形固定資産等への投資", "装置産業の拡大・更新投資規模", "補助（参考）", "大型案件年でスパイク"),
    ("組織規模", "連結従業員数", "グループ人員規模", "組織・オペレーション規模", "補助列（概算）", "外注・協力会社を含まない"),
    ("組織規模", "単体従業員数", "親会社本体の人員", "本社機能・コア技術者規模の目安", "補助（参考）", "ホールディングス体制で小さく見える"),
    ("組織規模", "グループ会社数 / 海外拠点数", "子会社・海外法人の広がり", "グローバル展開規模", "補助（参考）", "カウント定義が企業ごとに異なる"),
    ("組織規模", "生産拠点数 / プラント能力", "工場数・主要製品能力（例: エチレン）", "製造キャパシティ規模", "補助（参考）", "能力公表は限定的"),
    ("組織規模", "R&D拠点数・研究者数", "研究所・技術センターの数と人員", "技術開発インフラ規模", "本表で主要拠点を記載", "公開情報は拠点名・所在地中心"),
    ("事業規模", "事業セグメント数 / 多角化度", "石化・機能材・医薬・住宅等の広がり", "コングロマリット型の総合力", "業態区分列で表現", "セグメント再編で比較困難"),
    ("事業規模", "主力製品の世界シェア", "シリコンウェハ・塩ビ等のシェア", "特定領域での支配的規模", "備考で補足", "非公開・推定が多い"),
    ("事業規模", "海外売上高比率", "海外/連結売上", "グローバル規模", "補助（参考）", "為替影響大"),
    ("市場規模", "証券コード / 上場市場", "東証プライム等", "資本市場アクセスと開示水準", "掲載", "非上場大手（例: 一部専業）は別扱い"),
    ("市場規模", "信用格付", "R&I・JCR等", "資金調達規模・信用力", "補助（参考）", "未取得企業あり"),
    ("影響力", "業界団体での位置づけ", "日化協加盟・役員企業等", "業界内プレゼンス", "補助（参考）", "定量化しにくい"),
    ("影響力", "特許出願・保有件数", "技術資産の量的指標", "知財規模", "補助（参考）", "質（被引用）も重要"),
    ("本表固有", "本社所在地", "登記上/実質本社", "意思決定拠点", "掲載", "東京本社＋大阪本社の二本社制あり"),
    ("本表固有", "主要R&D拠点", "代表的研究所・技術センター", "技術開発の地理的拠点", "掲載", "全拠点ではなく主要拠点を抜粋"),
    ("本表固有", "メールドメイン / 命名法則", "公式ドメインとローカル部規則", "対外連絡・営業リスト作成の実務", "掲載", "個人アドレスは推定例。実在人物の私的アドレスではない"),
]

for i, row in enumerate(criteria):
    r = 5 + i
    for c, v in enumerate(row, 1):
        ws1.cell(row=r, column=c, value=v)
        if row[4].startswith("主キー"):
            ws1.cell(row=r, column=5).fill = primary_fill
        elif "本表" in row[0] or row[4].startswith("掲載"):
            ws1.cell(row=r, column=5).fill = criteria_fill

style_body(ws1, 5, 4 + len(criteria), 6)
# re-apply special fills after style_body
for i, row in enumerate(criteria):
    r = 5 + i
    if row[4].startswith("主キー"):
        ws1.cell(row=r, column=5).fill = primary_fill
    elif row[4].startswith("掲載") or "本表で" in row[4]:
        ws1.cell(row=r, column=5).fill = criteria_fill

ws1["A29"] = "【順位の作り方】"
ws1["A29"].font = subtitle_font
ws1["A30"] = (
    "1) 主キー：連結売上高（EDINET等の直近開示、百万円単位を億円換算）で降順。"
    "2) 同順位時のタイブレーク：時価総額→営業利益→連結従業員数の順を推奨。"
    "3) 花王・資生堂・ユニ・チャーム等は東証「化学」セクターに含まれるが、"
    "消費財比重が大きい。純粋な素材・石化比較ではセグメント売上で再集計するのが望ましい。"
    "4) 東レは繊維・炭素繊維大手だが業種分類上「繊維製品」となることが多く、本表の化学セクター順位からは除外し参考注記とする。"
)
ws1["A30"].font = body_font
ws1["A30"].alignment = wrap
ws1["A30"].fill = note_fill
ws1.merge_cells("A30:F30")
ws1.row_dimensions[30].height = 70

set_widths(ws1, [12, 28, 40, 36, 22, 40])
ws1.freeze_panes = "A5"
ws1.auto_filter.ref = f"A4:F{4 + len(criteria)}"

# ============================================================
# Sheet 2: 企業規模ランキング
# ============================================================
ws2 = wb.create_sheet("2_企業規模ランキング")

ws2["A1"] = "日本の化学企業｜規模順リスト（連結売上高 大きい順）"
ws2["A1"].font = title_font
ws2.merge_cells("A1:R1")

ws2["A2"] = (
    "データ基準：金融庁EDINET開示ベースの化学セクター売上高ランキング（CorpoLog集計）を主に使用。"
    "売上高は億円（概算）。時価総額・従業員・R&Dは公開情報に基づく概算で、時点により変動。"
    "メール例は命名法則に沿った仮想例（info / 姓.名 形式）であり、特定個人の実在アドレスを示しません。"
)
ws2["A2"].font = body_font
ws2["A2"].alignment = wrap
ws2["A2"].fill = note_fill
ws2.merge_cells("A2:R2")
ws2.row_dimensions[2].height = 48

headers2 = [
    "順位",
    "企業名",
    "証券コード",
    "業態区分",
    "連結売上高（億円・概算）",
    "営業利益（億円・概算）",
    "営業利益率（%・概算）",
    "時価総額（億円・概算）",
    "連結従業員数（概算）",
    "R&D費（億円・概算）",
    "本社所在地",
    "主要R&D拠点",
    "公式サイトドメイン",
    "代表メールドメイン",
    "メール命名法則",
    "代表メール例（仮想）",
    "主力事業・備考",
    "データ注記",
]

for i, h in enumerate(headers2, 1):
    ws2.cell(row=4, column=i, value=h)
style_header(ws2, 4, len(headers2))
ws2.row_dimensions[4].height = 40

# Data: rank, name, code, type, revenue(oku), op_profit, op_margin, mcap, employees, rd,
# hq, rd_sites, web_domain, mail_domain, mail_rule, mail_examples, business, note
# Revenue from EDINET/CorpoLog (億円)
# Other figures approximated from Buffett Code / public IR (2024-2025期付近)

companies = [
    (1, "三菱ケミカルグループ", "4188", "総合化学HD", 37040, 1967, 4.5, 13640, 66000, 1500,
     "東京都千代田区丸の内1-1-1（パレスビル）※事業会社三菱ケミカル本社",
     "横浜 Science & Innovation Center（青葉区鴨志田町）；鶴見；筑波（牛久）；三重（四日市）；福岡（北九州）ほか",
     "www.mcgc.com / www.m-chemical.co.jp",
     "mcgc.com / m-chemical.co.jp",
     "①info@ / pr@ / ir@ 等の汎用 ②姓名ローマ字（姓.名 または 名.姓）③グループ再編でドメイン併存に注意",
     "info@mcgc.com；ir@mcgc.com；yamada.taro@m-chemical.co.jp",
     "機能商品・素材・ヘルスケア（旧三菱ケミカルHD）。国内最大級の総合化学",
     "売上は直近EDINET開示ベース。HDと事業会社でドメインが分かれる"),
    (2, "富士フイルムホールディングス", "4901", "機能化学・ヘルスケア", 33570, 3302, 10.3, 37650, 74000, 1800,
     "東京都港区西麻布2-26-30 / 東京ミッドタウン本社機能",
     "神奈川県南足柄（先進研究所）；神奈川県足柄；富山；米国・欧州のバイオ拠点",
     "www.fujifilm.com",
     "fujifilm.com / fujifilm.co.jp",
     "①グローバルは fujifilm.com ②国内は co.jp 併存 ③名.姓@ または 部門エイリアス",
     "info@fujifilm.com；recruit@fujifilm.co.jp；taro.yamada@fujifilm.com",
     "ヘルスケア・高機能材料・イメージング。化学セクター売上2位級",
     "イメージング出自だが材料・医薬材料比重大"),
    (3, "旭化成", "3407", "総合化学・住宅・医薬", 30750, 2119, 7.0, 20320, 49000, 900,
     "東京都千代田区有楽町1-1-2 日比谷三井タワー（東京ミッドタウン日比谷）",
     "静岡県富士市（富士支社・研究開発）；宮崎県延岡；川崎；水島（倉敷）ほか",
     "www.asahi-kasei.com",
     "asahi-kasei.co.jp / asahi-kasei.com",
     "①asahi-kasei.co.jp が国内標準 ②姓_名 / 姓.名 ③事業会社（ホームズ等）は別ドメイン",
     "info@asahi-kasei.co.jp；yamada.t@asahi-kasei.co.jp",
     "マテリアル・住宅・ヘルスケアの3領域",
     "従業員・本社は会社概要（2026/3時点公表値付近）"),
    (4, "信越化学工業", "4063", "機能化学・半導体材料", 25740, 7421, 29.0, 102600, 27000, 700,
     "東京都千代田区丸の内1-4-1 丸の内永楽ビルディング",
     "群馬県安中（シリコーン電子材料技術研・精密機能材料研）；茨城県神栖（塩ビ・高分子）；新潟県上越（合成技術・新機能材料）；福井県越前（磁性材料）",
     "www.shinetsu.co.jp",
     "shinetsu.co.jp",
     "①shinetsu.co.jp に統一傾向 ②総務・広報経由の問い合わせが基本 ③個人は 姓 または 頭文字.姓",
     "info@shinetsu.co.jp；pr@shinetsu.co.jp；t.yamada@shinetsu.co.jp",
     "塩ビ・シリコーン・半導体シリコンウェハで世界トップクラス。利益率・時価総額突出",
     "時価総額は市況で変動大。規模＝売上だけでなく収益力でも突出"),
    (5, "住友化学", "4005", "総合化学・農薬・電子材料", 23290, 1930, 7.4, 7720, 32000, 1700,
     "東京本社：東京都中央区日本橋2-7-1 東京日本橋タワー／大阪本社：大阪市中央区北浜4-5-33",
     "愛媛（新居浜）；大阪（此花）；千葉（袖ケ浦）；宝塚；筑波；大江ほか多数の研究所",
     "www.sumitomo-chem.co.jp",
     "sumitomo-chem.co.jp",
     "①sumitomo-chem.co.jp ②姓_名@ または 姓.名@ ③事業部エイリアスあり",
     "info@sumitomo-chem.co.jp；yamada@sumitomo-chem.co.jp",
     "石化・情報電子化学・健康・農業関連",
     "二本社制。R&D拠点が国内に広く分散"),
    (6, "日本ペイントホールディングス", "4612", "塗料・表面処理", 17740, 1876, 11.5, 23690, 40000, 250,
     "大阪府大阪市北区大淀北2-1-2（大阪本社）／東京オフィスあり",
     "大阪府寝屋川・大阪研究所系；海外（アジア・欧州）技術センター多数",
     "www.nipponpaint-holdings.com",
     "nipponpaint.jp / nipponpaint.com",
     "①HDと地域法人でドメイン分岐 ②contact@ / info@ ③現地法人は国別ドメイン",
     "info@nipponpaint.jp；contact@nipponpaint-holdings.com",
     "自動車・建築用塗料で世界有数。海外売上比率が高い",
     "持株会社体制。製造・研究は事業会社側"),
    (7, "花王", "4452", "消費財化学・ケミカル", 16890, 1466, 9.0, 28320, 35000, 600,
     "東京都中央区日本橋茅場町1-14-10",
     "栃木県（栃木研究所）；和歌山；東京・隅田ほか；ケミカル事業の技術拠点",
     "www.kao.com",
     "kao.com / kao.co.jp",
     "①kao.co.jp / kao.com ②名.姓@（グローバル寄り）③採用は専用ポータル",
     "info@kao.com；taro.yamada@kao.co.jp",
     "化粧品・洗剤＋油脂・界面活性剤等ケミカル",
     "化学セクターだが消費財比重大。素材比較時は要調整"),
    (8, "三井化学", "4183", "総合化学", 16690, 783, 4.3, 8480, 19000, 500,
     "東京都中央区八重洲2-2-1 東京ミッドタウン八重洲",
     "VISION HUB® SODEGAURA（千葉県袖ケ浦市長浦）；茂原研究・開発センター；Design & Solution Center（新潟県三条）",
     "jp.mitsuichemicals.com / www.mitsuichemicals.com",
     "mitsuichemicals.com",
     "①mitsuichemicals.com が標準 ②姓.名@ ③旧 mitsui-chem.co.jp からの移行に注意",
     "info@mitsuichemicals.com；yamada.taro@mitsuichemicals.com",
     "ライフ＆ヘルスケア、モビリティ、ICT、基盤素材",
     "袖ケ浦に研究者が集約"),
    (9, "日本酸素ホールディングス", "4091", "産業ガス", 13600, 1659, 12.7, 19700, 20000, 350,
     "東京都品川区大海1-3-26 豊洲フォレシアほか（公表本社所在地に準拠）",
     "山梨・川崎等の技術開発拠点；海外（Matheson等）技術センター",
     "www.nipponsanso-hd.co.jp",
     "nipponsanso-hd.co.jp / tnsco.co.jp",
     "①HDドメインと事業会社ドメイン併存 ②info@ / 問い合わせフォーム優先",
     "info@nipponsanso-hd.co.jp",
     "酸素・窒素・アルゴン等産業ガスのグローバル大手",
     "旧大陽日酸グループ。海外比率高"),
    (10, "レゾナック・ホールディングス", "4004", "機能化学・半導体材料", 13470, 890, 6.4, 15520, 26000, 500,
     "東京都港区芝大門1-13-9",
     "神奈川県川崎（旧昭和電工研）；群馬；大分；千葉ほか半導体材料開発拠点",
     "www.resonac.com",
     "resonac.com / resonac.co.jp",
     "①resonac.com へブランド統一 ②旧 sdk.co.jp（昭和電工）からの変更に注意 ③姓.名@",
     "info@resonac.com；yamada.taro@resonac.com",
     "半導体材料・無機/有機化学品。旧昭和電工",
     "社名変更後のドメイン切替過渡期あり"),
    (11, "積水化学工業", "4204", "住宅・高機能プラスチックス", 13090, 1080, 8.3, 11490, 27000, 350,
     "大阪府大阪市北区西天満2-4-4／東京支社：東京都港区虎ノ門",
     "京都・水無瀬；茨城；大阪ほかの研究所・開発センター",
     "www.sekisuichemical.com",
     "sekisui.com / sekisui.co.jp",
     "①sekisui.co.jp / sekisui.com ②部門コード付きエイリアスあり",
     "info@sekisui.co.jp",
     "住宅、環境・ライフライン、高機能プラスチックス",
     "住宅セグメントが大きく素材専業と比較注意"),
    (12, "エア・ウォーター", "4088", "産業ガス・総合", 10670, 752, 7.0, 5420, 20000, 180,
     "大阪府大阪市中央区南船場2-12-8",
     "北海道・関西等の開発拠点；医療・ガス技術センター",
     "www.awi.co.jp",
     "awi.co.jp",
     "①awi.co.jp ②問い合わせは Webフォーム多用 ③姓名ローマ字",
     "info@awi.co.jp",
     "産業ガス、医療、農業・食品関連",
     "地域密着型M&Aで規模拡大"),
    (13, "DIC", "4631", "インキ・顔料・機能材料", 10520, 445, 4.2, 3570, 21000, 200,
     "東京都中央区日本橋3-7-20 ディーアイシービル",
     "千葉県佐倉（総合研究所）；小田原；海外技術拠点",
     "www.dic-global.com",
     "dic.co.jp / dic-global.com",
     "①dic.co.jp ②グローバルは dic-global.com ③姓.名@",
     "info@ma.dic.co.jp；yamada@dic.co.jp",
     "印刷インキ世界大手、顔料・樹脂・電子材料",
     "旧大日本インキ化学"),
    (14, "日東電工", "6988", "機能フィルム・電子材料", 10280, 1857, 18.3, 23640, 28000, 400,
     "大阪府茨木市下穂積1-1-2",
     "大阪府茨木・豊橋；滋賀；東北等の開発拠点；海外R&D",
     "www.nitto.com",
     "nitto.com / nitto.co.jp",
     "①nitto.co.jp / nitto.com ②名.姓@（海外連携）③採用専用アドレス",
     "info@nitto.co.jp；taro.yamada@nitto.com",
     "光学フィルム、電子材料、産業テープ",
     "高収益の機能材メーカー"),
    (15, "東ソー", "4042", "総合化学・塩ビ・機能材", 10200, 989, 9.3, 8010, 14000, 250,
     "東京都港区芝3-8-2 芝公園ファーストビル",
     "山口県周南（南陽）研究所；四日市；東京研究センター系",
     "www.tosoh.co.jp",
     "tosoh.co.jp",
     "①tosoh.co.jp ②姓.名@ または 姓@",
     "info@tosoh.co.jp；yamada@tosoh.co.jp",
     "クロルアルカリ、石油化学、機能材料（ジルコニア・石英等）",
     "周南コンビナート中核"),
    (16, "資生堂", "4911", "化粧品（化学セクター分類）", 9700, 76, 0.8, 10370, 33000, 300,
     "東京都港区東新橋1-8-10",
     "横浜（グローバルイノベーションセンター）；掛川ほか",
     "corp.shiseido.com",
     "shiseido.com / shiseido.co.jp",
     "①shiseido.com ②ブランド別ドメイン多数 ③素材メーカー比較時は除外推奨",
     "corp.contact@shiseido.com",
     "化粧品・スキンケア中心",
     "業種「化学」だが消費財。素材ランキングからは外す判断も可"),
    (17, "ユニ・チャーム", "8113", "衛生材料（化学セクター分類）", 9450, 820, 8.3, 16310, 16000, 180,
     "東京都港区三田3-5-27 住友不動産三田ツインビル西館",
     "香川県観音寺・四国拠点；海外開発拠点",
     "www.unicharm.co.jp",
     "unicharm.co.jp",
     "①unicharm.co.jp ②問い合わせフォーム中心",
     "info@unicharm.co.jp",
     "紙おむつ・生理用品・ペットケア",
     "不織布・吸収体材料技術。消費財分類寄り"),
    (18, "カネカ", "4118", "機能樹脂・食品・医薬", 8120, 401, 5.0, 2880, 12000, 150,
     "大阪府大阪市北区中之島2-3-18／東京：港区赤坂",
     "大阪府摂津；兵庫県高砂；滋賀；海外R&D",
     "www.kaneka.co.jp",
     "kaneka.co.jp",
     "①kaneka.co.jp ②姓_名@ 形式が多い",
     "info@kaneka.co.jp；yamada_taro@kaneka.co.jp",
     "変成シリコーン、光学材料、医薬原薬、食品",
     "大阪発の機能化学"),
    (19, "クラレ", "3405", "機能樹脂・繊維", 8080, 851, 10.3, 5100, 12000, 140,
     "東京都千代田区大手町2-6-4 常盤橋タワー",
     "岡山県倉敷（清音・玉島）；新潟；茨城；海外技術拠点",
     "www.kuraray.com",
     "kuraray.com / kuraray.co.jp",
     "①kuraray.co.jp / kuraray.com ②名.姓@",
     "info@kuraray.co.jp；taro.yamada@kuraray.com",
     "ポバール、EVAL、人工皮革クラリーノ、エラストマー",
     "高機能ポリマーに強み"),
    (20, "三菱瓦斯化学", "4182", "ガス化学・半導体材料", 7380, 509, 6.6, 6040, 8000, 150,
     "東京都千代田区丸の内2-5-2 三菱ビル",
     "新潟；東京研究所；山形；茨城（鹿島）ほか",
     "www.mgc.co.jp",
     "mgc.co.jp",
     "①mgc.co.jp ②姓.名@ ③エンジニアリング塑料・電子材料部門エイリアス",
     "info@mgc.co.jp；yamada.taro@mgc.co.jp",
     "メタノール、過酸化水素、半導体用薬品・BT材料",
     "電子材料で存在感"),
    (21, "関西ペイント", "4613", "塗料", 5900, 521, 8.8, 4500, 16000, 130,
     "大阪府大阪市中央区今橋2-6-14",
     "神奈川県平塚；大阪；海外技術センター",
     "www.kansai.co.jp",
     "kansai.co.jp",
     "①kansai.co.jp ②info@ / 営業所アドレス",
     "info@kansai.co.jp",
     "自動車・建築用塗料",
     "日本ペイントと並ぶ国内塗料大手"),
    (22, "ダイセル", "4202", "セルロース・有機化学品", 5800, 610, 10.4, 3890, 11000, 220,
     "大阪府大阪市北区大深町3-1 グランフロント大阪",
     "兵庫県姫路（広畑）；新井；東京；イノベーション・パーク系",
     "www.daicel.com",
     "daicel.com / daicel.co.jp",
     "①daicel.com へ寄せる傾向 ②姓.名@",
     "info@jp.daicel.com；yamada.taro@daicel.com",
     "酢酸セルロース、火工品（エアバッグ用）、有機合成",
     "ヘルスケア・スマート材料へシフト"),
    (23, "UBE", "4208", "化学品・建設資材", 4620, 180, 3.7, 2620, 8000, 130,
     "東京都港区芝浦1-2-1 シーバンスN館",
     "山口県宇部；堺；千葉；海外R&D",
     "www.ube.com",
     "ube.com / ube.co.jp",
     "①ube.com / ube-ind.co.jp 系の移行履歴あり ②旧宇部興産ドメインに注意",
     "info@ube.com",
     "ナイロン・ファイン化学品、ポリイミド、機械・セメント関連",
     "社名をUBEへ。ドメインもブランド統一中"),
    (24, "ライオン", "4912", "日用品・化学品", 4220, 284, 6.9, 4670, 8000, 100,
     "東京都墨田区本所1-3-7",
     "東京都江戸川区（研究開発本部）；小田原ほか",
     "www.lion.co.jp",
     "lion.co.jp",
     "①lion.co.jp ②問い合わせフォーム中心",
     "info@lion.co.jp",
     "洗剤・オーラルケア＋化学品",
     "消費財比重大"),
    (25, "東洋紡", "3101", "フィルム・機能素材", 4220, 200, 4.7, 1800, 10000, 150,
     "大阪府大阪市北区梅田1-13-1 大阪梅田ツインタワーズ・サウス",
     "滋賀県大津・犬上；福井；富山の研究所",
     "www.toyobo.co.jp",
     "toyobo.jp / toyobo.co.jp",
     "①toyobo.co.jp ②姓.名@",
     "info@toyobo.jp",
     "包装フィルム、工業フィルム、バイオ・ヘルスケア材料",
     "繊維出自の機能素材メーカー"),
    (26, "ADEKA", "4401", "添加剤・食品・化学品", 4170, 410, 10.1, 4640, 5000, 120,
     "東京都荒川区東尾久7-2-35",
     "埼玉県浦和；千葉；海外技術拠点",
     "www.adeka.co.jp",
     "adeka.co.jp",
     "①adeka.co.jp ②姓.名@",
     "info@adeka.co.jp；yamada.taro@adeka.co.jp",
     "ポリマー添加剤、食品、電子材料、化学品",
     "高付加価値添加剤に強み"),
    (27, "日本ゼオン", "4205", "合成ゴム・光学材料", 4120, 293, 7.0, 3790, 4000, 100,
     "東京都千代田区丸の内1-6-2 新丸の内センタービル",
     "神奈川県川崎（総合開発センター）；富山；三重",
     "www.zeon.co.jp",
     "zeon.co.jp",
     "①zeon.co.jp ②姓.名@",
     "info@zeon.co.jp",
     "合成ゴム、光学フィルム、スペシャリティケミカルズ",
     "光学用途で高シェア製品あり"),
    (28, "日本触媒", "4114", "触媒・吸水性樹脂", 4000, 191, 4.7, 3340, 5000, 110,
     "大阪府大阪市中央区高麗橋4-1-1 画廊ビル",
     "大阪府吹田（研究所）；姫路；川崎",
     "www.shokubai.co.jp",
     "shokubai.co.jp",
     "①shokubai.co.jp ②姓.名@",
     "info@shokubai.co.jp",
     "アクリル酸、高吸水性樹脂（SAP）、触媒",
     "SAP世界有数"),
    (29, "デンカ", "4061", "無機・有機・電子材料", 3840, 144, 3.6, 2630, 7000, 130,
     "東京都中央区日本橋室町2-1-1 日本橋三井タワー",
     "新潟県青海；千葉；大牟田；伊勢崎の技術・開発拠点",
     "www.denka.co.jp",
     "denka.co.jp",
     "①denka.co.jp ②姓.名@",
     "info@denka.co.jp",
     "クロロプレン、電子材料、ワクチン原液等",
     "旧電気化学工業"),
    (30, "artience（旧東洋インキSC）", "4634", "インキ・機能材料", 3500, 204, 5.8, 2500, 8000, 120,
     "東京都中央区京橋2-2-1 京橋エドグラン",
     "埼玉県川越；東京；海外技術拠点",
     "www.artiencegroup.com",
     "artiencegroup.com / toyoink.co.jp",
     "①ブランド変更で artience 系へ ②旧 toyoink.co.jp 併存期間あり",
     "info@artiencegroup.com",
     "印刷材料、カラー、機能材料",
     "2024年東洋インキSCホールディングスから商号変更"),
    (31, "トクヤマ", "4043", "無機・電子材料・セメント", 3495, 250, 7.0, 2800, 6000, 100,
     "山口県周南市御影町1-1（本社）／東京本社：千代田区",
     "周南の研究部門；つくば；鹿島",
     "www.tokuyama.co.jp",
     "tokuyama.co.jp",
     "①tokuyama.co.jp ②姓.名@",
     "info@tokuyama.co.jp",
     "ソーダ灰、多結晶シリコン、電子材料、セメント",
     "周南に一大拠点"),
    (32, "住友ベークライト", "4203", "フェノール樹脂・高機能材料", 3200, 280, 8.5, 3500, 8000, 150,
     "東京都品川区東品川2-5-8 天王洲パークサイドビル",
     "栃木；静岡；神戸；宇都宮の研究所・技術センター",
     "www.sumibe.co.jp",
     "sumibe.co.jp",
     "①sumibe.co.jp ②姓.名@",
     "info@sumibe.co.jp",
     "半導体封止材、高機能プラスチック、医療材料",
     "フェノール樹脂世界大手級"),
    (33, "日産化学", "4021", "農薬・電子材料・化学品", 2796, 600, 21.0, 9000, 3000, 200,
     "東京都中央区日本橋2-5-1 日本橋高島屋三井ビルディング",
     "千葉県袖ケ浦；埼玉県白岡；富山の研究所",
     "www.nissanchem.co.jp",
     "nissanchem.co.jp",
     "①nissanchem.co.jp ②姓.名@ ③高利益率ニッチ型",
     "info@nissanchem.co.jp；yamada.taro@nissanchem.co.jp",
     "農薬、半導体材料、化学品、ヘルスケア",
     "売上規模以上に利益率・時価総額が高い典型"),
    (34, "日油", "4403", "油脂・化成品・ライフサイエンス", 2580, 350, 13.0, 5500, 4000, 100,
     "東京都渋谷区恵比寿4-20-3 恵比寿ガーデンプレイスタワー",
     "兵庫県尼崎；愛知；茨城の研究拠点",
     "www.nof.co.jp",
     "nof.co.jp",
     "①nof.co.jp ②姓.名@",
     "info@nof.co.jp",
     "脂肪酸誘導体、有機過酸化物、DDS材料",
     "ライフサイエンス材料が成長領域"),
    (35, "日本化薬", "4272", "機能化学品・医薬・安全性", 2419, 280, 11.0, 3200, 5000, 120,
     "東京都千代田区丸の内2-1-1 明治安田生命ビル",
     "東京都北区；群馬；山口の研究所",
     "www.nipponkayaku.co.jp",
     "nipponkayaku.co.jp",
     "①nipponkayaku.co.jp ②姓.名@",
     "info@nipponkayaku.co.jp",
     "機能化学品、医薬品、セイフティシステムズ",
     "火薬出自の多角化企業"),
    (36, "東京応化工業", "4186", "半導体フォトレジスト", 2370, 500, 21.0, 12000, 2000, 150,
     "神奈川県川崎市中原区中丸子150",
     "神奈川県相模原・御殿場；郡山；海外（韓国・台湾・米）技術拠点",
     "www.tok.co.jp",
     "tok.co.jp",
     "①tok.co.jp ②姓.名@ ③半導体サプライチェーン向け英語名併記多い",
     "info@tok.co.jp；yamada.taro@tok.co.jp",
     "フォトレジスト・半導体用高純度薬液",
     "売上規模は中位だが時価総額・利益率は上位級"),
    (37, "高砂香料工業", "4914", "香料・フレーバー", 2251, 180, 8.0, 2800, 3000, 80,
     "東京都大田区蒲田5-37-1",
     "神奈川県平塚；静岡；海外クリエイションセンター",
     "www.takasago.com",
     "takasago.com / takasago.co.jp",
     "①takasago.com ②名.姓@",
     "info@takasago.com",
     "フレーバー・フレグランス、アロマケミカル",
     "香料世界大手の一角"),
    (38, "東亞合成", "4045", "アクリル・接着剤", 1623, 150, 9.0, 2200, 3000, 70,
     "東京都港区西新橋1-14-1",
     "名古屋；徳島；埼玉の研究所",
     "www.toagosei.co.jp",
     "toagosei.co.jp",
     "①toagosei.co.jp ②姓.名@",
     "info@toagosei.co.jp",
     "アクリル製品、瞬間接着剤アロンアルフア、高純度無機",
     "機能化学品中心"),
    (39, "クレハ", "4023", "機能樹脂・化学品・包装", 1617, 180, 11.0, 2500, 4000, 90,
     "東京都中央区日本橋浜町3-3-2",
     "茨城県いわき；栃木；大阪の開発拠点",
     "www.kureha.co.jp",
     "kureha.co.jp",
     "①kureha.co.jp ②姓.名@",
     "info@kureha.co.jp",
     "PVDF、PGA、家庭用ラップ、農薬",
     "高機能樹脂（電池バインダー等）に強み"),
    (40, "JSR", "4185※", "半導体材料・ライフサイエンス", 4100, 200, 5.0, "非上場(概算)", 8000, 200,
     "東京都港区東新橋1-9-2 汐留シティセンター",
     "三重県四日市；神奈川；筑波；海外半導体材料R&D",
     "www.jsr.co.jp",
     "jsr.co.jp",
     "①jsr.co.jp ②姓.名@ ③非公開化後もドメイン維持",
     "info@jsr.co.jp；yamada.taro@jsr.co.jp",
     "半導体材料、ディスプレイ材料、ライフサイエンス",
     "2024年非公開化。売上は直近公開期の参考値。順位は参考扱い"),
]

# Insert Toray as note row after main list - actually include as reference
# Adjust JSR rank note - place after 39 as reference

for i, row in enumerate(companies):
    r = 5 + i
    for c, v in enumerate(row, 1):
        cell = ws2.cell(row=r, column=c, value=v)
        if c in (1, 3, 5, 6, 7, 8, 9, 10):
            cell.alignment = center

style_body(ws2, 5, 4 + len(companies), len(headers2))

# Reference section: Toray
ref_row = 5 + len(companies) + 1
ws2.cell(row=ref_row, column=1, value="【参考】業種分類上「化学」以外だが素材・化学に近い大手")
ws2.cell(row=ref_row, column=1).font = subtitle_font
ws2.merge_cells(start_row=ref_row, start_column=1, end_row=ref_row, end_column=18)

toray = ( "-", "東レ", "3402", "繊維・炭素繊維・樹脂（参考）", 25000, 1200, 5.0, 18000, 48000, 700,
    "東京都中央区日本橋室町2-1-1 日本橋三井タワー",
    "滋賀県大津（滋賀事業場・研究）；愛媛；名古屋；鎌倉の複合材料開発ほか",
    "www.toray.co.jp",
    "toray.co.jp / toray.com",
    "①toray.co.jp ②姓.名@ または 社員番号系 ③グローバル toray.com",
    "info@mail.toray",
    "合成繊維、炭素繊維複合材料、樹脂・ケミカル、ライフサイエンス",
    "東証業種は繊維製品。化学企業比較では売上上位相当のため参考掲載")
# fix toray email example
toray = list(toray)
toray[15] = "info@mail.toray.co.jp；yamada.t.aa@mail.toray.co.jp"
toray = tuple(toray)

r = ref_row + 1
for c, v in enumerate(toray, 1):
    cell = ws2.cell(row=r, column=c, value=v)
    cell.font = body_font
    cell.alignment = wrap
    cell.border = thin
    cell.fill = note_fill

ws2.cell(row=r + 2, column=1, value="【凡例】売上・利益・時価総額・従業員・R&Dは概算。正確な投資判断には各社有価証券報告書・決算短信を参照してください。メール例の個人名部分は法則説明用の仮名です。")
ws2.cell(row=r + 2, column=1).font = Font(name="Yu Gothic", size=9, italic=True)
ws2.merge_cells(start_row=r + 2, start_column=1, end_row=r + 2, end_column=18)

set_widths(ws2, [6, 22, 10, 18, 14, 12, 12, 12, 12, 10, 36, 40, 28, 24, 36, 36, 32, 28])
ws2.freeze_panes = "C5"
ws2.auto_filter.ref = f"A4:R{4 + len(companies)}"
ws2.row_dimensions[4].height = 45

# ============================================================
# Sheet 3: メールアドレス法則
# ============================================================
ws3 = wb.create_sheet("3_メールアドレス法則")

ws3["A1"] = "化学企業メールアドレス｜代表例と命名法則"
ws3["A1"].font = title_font
ws3.merge_cells("A1:E1")

ws3["A2"] = (
    "日本の大手化学企業では、個人アドレスの公開を控え、問い合わせフォームや代表・部門エイリアスを使うケースが増えています。"
    "営業・IR・採用では下表の法則で推定・確認するのが一般的です。スパム・無断送信は避け、公式の問い合わせ経路を優先してください。"
)
ws3["A2"].font = body_font
ws3["A2"].alignment = wrap
ws3["A2"].fill = note_fill
ws3.merge_cells("A2:E2")
ws3.row_dimensions[2].height = 48

ws3["A4"] = "A. ドメイン法則"
ws3["A4"].font = subtitle_font

headers3a = ["法則ID", "法則", "説明", "代表例", "該当しやすい企業"]
for i, h in enumerate(headers3a, 1):
    ws3.cell(row=5, column=i, value=h)
style_header(ws3, 5, 5)

domain_rules = [
    ("D1", "英字社名 + .co.jp", "国内法人の標準。.co.jp は日本登記法人向け属性型JPドメイン", "tosoh.co.jp / kaneka.co.jp / zeon.co.jp", "東ソー、カネカ、ゼオン、トクヤマ等"),
    ("D2", "短縮・略称 + .co.jp", "長い社名をローマ字略称化（頭文字連結や慣用略称）", "mgc.co.jp（Mitsubishi Gas Chemical）/ awi.co.jp（Air Water）/ tok.co.jp（Tokyo Ohka）", "三菱ガス化学、エア・ウォーター、東京応化"),
    ("D3", "ハイフン区切り社名", "複合語をハイフンで可読化", "sumitomo-chem.co.jp / asahi-kasei.co.jp / m-chemical.co.jp", "住友化学、旭化成、三菱ケミカル"),
    ("D4", "グローバル .com 統一", "海外売上比率が高い企業は .com を主ドメイン化", "resonac.com / mitsuichemicals.com / nitto.com / kao.com", "レゾナック、三井化学、日東電工、花王"),
    ("D5", "HD / 事業会社でドメイン分離", "持株会社と事業会社・ブランドでメールドメインが異なる", "mcgc.com（G）と m-chemical.co.jp（事業会社）", "三菱ケミカルG、日本酸素HD、日本ペイントHD"),
    ("D6", "社名変更後の移行ドメイン", "再編・リブランディングで旧ドメインが一定期間生き残る", "旧 sdk.co.jp → resonac.com／旧 toyoink → artience／旧 ube-ind", "レゾナック、artience、UBE"),
    ("D7", "mail. サブドメイン", "メール専用サブドメインを切る企業がある", "〜@mail.toray.co.jp", "東レ（参考）など"),
]

for i, row in enumerate(domain_rules):
    r = 6 + i
    for c, v in enumerate(row, 1):
        ws3.cell(row=r, column=c, value=v)
style_body(ws3, 6, 5 + len(domain_rules), 5)

ws3["A14"] = "B. ローカル部（@より前）の命名法則"
ws3["A14"].font = subtitle_font

headers3b = ["法則ID", "パターン", "書式テンプレート", "仮想例（山田太郎の場合）", "備考"]
for i, h in enumerate(headers3b, 1):
    ws3.cell(row=15, column=i, value=h)
style_header(ws3, 15, 5)

local_rules = [
    ("L1", "代表・汎用エイリアス", "info@ / contact@ / pr@ / ir@ / recruit@ / support@", "info@example.co.jp", "最も安全に使える公開アドレス。まずここを当たる"),
    ("L2", "姓のみ", "lastname@", "yamada@example.co.jp", "大企業では衝突回避で減少傾向"),
    ("L3", "姓.名", "lastname.firstname@", "yamada.taro@example.co.jp", "日系メーカーで最も多い形式の一つ"),
    ("L4", "名.姓", "firstname.lastname@", "taro.yamada@example.co.jp", "外資連携・グローバル企業で多い"),
    ("L5", "頭文字.姓", "f.lastname@", "t.yamada@example.co.jp", "信越化学などで見られるコンパクト型"),
    ("L6", "姓_名", "lastname_firstname@", "yamada_taro@example.co.jp", "アンダースコア区切り派"),
    ("L7", "姓 + 名頭文字", "lastnamef@ / lastname_t@", "yamadat@example.co.jp", "同一姓の衝突回避"),
    ("L8", "部門・工場コード", "dept@ / plant-code@", "polymer-sales@example.co.jp", "部署共通アドレス。異動に強い"),
    ("L9", "数字・社員番号付与", "lastname123@ / id@", "yamada01@example.co.jp", "大規模組織の一意性確保"),
    ("L10", "問い合わせフォームのみ", "（公開個人メールなし）", "Webフォーム経由", "近年増加。ドメインだけ分かっても個人は引けない"),
]

for i, row in enumerate(local_rules):
    r = 16 + i
    for c, v in enumerate(row, 1):
        ws3.cell(row=r, column=c, value=v)
style_body(ws3, 16, 15 + len(local_rules), 5)

ws3["A28"] = "C. 主要企業ドメイン早見（代表）"
ws3["A28"].font = subtitle_font

headers3c = ["企業名", "代表ドメイン", "よく使う公開エイリアス例", "個人アドレス推定の第一候補", "注意"]
for i, h in enumerate(headers3c, 1):
    ws3.cell(row=29, column=i, value=h)
style_header(ws3, 29, 5)

quick = [
    ("三菱ケミカルグループ", "mcgc.com / m-chemical.co.jp", "info@ / ir@", "yamada.taro@m-chemical.co.jp", "HDと事業会社を混同しない"),
    ("富士フイルムHD", "fujifilm.com / fujifilm.co.jp", "info@ / recruit@", "taro.yamada@fujifilm.com", ".com/.co.jp併存"),
    ("旭化成", "asahi-kasei.co.jp", "info@", "yamada.t@asahi-kasei.co.jp", "事業会社は別ドメイン"),
    ("信越化学工業", "shinetsu.co.jp", "info@ / pr@", "t.yamada@shinetsu.co.jp", "個人公開は少ない"),
    ("住友化学", "sumitomo-chem.co.jp", "info@", "yamada@sumitomo-chem.co.jp", "ハイフンドメイン"),
    ("三井化学", "mitsuichemicals.com", "info@", "yamada.taro@mitsuichemicals.com", "旧ドメインからの移行"),
    ("レゾナックHD", "resonac.com", "info@", "yamada.taro@resonac.com", "旧昭和電工ドメイン注意"),
    ("東ソー", "tosoh.co.jp", "info@", "yamada@tosoh.co.jp", "シンプル.co.jp"),
    ("DIC", "dic.co.jp", "info@", "yamada@dic.co.jp", "グローバルサイトは別"),
    ("日東電工", "nitto.com / nitto.co.jp", "info@", "taro.yamada@nitto.com", "高海外比率"),
    ("三菱ガス化学", "mgc.co.jp", "info@", "yamada.taro@mgc.co.jp", "略称3文字"),
    ("東京応化工業", "tok.co.jp", "info@", "yamada.taro@tok.co.jp", "半導体材料メーカー"),
    ("日産化学", "nissanchem.co.jp", "info@", "yamada.taro@nissanchem.co.jp", "自動車の日産自動車とは別"),
    ("UBE", "ube.com", "info@", "yamada.taro@ube.com", "旧ube-ind系に注意"),
    ("JSR", "jsr.co.jp", "info@", "yamada.taro@jsr.co.jp", "非公開化後もドメイン継続"),
]

for i, row in enumerate(quick):
    r = 30 + i
    for c, v in enumerate(row, 1):
        ws3.cell(row=r, column=c, value=v)
style_body(ws3, 30, 29 + len(quick), 5)

ws3["A47"] = "D. 実務上の確認手順（推奨）"
ws3["A47"].font = subtitle_font
ws3["A48"] = (
    "1) 公式サイトの「お問い合わせ」「IR」「採用」から公開エイリアスまたはフォームを確認する。\n"
    "2) メールを送る必要がある場合は、名刺・展示会・プレスリリース・学会予稿の連絡先を一次情報とする。\n"
    "3) ドメインだけが分かっている状態で個人名を機械生成して大量送信しない（迷惑メール・個人情報保護の観点）。\n"
    "4) 社名変更直後は旧ドメインがバウンスまたは転送されることがあるため、両方を記録する。\n"
    "5) グループ会社（例: 旭化成ホームズ、三井化学クロップ＆ライフソリューション）は親会社とドメインが異なる。"
)
ws3["A48"].font = body_font
ws3["A48"].alignment = wrap
ws3["A48"].fill = criteria_fill
ws3.merge_cells("A48:E48")
ws3.row_dimensions[48].height = 90

set_widths(ws3, [10, 28, 55, 42, 36])
ws3.freeze_panes = "A6"

# ============================================================
# Sheet 4: 注記・出典
# ============================================================
ws4 = wb.create_sheet("4_注記と出典")
ws4["A1"] = "注記・出典・使い方"
ws4["A1"].font = title_font

notes = [
    "【作成目的】日本の化学企業を「規模」で横断比較し、本社・R&D拠点・メール連絡の実務情報を1冊にまとめる。",
    "【主キー】連結売上高（大きい順）。EDINET開示を集計した化学セクターランキング（CorpoLog等）を基準。",
    "【補助指標】営業利益・利益率・時価総額・従業員・R&D費を併記し、売上だけでは見えない「稼ぐ規模」「組織規模」「投資規模」を可視化。",
    "【時点】売上高は直近提出の有価証券報告書ベース（企業によりFY表記が異なる）。時価総額は変動するため概算。",
    "【消費財の扱い】花王・資生堂・ユニ・チャーム・ライオンは東証「化学」に含まれうるが、素材・石化比較ではセグメント切り出しを推奨。",
    "【参考掲載】東レは業種「繊維製品」だが素材企業として規模が大きいため参考行に記載。",
    "【JSR】非公開化後は時価総額なし。売上は直近公開期の参考値で順位も参考扱い。",
    "【メール】個人名付きアドレスは法則説明用の仮想例。実在の個人を指さない。連絡は公式窓口を優先。",
    "【出典例】金融庁EDINET；各社コーポレートサイト（拠点一覧・会社概要）；Buffett Code 業界ページ；CorpoLog Chemicals Revenue Ranking。",
    "【更新方法】年1回、各社有報の売上高・従業員を更新し、社名変更・本社移転・ドメイン変更を反映する。",
]

for i, t in enumerate(notes):
    ws4.cell(row=3 + i, column=1, value=t).font = body_font
    ws4.cell(row=3 + i, column=1).alignment = wrap
    ws4.merge_cells(start_row=3 + i, start_column=1, end_row=3 + i, end_column=4)
    ws4.row_dimensions[3 + i].height = 30

ws4["A15"] = "シート構成"
ws4["A15"].font = subtitle_font
ws4["A16"] = "1_規模基準 … 規模判定に使う全要素と本表での役割"
ws4["A17"] = "2_企業規模ランキング … 売上順の企業一覧（本社・R&D・メール法則付き）"
ws4["A18"] = "3_メールアドレス法則 … ドメイン法則・ローカル部法則・早見表"
ws4["A19"] = "4_注記と出典 … 本ファイルの前提と更新方針"
for r in range(16, 20):
    ws4.cell(row=r, column=1).font = body_font

set_widths(ws4, [100, 20, 20, 20])

# ============================================================
# Sheet 0-like summary at front? Already have sheet1 as criteria
# Add a compact ranking-only sheet for quick view
# ============================================================
ws0 = wb.create_sheet("0_早見ランキング", 0)
ws0["A1"] = "早見｜連結売上高順（上位40＋参考）"
ws0["A1"].font = title_font
ws0.merge_cells("A1:F1")

headers0 = ["順位", "企業名", "証券コード", "売上高（億円）", "本社（都道府県）", "メールドメイン"]
for i, h in enumerate(headers0, 1):
    ws0.cell(row=3, column=i, value=h)
style_header(ws0, 3, 6)

def pref(hq: str) -> str:
    for p in ["東京都", "大阪府", "神奈川県", "山口県", "愛知県", "千葉県"]:
        if hq.startswith(p):
            return p
    if "東京" in hq[:10]:
        return "東京都"
    if "大阪" in hq[:10]:
        return "大阪府"
    return hq.split("県")[0] + "県" if "県" in hq[:8] else hq[:6]

for i, row in enumerate(companies):
    r = 4 + i
    domain = row[13].split("/")[0].strip()
    vals = [row[0], row[1], row[2], row[4], pref(row[10]), domain]
    for c, v in enumerate(vals, 1):
        ws0.cell(row=r, column=c, value=v)
style_body(ws0, 4, 3 + len(companies), 6)

# toray ref
r = 4 + len(companies)
ws0.cell(row=r, column=1, value="-")
ws0.cell(row=r, column=2, value="東レ（参考・繊維製品）")
ws0.cell(row=r, column=3, value="3402")
ws0.cell(row=r, column=4, value=25000)
ws0.cell(row=r, column=5, value="東京都")
ws0.cell(row=r, column=6, value="toray.co.jp")
for c in range(1, 7):
    ws0.cell(row=r, column=c).fill = note_fill
    ws0.cell(row=r, column=c).border = thin
    ws0.cell(row=r, column=c).font = body_font

set_widths(ws0, [8, 28, 12, 16, 14, 28])
ws0.freeze_panes = "A4"
ws0.auto_filter.ref = f"A3:F{3 + len(companies)}"

out = "/workspace/日本化学企業_規模順リスト.xlsx"
wb.save(out)
print(f"Saved: {out}")
print(f"Sheets: {wb.sheetnames}")
print(f"Companies: {len(companies)}")
