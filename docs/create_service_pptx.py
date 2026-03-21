"""
ハナサク サービス紹介資料 PPTX生成スクリプト

AICU標準PPTXレイアウトルール（aicu-spring-fes-2026/docs/create_pptx.py）に準拠。
カラーのみハナサクブランドに差し替え。
"""
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
import re
import os

# --- ハナサクブランドカラー（AICU標準のダークネイビーに相当する位置づけ） ---
BRAND_DARK = RGBColor(0x5A, 0x20, 0x30)       # 表紙・セクション背景（ダーク）
BRAND_MID = RGBColor(0x7A, 0x30, 0x45)        # セクション区切り
BRAND_PRIMARY = RGBColor(0xC0, 0x60, 0x70)     # アクセント
BRAND_LIGHT = RGBColor(0xF5, 0xE6, 0xEA)       # 交互行背景
BRAND_SUBTITLE = RGBColor(0xE8, 0xA0, 0xA8)    # サブタイトル文字

TEXT_MAIN = RGBColor(0x33, 0x33, 0x33)
TEXT_NOTE = RGBColor(0x88, 0x88, 0x88)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)


# ============================================================
# AICU標準ヘルパー関数（カラーのみ差し替え）
# ============================================================

def add_title_slide(prs, title, subtitle=""):
    """表紙スライド — ダーク背景 + 白文字"""
    slide = prs.slides.add_slide(prs.slide_layouts[6])  # blank
    bg = slide.background
    fill = bg.fill
    fill.solid()
    fill.fore_color.rgb = BRAND_DARK

    left, top, width, height = Inches(0.8), Inches(1.5), Inches(8.4), Inches(2)
    txBox = slide.shapes.add_textbox(left, top, width, height)
    tf = txBox.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = title
    p.font.size = Pt(32)
    p.font.bold = True
    p.font.color.rgb = WHITE
    p.alignment = PP_ALIGN.LEFT

    if subtitle:
        p2 = tf.add_paragraph()
        p2.text = subtitle
        p2.font.size = Pt(16)
        p2.font.color.rgb = BRAND_SUBTITLE
        p2.alignment = PP_ALIGN.LEFT
    return slide


def add_section_slide(prs, title):
    """セクション区切りスライド — ダーク背景 + 白文字中央"""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    bg = slide.background
    fill = bg.fill
    fill.solid()
    fill.fore_color.rgb = BRAND_MID

    left, top, width, height = Inches(0.8), Inches(2.5), Inches(8.4), Inches(1.5)
    txBox = slide.shapes.add_textbox(left, top, width, height)
    tf = txBox.text_frame
    p = tf.paragraphs[0]
    p.text = title
    p.font.size = Pt(28)
    p.font.bold = True
    p.font.color.rgb = WHITE
    p.alignment = PP_ALIGN.CENTER
    return slide


def add_content_slide(prs, title, bullets, note=""):
    """コンテンツスライド — タイトル+アンダーライン+bullet"""
    slide = prs.slides.add_slide(prs.slide_layouts[6])

    # Title
    left, top, width, height = Inches(0), Inches(0), Inches(10), Inches(1.0)
    txBox = slide.shapes.add_textbox(left, top, width, height)
    tf = txBox.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = "  " + title
    p.font.size = Pt(22)
    p.font.bold = True
    p.font.color.rgb = BRAND_DARK
    p.alignment = PP_ALIGN.LEFT

    # Underline
    shape = slide.shapes.add_shape(1, Inches(0.5), Inches(0.95), Inches(9), Inches(0.03))
    shape.fill.solid()
    shape.fill.fore_color.rgb = BRAND_DARK
    shape.line.fill.background()

    # Content
    txBox2 = slide.shapes.add_textbox(Inches(0.6), Inches(1.2), Inches(8.8), Inches(5.5))
    tf2 = txBox2.text_frame
    tf2.word_wrap = True

    for i, bullet in enumerate(bullets):
        if i == 0:
            p = tf2.paragraphs[0]
        else:
            p = tf2.add_paragraph()

        indent = 0
        text = bullet
        if text.startswith("    ") or text.startswith("  "):
            indent = 1
            text = text.strip()

        text = text.lstrip("- ").lstrip("• ")
        text = re.sub(r'\*\*(.*?)\*\*', r'\1', text)
        text = re.sub(r'\[(.*?)\]', r'\1', text)

        p.text = text
        p.font.size = Pt(14) if indent == 0 else Pt(12)
        p.font.color.rgb = TEXT_MAIN
        p.space_after = Pt(6)
        p.level = indent

        if "◎" in bullet or "★" in bullet or bullet.strip().startswith("→"):
            p.font.bold = True
            p.font.color.rgb = BRAND_DARK

    if note:
        p_note = tf2.add_paragraph()
        p_note.text = ""
        p_note = tf2.add_paragraph()
        p_note.text = note
        p_note.font.size = Pt(11)
        p_note.font.italic = True
        p_note.font.color.rgb = TEXT_NOTE

    return slide


def add_table_slide(prs, title, headers, rows):
    """テーブルスライド — AICU標準準拠"""
    slide = prs.slides.add_slide(prs.slide_layouts[6])

    # Title
    txBox = slide.shapes.add_textbox(Inches(0), Inches(0), Inches(10), Inches(1.0))
    tf = txBox.text_frame
    p = tf.paragraphs[0]
    p.text = "  " + title
    p.font.size = Pt(22)
    p.font.bold = True
    p.font.color.rgb = BRAND_DARK

    # Underline
    shape = slide.shapes.add_shape(1, Inches(0.5), Inches(0.95), Inches(9), Inches(0.03))
    shape.fill.solid()
    shape.fill.fore_color.rgb = BRAND_DARK
    shape.line.fill.background()

    # Table
    num_rows = len(rows) + 1
    num_cols = len(headers)
    table_width = Inches(9)

    table_shape = slide.shapes.add_table(
        num_rows, num_cols, Inches(0.5), Inches(1.3),
        table_width, Inches(0.4 * min(num_rows, 12))
    )
    table = table_shape.table

    # Header
    for j, header in enumerate(headers):
        cell = table.cell(0, j)
        cell.text = re.sub(r'\*\*(.*?)\*\*', r'\1', header)
        for paragraph in cell.text_frame.paragraphs:
            paragraph.font.size = Pt(11)
            paragraph.font.bold = True
            paragraph.font.color.rgb = WHITE
        cell.fill.solid()
        cell.fill.fore_color.rgb = BRAND_DARK

    # Data
    for i, row in enumerate(rows):
        for j, val in enumerate(row):
            if j < num_cols:
                cell = table.cell(i + 1, j)
                cell.text = re.sub(r'\*\*(.*?)\*\*', r'\1', str(val))
                for paragraph in cell.text_frame.paragraphs:
                    paragraph.font.size = Pt(10)
                    paragraph.font.color.rgb = TEXT_MAIN
                if i % 2 == 0:
                    cell.fill.solid()
                    cell.fill.fore_color.rgb = BRAND_LIGHT

    return slide


# ============================================================
# ハナサク サービス紹介資料
# ============================================================

def create_presentation():
    prs = Presentation()
    prs.slide_width = Inches(10)
    prs.slide_height = Inches(7.5)

    # --- 表紙 ---
    add_title_slide(prs,
        "ハナサク\nサービス紹介資料",
        "福井の女の子が、自分の言葉で未来を語れるようになる就活伴走サービス\n\n西川 舞衣子（EMKO）")

    # --- サービス概要 ---
    add_content_slide(prs, "サービス概要", [
        "「ハナサク」は、福井県の女子学生とその保護者を対象としたオンライン就活伴走コーチングです。",
        "就活塾でもエージェントでもない、あなたの「伴走者」として一貫してサポートします。",
        "",
        "◎ 就活伴走コーチング",
        "  対象: 女子学生 + 保護者",
        "  形式: オンライン面談（30分/回）",
        "  LINE質問対応・親御さんへの月次レポート",
        "",
        "◎ 女性リーダー向けメンタリング",
        "  対象: 管理職・リーダー候補の女性",
        "  形式: オンライン1on1（45分/回）",
        "  キャリア相談・壁打ち",
    ])

    # --- こんな悩みありませんか ---
    add_content_slide(prs, "こんな悩み、ありませんか？", [
        "◎ 学生さん",
        "  キャリアセンターの面談が表面的で、本当に知りたいことが聞けない",
        "  東京の就活塾は高すぎるし、地方から通えない",
        "  面接で「結婚は？出産は？」と聞かれたとき、どう答えればいいかわからない",
        "  地元に残るか東京に出るか、相談できる大人がいない",
        "",
        "◎ 親御さん",
        "  子供の就活状況が見えず不安（LINEで聞いても「大丈夫」としか返ってこない）",
        "  自分たちの時代と就活が違いすぎて、アドバイスできない",
        "  高い学費を払っているのに、キャリアセンターだけで大丈夫か心配",
        "",
        "◎ 女性リーダー",
        "  管理職になったが、相談できる同じ立場の人がいない",
        "  チームマネジメントや社内政治で壁を感じている",
        "  仕事とライフイベントの両立に悩んでいる",
    ])

    # --- 3つの特徴 ---
    add_content_slide(prs, "ハナサクの3つの特徴", [
        "★ 1. 福井出身コーチによる1on1面談",
        "  福井県出身・高志高校卒のコーチが、地元の空気感を理解した上であなたの言葉を引き出します",
        "  オンライン（Google Meet）なので、全国どこからでも受けられます",
        "",
        "★ 2. 親子で安心の伴走型サポート",
        "  長期プランでは親御さんへの月次レポートを標準提供",
        "  お子さんの成長ポイント、来月の計画、親へのお願い事項をお伝えします",
        "  プライバシーは厳守します",
        "",
        "★ 3. 女子学生専門だからできること",
        "  女性活躍推進プログラムで社外メンターを務めた経験を活かし、女子学生特有の悩みに寄り添います",
        "  外見への言及、ライフプランへの圧、自己主張への抵抗感にも対応",
    ])

    # --- 料金（就活） ---
    add_table_slide(prs,
        "料金プラン｜就活伴走コーチング（30分/回）",
        ["プラン", "料金（税込）", "内容", "1回あたり"],
        [
            ["都度払い", "3,000円/回", "オンライン面談 1回", "3,000円"],
            ["4回チケット", "10,000円", "好きなタイミングで4回（有効期限3か月）", "2,500円"],
            ["3か月伴走プラン", "15,000円", "月2回×3か月（計6回）+ LINE質問対応", "2,500円"],
            ["6か月伴走プラン", "27,000円", "月2回×6か月（計12回）+ LINE + 月次レポート", "2,250円"],
        ])

    # --- 料金（女性リーダー） ---
    add_table_slide(prs,
        "料金プラン｜女性リーダー向けメンタリング（45分/回）",
        ["プラン", "料金（税込）", "内容", "1回あたり"],
        [
            ["都度払い", "4,000円/回", "オンライン1on1 1回", "4,000円"],
            ["3か月プラン", "20,000円", "月2回×3か月（計6回）+ LINE相談対応", "約3,300円"],
        ])

    # 注記を手動追加
    slide = prs.slides[-1]
    txBox = slide.shapes.add_textbox(Inches(0.6), Inches(3.5), Inches(8.8), Inches(1.5))
    tf = txBox.text_frame
    tf.word_wrap = True
    for i, text in enumerate(["初回面談は無料", "お友達紹介で、紹介した方・された方どちらも次回1,000円OFF"]):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.text = "★ " + text
        p.font.size = Pt(13)
        p.font.bold = True
        p.font.color.rgb = BRAND_DARK
        p.space_after = Pt(8)

    # --- 競合比較 ---
    add_table_slide(prs,
        "他の就活サービスとの違い",
        ["", "大学キャリアセンター", "就活塾（東京系）", "就活エージェント", "ハナサク"],
        [
            ["料金", "無料", "15〜40万円", "無料", "3,000円/回〜"],
            ["地方対応", "各大学内", "東京圏中心", "オンライン可", "福井特化"],
            ["女子特化", "なし", "なし", "一部あり", "専門"],
            ["親への対応", "なし", "なし", "なし", "月次レポート"],
            ["継続性", "単発", "3〜6か月", "内定まで", "最大2年伴走"],
            ["運営者", "大学職員", "サラリーマン講師", "人材会社社員", "社外メンター経験者"],
        ])

    # --- コーチ紹介 ---
    add_content_slide(prs, "コーチ紹介", [
        "◎ 西川 舞衣子（にしかわ まいこ）",
        "",
        "- 福井県出身 / 高志高校卒",
        "- デジタルハリウッド大学大学院修了",
        "- 労働市場などの調査を行うシンクタンクに勤務",
        "- 大手企業の女性活躍推進プログラムで社外メンターを務めた経験",
        "- 2児の母",
        "",
        "→ 福井の風土を知り、東京の就活事情も知る立場から、",
        "  地方の女子学生が「私はこう生きたい」と自分の言葉で語れるようになることを目指しています。",
    ])

    # --- ご利用の流れ ---
    add_content_slide(prs, "ご利用の流れ", [
        "★ STEP 1：無料相談を予約",
        "  Googleカレンダーから空き日時を選ぶだけ（アカウント登録不要）",
        "",
        "★ STEP 2：初回面談（無料）",
        "  現状をヒアリングし、最適なプランをご提案",
        "",
        "★ STEP 3：プランを選んでスタート",
        "  お支払いはクレジットカード / Apple Pay / Google Pay",
        "",
        "★ STEP 4：定期面談で伴走",
        "  一緒に「あなたの言葉」を見つけていきます",
    ], note="無理な勧誘は一切しません。まずはお気軽にご相談ください。")

    # --- お問い合わせ（セクションスライドで締め） ---
    slide = add_section_slide(prs, "まずは無料で相談してみませんか？")
    # 追加テキスト
    txBox = slide.shapes.add_textbox(Inches(0.8), Inches(4.0), Inches(8.4), Inches(2.5))
    tf = txBox.text_frame
    tf.word_wrap = True
    contacts = [
        "LP: https://mai-san447.github.io/hanasaku-career/",
        "無料相談予約: https://calendar.app.google/u8SkpMApaMfaRoSx6",
        "LINE公式: https://lin.ee/qURXYyS",
        "",
        "© 2026 ハナサク / EMKO",
    ]
    for i, text in enumerate(contacts):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.text = text
        p.font.size = Pt(14) if i < 3 else Pt(11)
        p.font.color.rgb = BRAND_SUBTITLE if i < 3 else TEXT_NOTE
        p.alignment = PP_ALIGN.CENTER
        p.space_after = Pt(8)

    # Save
    output_dir = os.path.dirname(os.path.abspath(__file__))
    output_path = os.path.join(output_dir, "hanasaku_service_overview.pptx")
    prs.save(output_path)
    print(f"Created: {output_path}")


if __name__ == "__main__":
    create_presentation()
