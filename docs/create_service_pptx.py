from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.oxml.ns import qn
import os

# ハナサクのブランドカラー
PRIMARY = RGBColor(0xC0, 0x60, 0x70)       # #C06070
PRIMARY_DARK = RGBColor(0x9A, 0x40, 0x50)  # #9A4050
PRIMARY_LIGHT = RGBColor(0xF5, 0xE6, 0xEA) # #F5E6EA
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
TEXT_DARK = RGBColor(0x3A, 0x3A, 0x3A)
TEXT_LIGHT = RGBColor(0x88, 0x88, 0x88)
BG_WARM = RGBColor(0xFD, 0xF8, 0xF6)


def set_slide_bg(slide, color):
    bg = slide.background
    fill = bg.fill
    fill.solid()
    fill.fore_color.rgb = color


def add_textbox(slide, left, top, width, height, text, font_size=14,
                bold=False, color=TEXT_DARK, align=PP_ALIGN.LEFT, line_spacing=1.5):
    txBox = slide.shapes.add_textbox(Inches(left), Inches(top), Inches(width), Inches(height))
    tf = txBox.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = text
    p.font.size = Pt(font_size)
    p.font.bold = bold
    p.font.color.rgb = color
    p.alignment = align
    p.line_spacing = Pt(font_size * line_spacing)
    return tf


def add_multiline(slide, left, top, width, height, lines, font_size=12,
                  color=TEXT_DARK, bold=False, align=PP_ALIGN.LEFT, line_spacing=1.6):
    txBox = slide.shapes.add_textbox(Inches(left), Inches(top), Inches(width), Inches(height))
    tf = txBox.text_frame
    tf.word_wrap = True
    for i, line in enumerate(lines):
        if i == 0:
            p = tf.paragraphs[0]
        else:
            p = tf.add_paragraph()
        p.text = line
        p.font.size = Pt(font_size)
        p.font.color.rgb = color
        p.font.bold = bold
        p.alignment = align
        p.line_spacing = Pt(font_size * line_spacing)
    return tf


def add_rect(slide, left, top, width, height, fill_color, line=False):
    shape = slide.shapes.add_shape(1, Inches(left), Inches(top), Inches(width), Inches(height))
    shape.fill.solid()
    shape.fill.fore_color.rgb = fill_color
    if not line:
        shape.line.fill.background()
    return shape


def add_table_slide(prs, title, headers, rows, col_widths=None):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    set_slide_bg(slide, WHITE)

    # Title bar
    add_rect(slide, 0, 0, 10, 0.9, PRIMARY)
    add_textbox(slide, 0.5, 0.15, 9, 0.6, title, font_size=20, bold=True, color=WHITE)

    # Table
    num_rows = len(rows) + 1
    num_cols = len(headers)
    tbl_left = Inches(0.4)
    tbl_top = Inches(1.2)
    tbl_width = Inches(9.2)
    tbl_height = Inches(0.4 * num_rows)

    table_shape = slide.shapes.add_table(num_rows, num_cols, tbl_left, tbl_top, tbl_width, tbl_height)
    table = table_shape.table

    if col_widths:
        for i, w in enumerate(col_widths):
            table.columns[i].width = Inches(w)

    # Header row
    for j, h in enumerate(headers):
        cell = table.cell(0, j)
        cell.text = h
        for paragraph in cell.text_frame.paragraphs:
            paragraph.font.size = Pt(10)
            paragraph.font.bold = True
            paragraph.font.color.rgb = WHITE
            paragraph.alignment = PP_ALIGN.CENTER
        # Header bg
        tcPr = cell._tc.get_or_add_tcPr()
        solidFill = tcPr.makeelement(qn('a:solidFill'), {})
        srgbClr = solidFill.makeelement(qn('a:srgbClr'), {'val': 'C06070'})
        solidFill.append(srgbClr)
        tcPr.append(solidFill)

    # Data rows
    for i, row in enumerate(rows):
        for j, val in enumerate(row):
            cell = table.cell(i + 1, j)
            cell.text = val
            for paragraph in cell.text_frame.paragraphs:
                paragraph.font.size = Pt(9)
                paragraph.font.color.rgb = TEXT_DARK
                paragraph.alignment = PP_ALIGN.CENTER if j > 0 else PP_ALIGN.LEFT
            # Alternate row bg
            bg_val = 'F5E6EA' if i % 2 == 0 else 'FFFFFF'
            tcPr = cell._tc.get_or_add_tcPr()
            solidFill = tcPr.makeelement(qn('a:solidFill'), {})
            srgbClr = solidFill.makeelement(qn('a:srgbClr'), {'val': bg_val})
            solidFill.append(srgbClr)
            tcPr.append(solidFill)

    return slide


def create_presentation():
    prs = Presentation()
    prs.slide_width = Inches(10)
    prs.slide_height = Inches(7.5)

    # ===== Slide 1: 表紙 =====
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    set_slide_bg(slide, WHITE)

    # 上部の装飾バー
    add_rect(slide, 0, 0, 10, 0.15, PRIMARY)

    # 花の装飾（丸）
    shape = slide.shapes.add_shape(
        9, Inches(7.5), Inches(0.5), Inches(2.5), Inches(2.5)  # oval
    )
    shape.fill.solid()
    shape.fill.fore_color.rgb = PRIMARY_LIGHT
    shape.line.fill.background()

    shape2 = slide.shapes.add_shape(
        9, Inches(8.2), Inches(5.5), Inches(1.8), Inches(1.8)
    )
    shape2.fill.solid()
    shape2.fill.fore_color.rgb = RGBColor(0xFA, 0xF0, 0xF0)
    shape2.line.fill.background()

    add_textbox(slide, 0.8, 2.0, 7, 1.2, "ハナサク", font_size=44, bold=True, color=PRIMARY)
    add_textbox(slide, 0.8, 3.3, 7, 0.8,
                "福井の女の子が、自分の言葉で未来を語れるようになる就活伴走サービス",
                font_size=16, color=TEXT_DARK)
    add_textbox(slide, 0.8, 4.5, 7, 0.5, "サービス紹介資料", font_size=14, color=TEXT_LIGHT)
    add_textbox(slide, 0.8, 6.5, 7, 0.4, "西川 舞衣子（EMKO）", font_size=11, color=TEXT_LIGHT)

    # 下部バー
    add_rect(slide, 0, 7.35, 10, 0.15, PRIMARY)

    # ===== Slide 2: サービス概要 =====
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    set_slide_bg(slide, WHITE)
    add_rect(slide, 0, 0, 10, 0.9, PRIMARY)
    add_textbox(slide, 0.5, 0.15, 9, 0.6, "サービス概要", font_size=20, bold=True, color=WHITE)

    add_multiline(slide, 0.6, 1.3, 8.8, 2.5, [
        "「ハナサク」は、福井県の女子学生とその保護者を対象とした",
        "オンライン就活伴走コーチングサービスです。",
        "",
        "就活塾でもエージェントでもない、あなたの「伴走者」として、",
        "面接対策から親御さんへの進捗レポートまで一貫してサポートします。",
    ], font_size=14, color=TEXT_DARK, line_spacing=1.8)

    # 2つのサービス枠
    add_rect(slide, 0.6, 4.0, 4.0, 2.2, PRIMARY_LIGHT)
    add_textbox(slide, 0.9, 4.2, 3.4, 0.5, "就活伴走コーチング", font_size=14, bold=True, color=PRIMARY_DARK)
    add_multiline(slide, 0.9, 4.8, 3.4, 1.2, [
        "対象: 女子学生 + 保護者",
        "形式: オンライン面談（30分）",
        "LINE質問対応・月次レポート",
    ], font_size=11, color=TEXT_DARK, line_spacing=1.8)

    add_rect(slide, 5.4, 4.0, 4.0, 2.2, PRIMARY_LIGHT)
    add_textbox(slide, 5.7, 4.2, 3.4, 0.5, "女性リーダー向けメンタリング", font_size=14, bold=True, color=PRIMARY_DARK)
    add_multiline(slide, 5.7, 4.8, 3.4, 1.2, [
        "対象: 管理職・リーダー候補の女性",
        "形式: オンライン1on1（45分）",
        "キャリア相談・壁打ち",
    ], font_size=11, color=TEXT_DARK, line_spacing=1.8)

    # ===== Slide 3: こんな悩み、ありませんか？ =====
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    set_slide_bg(slide, WHITE)
    add_rect(slide, 0, 0, 10, 0.9, PRIMARY)
    add_textbox(slide, 0.5, 0.15, 9, 0.6, "こんな悩み、ありませんか？", font_size=20, bold=True, color=WHITE)

    # 学生カード
    add_rect(slide, 0.4, 1.2, 4.3, 2.6, RGBColor(0xFD, 0xF0, 0xF2))
    add_textbox(slide, 0.6, 1.35, 3.9, 0.4, "学生さん", font_size=13, bold=True, color=PRIMARY)
    add_multiline(slide, 0.6, 1.8, 3.9, 2.0, [
        "・キャリアセンターの面談が表面的",
        "・東京の就活塾は高すぎて通えない",
        "・面接で「結婚は？」と聞かれたら…",
        "・地元に残るか東京に出るか相談できない",
    ], font_size=11, color=TEXT_DARK, line_spacing=1.8)

    # 親カード
    add_rect(slide, 5.3, 1.2, 4.3, 2.6, RGBColor(0xE6, 0xF0, 0xEA))
    add_textbox(slide, 5.5, 1.35, 3.9, 0.4, "親御さん", font_size=13, bold=True, color=RGBColor(0x5A, 0x7A, 0x6A))
    add_multiline(slide, 5.5, 1.8, 3.9, 2.0, [
        "・子供の就活状況が見えず不安",
        "・自分たちの時代と就活が違いすぎる",
        "・キャリアセンターだけで大丈夫？",
        "・「オヤカク」文化、どう関わるべき？",
    ], font_size=11, color=TEXT_DARK, line_spacing=1.8)

    # 女性リーダーカード
    add_rect(slide, 0.4, 4.2, 9.2, 2.2, RGBColor(0xF0, 0xE8, 0xF5))
    add_textbox(slide, 0.6, 4.35, 8.8, 0.4, "女性リーダー", font_size=13, bold=True, color=RGBColor(0x7A, 0x5A, 0x8A))
    add_multiline(slide, 0.6, 4.8, 4.0, 1.5, [
        "・管理職になったが相談相手がいない",
        "・チームマネジメントや社内政治で壁を感じる",
    ], font_size=11, color=TEXT_DARK, line_spacing=1.8)
    add_multiline(slide, 5.0, 4.8, 4.0, 1.5, [
        "・仕事とライフイベントの両立に悩んでいる",
        "・女性リーダーとしてのロールモデルが少ない",
    ], font_size=11, color=TEXT_DARK, line_spacing=1.8)

    # ===== Slide 4: 3つの特徴 =====
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    set_slide_bg(slide, WHITE)
    add_rect(slide, 0, 0, 10, 0.9, PRIMARY)
    add_textbox(slide, 0.5, 0.15, 9, 0.6, "ハナサクの3つの特徴", font_size=20, bold=True, color=WHITE)

    features = [
        ("01", "福井出身コーチによる\n1on1面談",
         "福井県出身・高志高校卒のコーチが、\n地元の空気感を理解した上で\nあなたの言葉を引き出します。\nオンライン（Google Meet）で全国対応。"),
        ("02", "親子で安心の\n伴走型サポート",
         "長期プランでは親御さんへの\n月次レポートを標準提供。\nお子さんの成長・計画・お願い事項を\nお伝えします。プライバシーは厳守。"),
        ("03", "女子学生専門\nだからできること",
         "社外メンター経験を活かし、\n女子学生特有の悩みに寄り添います。\n外見への言及、ライフプランへの圧、\n自己主張への抵抗感にも対応。"),
    ]

    for i, (num, title, desc) in enumerate(features):
        x = 0.4 + i * 3.1
        # Number circle
        shape = slide.shapes.add_shape(9, Inches(x + 0.9), Inches(1.2), Inches(0.7), Inches(0.7))
        shape.fill.solid()
        shape.fill.fore_color.rgb = PRIMARY
        shape.line.fill.background()
        tf = shape.text_frame
        tf.paragraphs[0].text = num
        tf.paragraphs[0].font.size = Pt(18)
        tf.paragraphs[0].font.bold = True
        tf.paragraphs[0].font.color.rgb = WHITE
        tf.paragraphs[0].alignment = PP_ALIGN.CENTER
        tf.vertical_anchor = MSO_ANCHOR.MIDDLE

        add_textbox(slide, x, 2.2, 2.8, 0.9, title, font_size=13, bold=True, color=PRIMARY_DARK, align=PP_ALIGN.CENTER)
        add_textbox(slide, x, 3.3, 2.8, 3.0, desc, font_size=10, color=TEXT_DARK, align=PP_ALIGN.CENTER, line_spacing=1.8)

    # ===== Slide 5: 料金プラン（就活） =====
    add_table_slide(prs,
        "料金プラン｜就活伴走コーチング（30分/回）",
        ["プラン", "料金（税込）", "内容", "1回あたり"],
        [
            ["都度払い", "3,000円/回", "オンライン面談 1回", "3,000円"],
            ["4回チケット", "10,000円", "好きなタイミングで4回（有効期限3か月）", "2,500円"],
            ["3か月伴走プラン", "15,000円", "月2回×3か月（計6回）+ LINE質問対応", "2,500円"],
            ["6か月伴走プラン", "27,000円", "月2回×6か月（計12回）+ LINE + 月次レポート", "2,250円"],
        ],
        col_widths=[2.0, 1.5, 3.7, 1.5],
    )

    # ===== Slide 6: 料金プラン（女性リーダー） =====
    slide = add_table_slide(prs,
        "料金プラン｜女性リーダー向けメンタリング（45分/回）",
        ["プラン", "料金（税込）", "内容", "1回あたり"],
        [
            ["都度払い", "4,000円/回", "オンライン1on1 1回", "4,000円"],
            ["3か月プラン", "20,000円", "月2回×3か月（計6回）+ LINE相談対応", "約3,300円"],
        ],
        col_widths=[2.0, 1.5, 3.7, 1.5],
    )
    add_multiline(slide, 0.6, 3.5, 8.8, 1.5, [
        "・初回面談は無料",
        "・お友達紹介で、紹介した方・された方どちらも次回1,000円OFF",
    ], font_size=12, color=PRIMARY_DARK, bold=True, line_spacing=2.0)

    # ===== Slide 7: 競合比較 =====
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
        ],
        col_widths=[1.2, 2.0, 2.0, 2.0, 2.0],
    )

    # ===== Slide 8: コーチ紹介 =====
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    set_slide_bg(slide, WHITE)
    add_rect(slide, 0, 0, 10, 0.9, PRIMARY)
    add_textbox(slide, 0.5, 0.15, 9, 0.6, "コーチ紹介", font_size=20, bold=True, color=WHITE)

    # Profile area
    add_rect(slide, 0.6, 1.5, 8.8, 4.5, BG_WARM)

    add_textbox(slide, 1.0, 1.8, 7, 0.6, "西川 舞衣子（にしかわ まいこ）", font_size=20, bold=True, color=PRIMARY_DARK)
    add_textbox(slide, 1.0, 2.5, 7, 0.4, "ハナサク代表 / EMKO", font_size=12, color=TEXT_LIGHT)

    add_multiline(slide, 1.0, 3.2, 8.0, 2.5, [
        "福井県出身。高志高校卒業後、東京でキャリアを積み、",
        "デジタルハリウッド大学大学院修了。",
        "",
        "労働市場などの調査を行うシンクタンクに勤務。",
        "大手企業の女性活躍推進プログラムで社外メンターを務めた経験を持つ。2児の母。",
        "",
        "福井の風土を知り、東京の就活事情も知る立場から、地方の女子学生が",
        "「私はこう生きたい」と自分の言葉で語れるようになることを目指しています。",
    ], font_size=11, color=TEXT_DARK, line_spacing=1.8)

    # Tags
    tags = ["福井出身", "高志高校卒", "シンクタンク勤務", "社外メンター", "2児の母", "DHU大学院修了"]
    for i, tag in enumerate(tags):
        x = 1.0 + i * 1.4
        add_rect(slide, x, 5.3, 1.25, 0.35, PRIMARY_LIGHT)
        add_textbox(slide, x, 5.33, 1.25, 0.3, tag, font_size=8, bold=True, color=PRIMARY_DARK, align=PP_ALIGN.CENTER)

    # ===== Slide 9: ご利用の流れ =====
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    set_slide_bg(slide, WHITE)
    add_rect(slide, 0, 0, 10, 0.9, PRIMARY)
    add_textbox(slide, 0.5, 0.15, 9, 0.6, "ご利用の流れ", font_size=20, bold=True, color=WHITE)

    steps = [
        ("STEP 1", "無料相談を予約", "Googleカレンダーから\n空き日時を選ぶだけ\n（アカウント登録不要）"),
        ("STEP 2", "初回面談（無料）", "現状をヒアリングし\n最適なプランをご提案"),
        ("STEP 3", "プランを選んで\nスタート", "クレジットカード\nApple Pay / Google Pay"),
        ("STEP 4", "定期面談で伴走", "一緒に「あなたの言葉」を\n見つけていきます"),
    ]

    for i, (step, title, desc) in enumerate(steps):
        x = 0.3 + i * 2.4
        # Step badge
        add_rect(slide, x + 0.3, 1.3, 1.5, 0.35, PRIMARY)
        add_textbox(slide, x + 0.3, 1.32, 1.5, 0.3, step, font_size=10, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
        # Title
        add_textbox(slide, x, 1.9, 2.1, 0.8, title, font_size=13, bold=True, color=PRIMARY_DARK, align=PP_ALIGN.CENTER)
        # Desc
        add_textbox(slide, x, 2.9, 2.1, 1.5, desc, font_size=10, color=TEXT_DARK, align=PP_ALIGN.CENTER, line_spacing=1.8)

        # Arrow between steps
        if i < 3:
            add_textbox(slide, x + 2.1, 2.0, 0.3, 0.6, ">", font_size=20, bold=True, color=PRIMARY, align=PP_ALIGN.CENTER)

    add_multiline(slide, 0.6, 5.0, 8.8, 1.5, [
        "無理な勧誘は一切しません。まずはお気軽にご相談ください。",
    ], font_size=12, color=TEXT_LIGHT, align=PP_ALIGN.CENTER)

    # ===== Slide 10: お問い合わせ =====
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    set_slide_bg(slide, PRIMARY)

    add_textbox(slide, 0.5, 1.5, 9, 1.0,
                "まずは無料で相談してみませんか？",
                font_size=28, bold=True, color=WHITE, align=PP_ALIGN.CENTER)

    add_textbox(slide, 0.5, 2.8, 9, 0.5,
                "初回面談は無料です。お気軽にどうぞ。",
                font_size=14, color=RGBColor(0xF5, 0xD0, 0xD5), align=PP_ALIGN.CENTER)

    # Contact items
    contacts = [
        ("LP", "https://mai-san447.github.io/hanasaku-career/"),
        ("無料相談予約（Googleカレンダー）", "https://calendar.app.google/u8SkpMApaMfaRoSx6"),
        ("LINE公式アカウント", "https://lin.ee/qURXYyS"),
    ]
    for i, (label, url) in enumerate(contacts):
        y = 3.8 + i * 0.7
        add_rect(slide, 1.5, y, 7, 0.55, PRIMARY_DARK)
        add_textbox(slide, 1.7, y + 0.05, 2.5, 0.4, label, font_size=11, bold=True, color=WHITE)
        add_textbox(slide, 4.2, y + 0.05, 4.0, 0.4, url, font_size=10, color=RGBColor(0xF5, 0xD0, 0xD5))

    add_textbox(slide, 0.5, 6.5, 9, 0.5,
                "© 2026 ハナサク / EMKO",
                font_size=10, color=RGBColor(0xE8, 0xA0, 0xA8), align=PP_ALIGN.CENTER)

    # Save
    output_dir = os.path.dirname(os.path.abspath(__file__))
    output_path = os.path.join(output_dir, "hanasaku_service_overview.pptx")
    prs.save(output_path)
    print(f"Created: {output_path}")


if __name__ == "__main__":
    create_presentation()
