from pptx_helpers import *

def build(prs):
    s = content_slide(prs, "図解部品②", breadcrumb="部品カタログ：図解",
                      lead="zone（システム概念図の領域分け） / section_tab")
    body = section_tab(s, "4.1 システムの基本構成")
    l, r = split_cols(body, [1, 2], 30)
    inner = zone(s, l, "利用者", style="gray")
    icon_label(s, (inner[0] + 20, inner[1] + 10, 80, 90), "people", "職員")
    icon_label(s, (inner[0] + 130, inner[1] + 10, 80, 90), "phone", "住民", sub="スマホ")
    inner2 = zone(s, r, "共通基盤（本事業の範囲）", style="dashed")
    a = (inner2[0] + 20, inner2[1] + 20, 150, 60); b = (inner2[0] + 330, inner2[1] + 20, 150, 60)
    message_box(s, a, "申請受付", style="outline", size=13); message_box(s, b, "審査・決裁", style="outline", size=13)
    connect(s, a, b, label="申請データ")
    icon_label(s, (inner2[0] + 240, inner2[1] + 130, 80, 100), "database", "共通DB")
    connect(s, b, (inner2[0] + 240, inner2[1] + 130, 80, 100), kind="dotted")
    icon(s, inner2[0] + 420, inner2[1] + 140, 44, "lock", COLOR_PRIMARY_MID)   # icon は色も指定できる
    text(s, (inner2[0] + 380, inner2[1] + 190, 124, 24), "暗号化", size=12, bold=True, align="center")
