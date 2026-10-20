from pptx_helpers import *

def build(prs):
    s = content_slide(prs, "テキスト部品①", breadcrumb="部品カタログ：テキスト",
                      lead="numbered_list / callout / note_box / quote_box / tag")
    l, r = split_cols(BODY, [1, 1], 24)
    numbered_list(s, (l[0], l[1], l[2], 200), [("現状把握", "業務ヒアリングと資料調査を行う"),
                                                ("課題整理", "課題を分類し、優先度を付ける"),
                                                ("方針策定", "To-Be の方向性を定める")])
    callout(s, (l[0] + 40, l[1] + 230, 260, 50), "ここが最重要ポイント", tail=(l[0] + 20, l[1] + 300))
    note_box(s, (r[0], r[1], r[2], 120), ["対象は本社部門のみ", "子会社は次年度に実施"])
    quote_box(s, (r[0], r[1] + 136, r[2], 110), "月末は集計作業に追われ、分析まで手が回らない", speaker="経理部 担当者")
    x = r[0]
    for st in ["primary", "mid", "pale", "outline", "gray", "alert"]:
        x += tag(s, x, r[1] + 270, st) + 8
    legend(s, r[0], r[1] + 310, [("自社", COLOR_PRIMARY), ("委託先", COLOR_PRIMARY_MID), ("顧客", COLOR_GRAY)])
    number_badge(s, r[0], r[1] + 340, 1); number_badge(s, r[0] + 30, r[1] + 340, 2, style="mid")
    number_badge(s, r[0] + 60, r[1] + 340, "!", style="alert")
