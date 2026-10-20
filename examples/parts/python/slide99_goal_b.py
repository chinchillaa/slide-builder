from pptx_helpers import *


def build(prs):
    s = content_slide(prs, "目標管理部品②", breadcrumb="部品カタログ：目標管理",
                      lead="okr / target_actual")
    l, r = split_cols(BODY, [1, 1], 24)
    okr(s, l, "申請業務を完全にオンライン化する", [("オンライン申請率 90%", 72), ("処理日数を3日以内", 55), ("満足度 4.0 以上", 80)])
    target_actual(s, r, [("東日本", 120, 132), ("中日本", 100, 88), ("西日本", 80, 81)], unit="件")
