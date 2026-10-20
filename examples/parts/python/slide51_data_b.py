from pptx_helpers import *

def build(prs):
    s = content_slide(prs, "数値部品②", breadcrumb="部品カタログ：数値",
                      lead="progress_bars / rank_bars / percent_bar / donut_kpi")
    (a, b), (c, d) = [split_cols(r, [1, 1], 30) for r in split_rows(BODY, [1.2, 1], 30)]
    progress_bars(s, a, [("要件定義", 100), ("設計", 80), ("開発", 45), ("テスト", 10)], target=60)
    rank_bars(s, b, [("経理部", 18), ("営業部", 12), ("総務部", 6), ("人事部", 4)], unit="h", highlight=0)
    percent_bar(s, c, [("人件費", 45), ("システム費", 30), ("外注費", 15), ("その他", 10)])
    l2, r2 = split_cols(d, [1, 1], 10)
    donut_kpi(s, l2, 72, "移行完了率")
    donut_kpi(s, r2, 35, "教育受講率")
