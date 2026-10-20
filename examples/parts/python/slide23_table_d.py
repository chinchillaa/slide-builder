from pptx_helpers import *

def build(prs):
    s = content_slide(prs, "表部品④", breadcrumb="部品カタログ：表",
                      lead="heatmap_table / status_table / check_table / highlight_table")
    (a, b), (c, d) = [split_cols(r, [1, 1], 20) for r in split_rows(BODY, [1, 1], 14)]
    heatmap_table(s, a, ["東日本", "中日本", "西日本"], ["4月", "5月", "6月", "7月"],
                  [[120, 135, 160, 180], [80, 95, 90, 110], [60, 70, 85, 140]])
    status_table(s, b, ["タスク", "担当", "状況"], [["要件定義", "情シス", "完了"], ["ベンダー選定", "情シス", "進行中"],
                 ["データ移行", "業務部門", "遅延"], ["教育", "人事", "未着手"]], status_col=2, col_ratios=[2, 1.2, 1])
    check_table(s, c, ["申請", "承認", "集計"], ["本社", "支社", "子会社"],
                [[True, True, False], [True, False, False], [True, True, "一部"]])
    highlight_table(s, d, ["順位", "部署", "削減時間"], [["1", "経理部", "18h"], ["2", "営業部", "12h"], ["3", "総務部", "6h"]],
                    highlight_rows=(0,), col_ratios=[0.6, 2, 1.2], align=["c", "l", "r"])
