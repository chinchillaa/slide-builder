from pptx_helpers import *

def build(prs):
    s = content_slide(prs, "スケジュール・管理部品", breadcrumb="部品カタログ：スケジュール",
                      lead="schedule_list / action_items / kpi_trend")
    l, r = split_cols(BODY, [1, 1.15], 24)
    lt, lb = split_rows(l, [1.3, 1], 14)
    schedule_list(s, lt, [("10/20（月）", "キックオフ", "完了"), ("10/27（月）", "ヒアリング", "進行中"),
                          ("11/10（月）", "中間報告", "未着手"), ("11/25（火）", "最終報告", None)])
    a, b = split_cols(lb, [1, 1], 12)
    kpi_trend(s, a, "120", "月間利用者", [60, 72, 80, 95, 110, 120], unit="人", note="前月比 +9%")
    kpi_trend(s, b, "3.2", "平均処理日数", [5.1, 4.8, 4.0, 3.6, 3.4, 3.2], unit="日")
    action_items(s, r, [("ヒアリング対象者の確定", "鈴木", "10/24", "完了"), ("現行帳票の収集", "高橋", "10/31", "進行中"),
                        ("データ件数の調査", "田中", "10/31", "遅延"), ("予算案の作成", "佐藤", "11/7", "未着手")])
