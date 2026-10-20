from pptx_helpers import *

def build(prs):
    s = content_slide(prs, "フロー部品⑥", breadcrumb="部品カタログ：フロー",
                      lead="converge / diverge / phase_gate")
    top, bot = split_rows(BODY, [1, 1], 16)
    l, r = split_cols(top, [1, 1], 30)
    converge(s, l, ["販売データ", "在庫データ", "顧客データ"], "統合データ基盤")
    diverge(s, r, "統合データ基盤", ["経営ダッシュボード", "需要予測", "顧客分析"])
    phase_gate(s, bot, [("構想", ["目的の明確化"]), ("要件定義", ["業務要件", "機能要件"]), ("開発", ["設計・実装"]),
                        ("展開", ["教育", "移行"])], ["投資判断", "要件承認", "稼働判定"])
