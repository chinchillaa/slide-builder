from pptx_helpers import *

def build(prs):
    s = content_slide(prs, "表部品③", breadcrumb="部品カタログ：表",
                      lead="raci_table / key_value_table")
    l, r = split_cols(BODY, [1.2, 1], 24)
    raci_table(s, l, ["要件定義", "ベンダー選定", "開発", "受入テスト", "展開"], ["情シス", "業務部門", "ベンダー", "経営"],
               [["R", "C", "I", "A"], ["R", "C", "", "A"], ["A", "I", "R", ""], ["C", "R", "C", "A"], ["R", "R", "C", "I"]])
    key_value_table(s, r, [("案件名", "業務システム刷新"), ("期間", "2026年9月〜2027年3月"),
                           ("体制", ["PM 1名", "メンバー 3名"]), ("予算", "2,000万円（税抜）")], key_ratio=0.28)
