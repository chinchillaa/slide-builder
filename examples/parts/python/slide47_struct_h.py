from pptx_helpers import *

def build(prs):
    s = content_slide(prs, "構造部品⑧", breadcrumb="部品カタログ：構造",
                      lead="persona_card / role_cards")
    top, bot = split_rows(BODY, [1, 1], 14)
    persona_card(s, top, "経理部 主任（40代）", [("業務", "月次決算"), ("ITスキル", "Excel 中級"), ("課題", "締め作業が多い")],
                 ["集計の自動化", "ミスの削減"])
    role_cards(s, bot, [("PM", "佐藤", ["全体の進捗管理", "経営への報告"]), ("業務リーダー", "鈴木", ["業務要件の取りまとめ"]),
                        ("技術リーダー", "高橋", ["方式設計", "ベンダー管理"])])
