from pptx_helpers import *


def build(prs):
    s = content_slide(prs, "意思決定部品②", breadcrumb="部品カタログ：意思決定",
                      lead="assumptions / plan_cards")
    top, bot = split_rows(BODY, [0.8, 1.2], 16)
    assumptions(s, top, ["対象は本社部門", "既存データはすべて移行"], ["予算は2,000万円以内", "2028年3月までに稼働"])
    plan_cards(s, bot, [("ライト", "月額 5万円", ["利用者 50名まで", "メールサポート"]),
                        ("スタンダード", "月額 12万円", ["利用者 200名まで", "電話サポート", "データ連携"]),
                        ("エンタープライズ", "個別見積り", ["利用者 無制限", "専任担当"])], recommended=1)
