from pptx_helpers import *

def build(prs):
    s = content_slide(prs, "カード部品②", breadcrumb="部品カタログ：カード",
                      lead="point_cards / pros_cons / conclusion_bar")
    top, mid, bot = split_rows(BODY, [1.1, 1, 0.32], 12)
    point_cards(s, top, [("意思決定の迅速化", "日次でKPIを確認できる"), ("工数削減", "集計作業を月40h削減"),
                         ("品質向上", "転記ミスをゼロにする")])
    pros_cons(s, mid, ["初期費用が低い", "短期間で導入できる"], ["カスタマイズに制約がある", "月額費用が継続する"])
    conclusion_bar(s, bot, "SaaS 型の採用を推奨する")
