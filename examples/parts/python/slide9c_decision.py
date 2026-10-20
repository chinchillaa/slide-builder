from pptx_helpers import *


def build(prs):
    s = content_slide(prs, "意思決定部品①", breadcrumb="部品カタログ：意思決定",
                      lead="recommendation（推奨案と理由・次のアクション）")
    recommendation(s, BODY, [("A案：SaaS", "短期間・低コストで導入"), ("B案：パッケージ", "カスタマイズしやすい"),
                             ("C案：自社開発", "自由度は高いが高コスト")], 0,
                   ["費用が最も低い", "保守期限までに導入できる"], ["役員会で承認を得る（11月）", "ベンダーと契約する（12月）"])
