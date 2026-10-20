from pptx_helpers import *

def build(prs):
    s = content_slide(prs, "表部品①", breadcrumb="部品カタログ：表",
                      lead="comparison_table（推奨列を強調） / score_table（重み付け・合計）")
    top, bot = split_rows(BODY, [1, 1.1], 14)
    comparison_table(s, top, ["初期費用", "導入期間", "拡張性"], ["A案：SaaS", "B案：パッケージ", "C案：スクラッチ"],
                     [["◎ 500万円", "○ 1,500万円", "× 5,000万円"], ["◎ 3か月", "○ 6か月", "△ 12か月"],
                      ["△ 制約あり", "○ アドオン可", "◎ 自由"]], recommend=0)
    score_table(s, bot, ["費用", "期間", "機能適合", "運用負荷"], ["A案", "B案", "C案"],
                [[5, 3, 1], [5, 4, 2], [3, 4, 5], [4, 3, 2]], weights=[2, 1, 2, 1])
