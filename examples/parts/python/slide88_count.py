from pptx_helpers import *

def build(prs):
    s = content_slide(prs, "件数の分類・絞り込み", breadcrumb="部品カタログ：構造",
                      lead="count_tree（件数付きの分類ツリー） / reduction_bars（絞り込みの過程）")
    l, r = split_cols(BODY, [1.4, 1], 24)
    count_tree(s, l, ("対象手続", 336, [("申請", 195, [("個人", 120, []), ("法人", 75, [])]),
                                       ("届出", 104, []), ("その他", 37, [])]))
    reduction_bars(s, r, [("全手続", 407, "件数10万件未満"), ("候補", 293, "次年度に着手可能"), ("対象", 3, None)])
