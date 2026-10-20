from pptx_helpers import *

def build(prs):
    s = content_slide(prs, "番号注釈", breadcrumb="部品カタログ：線・矢印",
                      lead="annotate_points（任意の図に番号を打ち、欄外で説明する）")
    l, r = split_cols(BODY, [1.5, 1], 24)
    layer_stack(s, l, [("画面", ["申請画面", "承認画面"]), ("業務", ["審査", "通知"]), ("データ", ["申請DB"])])
    annotate_points(s, [(l[0] + 250, l[1] + 30), (l[0] + 470, l[1] + 160), (l[0] + 330, l[1] + 290)],
                    ["入力項目を半分に削減", "通知を自動化", "データを一元管理"], r)
