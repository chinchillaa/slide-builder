from pptx_helpers import *

def build(prs):
    s = content_slide(prs, "構造部品③", breadcrumb="部品カタログ：構造",
                      lead="matrix_2x2 / positioning_map")
    l, r = split_cols(BODY, [1, 1], 30)
    matrix_2x2(s, l, ("実現性", "低", "高"), ("効果", "低", "高"),
               [("検討", "効果は高いが難しい"), ("最優先", "すぐに着手"), ("見送り", ""), ("改善", "小さく実施")], highlight=1)
    positioning_map(s, r, ("価格", "低価格", "高価格"), ("機能", "シンプル", "多機能"),
                    [(0.25, 0.3, "A社"), (0.7, 0.8, "B社"), (0.6, 0.35, "C社"), (0.35, 0.72, "自社（目標）", True)])
