from pptx_helpers import *

def build(prs):
    s = content_slide(prs, "構造部品⑦", breadcrumb="部品カタログ：構造",
                      lead="swot / relation_map")
    l, r = split_cols(BODY, [1, 1.1], 24)
    swot(s, l, ["顧客基盤が厚い"], ["IT 人材が不足"], ["クラウドの普及"], ["競合の参入"])
    relation_map(s, r, {"a": ("自社", 0.5, 0.0, "label"), "b": ("販売代理店", 0.0, 0.5), "c": ("顧客", 0.5, 1.0, "pale"),
                        "d": ("システムベンダー", 1.0, 0.5)},
                 [("a", "b", "商品を卸す"), ("b", "c", "販売"), ("d", "a", "保守"), ("c", "a", "問合せ")])
