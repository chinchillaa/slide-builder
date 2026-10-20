from pptx_helpers import *


def build(prs):
    s = content_slide(prs, "フレームワーク部品①", breadcrumb="部品カタログ：フレームワーク",
                      lead="framework_grid（3c / pest / kpt / 4p / 4c）")
    top, bot = split_rows(BODY, [1, 1], 14)
    a, b = split_cols(top, [1.3, 1], 20)
    framework_grid(s, a, "3c", [["高齢化が進む"], ["大手が参入"], ["地域の顧客基盤"]])
    framework_grid(s, b, "pest", [["規制緩和"], ["景気回復"], ["人口減少"], ["AIの普及"]])
    c, d = split_cols(bot, [1.3, 1], 20)
    framework_grid(s, c, "kpt", [["朝会の継続"], ["レビュー待ちが長い"], ["レビュー担当を2名に"]])
    framework_grid(s, d, "4p", [["高品質"], ["中価格帯"], ["Web直販"], ["SNS広告"]])
