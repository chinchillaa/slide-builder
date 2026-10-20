from pptx_helpers import *

def build(prs):
    s = content_slide(prs, "グラフ部品①", breadcrumb="部品カタログ：グラフ",
                      lead="waterfall / scatter_chart / bubble_chart")
    a, b, c = split_cols(BODY, [1.3, 1, 1], 20)
    waterfall(s, a, ("2025年度", 100), [("値上げ", 20), ("数量減", -15), ("新商品", 12), ("為替", -7)],
              end_label="2026年度", title="売上の増減要因", unit="億円")
    scatter_chart(s, b, {"店舗": [(10, 20), (15, 28), (22, 35), (30, 41), (35, 39), (42, 55)]},
                  x_title="広告費", y_title="売上")
    bubble_chart(s, c, {"事業A": [(10, 60, 30)], "事業B": [(40, 30, 50)], "事業C": [(70, 75, 20)]},
                 x_title="成長率", y_title="利益率")
