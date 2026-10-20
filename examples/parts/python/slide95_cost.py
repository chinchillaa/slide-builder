from pptx_helpers import *


def build(prs):
    s = content_slide(prs, "費用・効果部品①", breadcrumb="部品カタログ：費用・効果",
                      lead="cost_table（小計・合計） / roi_chart（投資回収）")
    l, r = split_cols(BODY, [1, 1], 24)
    cost_table(s, l, ["費目", "2026年度", "2027年度"],
               [("ライセンス", 300, 300), ("構築", 1200, 200), ("小計：初期・運用", 1500, 500), ("教育", 100, 50),
                ("保守", 0, 150)], unit="万円")
    roi_chart(s, r, ["1年目", "2年目", "3年目", "4年目", "5年目"], [1500, 500, 300, 300, 300], [200, 800, 900, 900, 900],
              unit="万円")
