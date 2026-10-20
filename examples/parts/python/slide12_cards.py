from pptx_helpers import *

def build(prs):
    s = content_slide(prs, "カード部品①", breadcrumb="部品カタログ：カード",
                      lead="card_grid（titled_box の格子） / band_heading / sub_heading / divider")
    x, y, w, h = BODY
    used = band_heading(s, x, y, w, "検討の3つの観点")
    card_grid(s, (x, y + used, w, 150), [("業務", ["手作業を減らす", "標準化"]), ("システム", ["クラウド化", "データ連携"]),
                                         ("組織", ["役割の明確化", "教育"])], cols=3)
    used2 = sub_heading(s, x, y + used + 166, w / 2 - 10, "補足事項")
    bullets(s, (x, y + used + 166 + used2, w / 2 - 10, 120), ["移行は段階的に行う", "既存データはすべて移行する"])
    divider(s, x + w / 2 + 10, y + used + 180, w / 2 - 10, dash=True)
    card_grid(s, (x + w / 2 + 10, y + used + 196, w / 2 - 10, 140),
              [("短期", "既存運用の改善"), ("中長期", "新システムへの移行")], cols=2, style="pale")
