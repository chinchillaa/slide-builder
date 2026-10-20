from pptx_helpers import *

def build(prs):
    s = content_slide(prs, "フロー部品③", breadcrumb="部品カタログ：フロー",
                      lead="cycle / funnel / pyramid")
    a, b, c = split_cols(BODY, [1, 1, 1.4], 24)
    cycle(s, a, ["Plan", "Do", "Check", "Act"], center="継続的改善")
    funnel(s, b, [("認知", "10,000"), ("興味", "3,000"), ("検討", "800"), ("契約", "120")])
    pyramid(s, c, [("経営", "全社KPIで意思決定"), ("管理", "部門別に分析"), ("現場", "日次のデータ入力")])
