from pptx_helpers import *

def build(prs):
    s = content_slide(prs, "フロー部品④", breadcrumb="部品カタログ：フロー",
                      lead="stair_steps / timeline")
    top, bot = split_rows(BODY, [1, 1], 16)
    stair_steps(s, top, [("Lv1 可視化", "データを集める"), ("Lv2 分析", "要因を探る"), ("Lv3 予測", "将来を見通す"),
                         ("Lv4 最適化", "自動で判断する")])
    timeline(s, bot, [("2026/4", "構想策定", ""), ("2026/9", "要件定義", "業務部門と合意"), ("2027/1", "開発着手", ""),
                      ("2027/10", "稼働", "全社展開")])
