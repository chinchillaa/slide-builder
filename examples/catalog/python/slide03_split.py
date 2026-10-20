from pptx_helpers import *

def build(prs):
    s = content_slide(prs, "1.1. 背景・目的", breadcrumb="1. 本検討の概要",
                      lead="現行システムの老朽化と業務の属人化を背景に、**刷新方針の策定**を目的として現状調査を実施した。")
    areas, _ = split_columns(s, BODY, [1, 1], titles=["背景", "本検討の目的"], arrow=True)
    bullets(s, areas[0], [
        "現行システムは稼働から15年が経過し、保守期限が2028年3月に到来する",
        "業務手順が担当者ごとに異なり、引継ぎに時間がかかっている",
        (1, "月次処理の約4割が手作業"),
        "データが部署ごとに分散し、横断的な分析ができない",
    ], box=True)
    message_box(s, (areas[1][0], areas[1][1], areas[1][2], 80), "刷新方針（To-Be）の策定")
    message_box(s, (areas[1][0], areas[1][1] + 92, areas[1][2], 80), "移行ロードマップの策定", style="outline")
    message_box(s, (areas[1][0], areas[1][1] + 184, areas[1][2], 80), "概算費用と体制の見積り", style="pale")
    source(s, "情報システム部 現行システム資産台帳（2026年4月時点）")
