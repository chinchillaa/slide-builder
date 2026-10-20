from pptx_helpers import *


def build(prs):
    s = content_slide(prs, "費用・効果部品②", breadcrumb="部品カタログ：費用・効果",
                      lead="budget_actual_table（予実対比） / effect_breakdown（定量・定性効果）")
    top, bot = split_rows(BODY, [1, 1.15], 16)
    budget_actual_table(s, top, [("人件費", 1200, 1150), ("外注費", 800, 920), ("システム費", 500, 480)], unit="万円")
    effect_breakdown(s, bot, [("集計作業の削減", 480), ("転記ミス対応の削減", 120), ("紙・郵送費の削減", 60)],
                     ["意思決定が早くなる", "担当者の残業が減る", "データの信頼性が上がる"])
