from pptx_helpers import *

def build(prs):
    s = content_slide(prs, "フロー部品⑦", breadcrumb="部品カタログ：フロー",
                      lead="value_chain / step_matrix")
    top, bot = split_rows(BODY, [0.9, 1.2], 16)
    value_chain(s, top, ["調達", "製造", "物流", "販売", "サービス"], support=["全般管理（経営企画・経理）", "人事・労務管理"])
    step_matrix(s, bot, ["申請", "審査", "承認", "通知"], ["現状", "課題"],
                [["紙で提出", "目視で確認", "押印", "郵送"], ["記入ミスが多い", "時間がかかる", "出社が必要", "コストが高い"]])
