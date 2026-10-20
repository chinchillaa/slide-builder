from pptx_helpers import *


def build(prs):
    s = content_slide(prs, "顧客・利用者部品②", breadcrumb="部品カタログ：顧客・利用者",
                      lead="likert_chart（5段階評価） / empathy_map")
    l, r = split_cols(BODY, [1.2, 1], 24)
    likert_chart(s, l, ["操作は分かりやすい", "処理が速い", "また使いたい"], [[30, 40, 15, 10, 5], [20, 35, 25, 15, 5], [45, 35, 10, 5, 5]])
    empathy_map(s, r, "経理担当\n40代", ["月末は忙しい"], ["ミスが怖い"], ["Excelで集計"], ["もっと楽にしたい"])
