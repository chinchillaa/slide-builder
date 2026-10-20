from pptx_helpers import *

def build(prs):
    s = content_slide(prs, "2.2 システム構成", breadcrumb="2. 導入方針",
                      lead="既存のTeamsを入口にし、解決しない質問だけを担当者に引き継ぐため、社員の使い方を変えずに導入できる。")
    top, bot = split_rows(BODY, [5, 1], 14)
    l, r = split_cols(top, [1, 1.6], 70)
    zi = zone(s, l, "社内", style="gray")
    emp = (zi[0] + (zi[2] - 110) / 2, zi[1] + 10, 110, 100)
    desk = (zi[0] + (zi[2] - 110) / 2, zi[1] + zi[3] - 110, 110, 100)
    icon_label(s, emp, "pc", "社員", sub="Teamsで質問")
    icon_label(s, desk, "people", "ヘルプデスク担当")
    ci = zone(s, r, "クラウド（SaaS）", style="dashed")
    bot_box = (ci[0] + 20, ci[1] + (ci[3] - 70) / 2, 170, 70)
    message_box(s, bot_box, "AIチャットボット", style="primary", size=14)
    faq = (ci[0] + ci[2] - 130, ci[1] + 10, 110, 100)
    man = (ci[0] + ci[2] - 130, ci[1] + ci[3] - 110, 110, 100)
    icon_label(s, faq, "database", "社内FAQ")
    icon_label(s, man, "document", "業務マニュアル")
    connect(s, emp, bot_box, kind="both", label="質問・回答")
    connect(s, bot_box, desk, kind="dashed", label="未解決のみ引継ぎ")
    connect(s, bot_box, faq, kind="dotted", label="検索")
    connect(s, bot_box, man, kind="dotted", label="検索")
    note_box(s, bot, ["参照するのは社内FAQと業務マニュアルのみ。入力内容はAIの学習に使わない契約とする（3.2 の情報漏えい対策）"],
             label="注記")
