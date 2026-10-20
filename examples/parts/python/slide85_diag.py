from pptx_helpers import *

def build(prs):
    s = content_slide(prs, "図解部品③", breadcrumb="部品カタログ：図解",
                      lead="flow_before_after / issue_map")
    l, r = split_cols(BODY, [1.2, 1], 24)
    flow_before_after(s, l, ["紙の申請書を郵送", "窓口で受付", "手入力で登録", "郵送で通知"],
                      ["Webで申請", "自動で受付", "システムに自動登録", "メールで通知"], changed=(0, 2, 3))
    issue_map(s, r, ["システムの価値・メリット", "データの標準化", "ベネフィットとリスク", "医療機器との関係"], current=1)
