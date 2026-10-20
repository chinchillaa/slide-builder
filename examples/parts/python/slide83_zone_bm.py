from pptx_helpers import *

def build(prs):
    s = content_slide(prs, "図解部品①", breadcrumb="部品カタログ：図解",
                      lead="business_model（お金・サービス・モノ・情報の流れ） / zone")
    business_model(s, BODY,
                   {"u": ("利用者", 0, 0, "person"), "op": ("運営事業者", 1, 0, "building"),
                    "mk": ("機器メーカー", 2, 0, "factory"), "pf": ("決済事業者", 1, 1, "money"),
                    "app": ("アプリ", 0, 1, "phone")},
                   [("op", "u", "service", "充電サービス"), ("u", "op", "money", "利用料金"),
                    ("mk", "op", "goods", "充電器"), ("op", "mk", "money", "代金"), ("u", "app", "info", "検索"),
                    ("pf", "op", "money", "精算")],
                   columns=["ユーザー", "事業者", "サプライヤー"])
