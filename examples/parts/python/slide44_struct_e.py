from pptx_helpers import *

def build(prs):
    s = content_slide(prs, "構造部品⑤", breadcrumb="部品カタログ：構造",
                      lead="layer_stack / equation / versus")
    top, mid, bot = split_rows(BODY, [1.4, 0.55, 1], 14)
    layer_stack(s, top, [("プレゼンテーション", ["ダッシュボード", "モバイルアプリ", "帳票"]),
                         ("アプリケーション", ["販売管理", "在庫管理", "会計"]),
                         ("データ", ["データウェアハウス", "マスタ管理"])])
    equation(s, mid, ["業務の標準化", "システム化"], "工数 40% 削減")
    versus(s, bot, ("自社開発", ["自由度が高い", "費用が高い"]), ("SaaS 利用", ["導入が早い", "制約がある"]))
