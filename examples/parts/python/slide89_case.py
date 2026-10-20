import os
from pptx_helpers import *

A = os.path.join(os.path.dirname(__file__), "..", "assets")

def build(prs):
    s = content_slide(prs, "事例紹介", breadcrumb="部品カタログ：カード",
                      lead="case_card（画像・概要・ポイント・出所）")
    top, bot = split_rows(BODY, [1, 1], 14)
    case_card(s, top, "会津若松市：データ活用の基本方針", "全体像を描いたうえで小さく始め、早く展開する方針でデータ活用を進めている。",
              ["市民のニーズを起点にアプリを開発", "オープンデータを積極的に公開"], image_path=os.path.join(A, "photo.png"),
              label="先進事例①", src="会津若松市ホームページ")
    case_card(s, bot, "A県：申請のオンライン化", "全手続の8割をオンライン化し、窓口の待ち時間を半減した。",
              ["共通基盤を県と市町村で共同利用", "職員向けの研修を毎月実施"], label="先進事例②")
