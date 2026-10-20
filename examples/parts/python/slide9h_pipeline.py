from pptx_helpers import *


def build(prs):
    s = content_slide(prs, "図解部品④", breadcrumb="部品カタログ：図解",
                      lead="pipeline（データの流れ）")
    top, bot = split_rows(BODY, [1, 1], 20)
    pipeline(s, top, [("pc", "各拠点の端末"), ("server", "収集サーバ"), ("database", "データ基盤"), ("chart", "ダッシュボード")],
             edge_labels=["日次で送信", "整形", "集計"])
    pipeline(s, bot, [("phone", "住民のスマホ"), ("cloud", "クラウド"), ("building", "市役所")], edge_labels=["申請", "連携"])
