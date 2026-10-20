from pptx_helpers import *

def build(prs):
    s = content_slide(prs, "フロー部品①", breadcrumb="部品カタログ：フロー",
                      lead="chevron_row / flow_boxes / vertical_steps / ipo")
    x, y, w, h = BODY
    chevron_row(s, (x, y, w, 40), ["現状調査", "課題整理", "方針策定", "計画策定"], active=1)
    flow_boxes(s, (x, y + 54, w, 60), [("申請", "Webフォーム"), ("一次承認", "課長"), ("二次承認", "部長"), ("処理", "経理部")])
    l, r = split_cols((x, y + 130, w, h - 130), [1, 1.6], 24)
    vertical_steps(s, l, [("キックオフ", "体制と進め方の合意"), ("ヒアリング", "全8部署"), ("報告", "経営会議")])
    ipo(s, r, ["売上データ", "在庫データ"], ["集計・突合", "異常値の検出"], ["月次レポート", "アラート"])
