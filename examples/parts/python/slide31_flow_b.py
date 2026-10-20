from pptx_helpers import *

def build(prs):
    s = content_slide(prs, "フロー部品②", breadcrumb="部品カタログ：フロー",
                      lead="swimlane（担当別の業務フロー）")
    swimlane(s, BODY, ["申請者", "上長", "経理部"], 5,
             steps=[(0, 0, "申請書を作成"), (0, 1, "申請"), (1, 2, "内容を確認"), (1, 3, "承認"), (2, 4, "支払処理"),
                    (0, 3, "差戻し対応")],
             arrows=[(0, 1), (1, 2), (2, 3), (3, 4), (2, 5)],
             col_headers=["準備", "申請", "確認", "承認", "処理"])
