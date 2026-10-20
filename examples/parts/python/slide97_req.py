from pptx_helpers import *


def build(prs):
    s = content_slide(prs, "要件・仕様部品", breadcrumb="部品カタログ：要件・仕様",
                      lead="requirements_table（MoSCoW） / grouped_table（1列目を結合）")
    top, bot = split_rows(BODY, [1, 1.1], 16)
    requirements_table(s, top, [("F-01", "申請", "Webから申請できる", "Must"), ("F-02", "申請", "下書きを保存できる", "Should"),
                                ("F-03", "通知", "スマホに通知できる", "Could"), ("F-04", "連携", "旧システムとの連携", "Won't")])
    grouped_table(s, bot, ["分類", "機能", "概要"], [("申請", "申請入力", "画面から入力する"), ("申請", "添付", "ファイルを添付する"),
                                                    ("審査", "審査", "担当者が内容を確認する"), ("審査", "差戻し", "修正を依頼する"),
                                                    ("通知", "結果通知", "メールで通知する")], col_ratios=[1, 1.4, 3])
