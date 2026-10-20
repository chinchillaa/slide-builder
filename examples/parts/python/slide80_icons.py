from pptx_helpers import *

def build(prs):
    s = content_slide(prs, "アイコン部品", breadcrumb="部品カタログ：アイコン",
                      lead="icon（18種） / icon_label / icon_steps")
    top, bot = split_rows(BODY, [1.15, 1], 16)
    cells = grid_areas(top, 2, 9, 10)
    for k, a in zip(ICONS, cells):
        icon_label(s, a, k, k)
    icon_steps(s, bot, [("document", "申請", "Webで入力"), ("person", "確認", "担当者がチェック"),
                        ("server", "登録", "システムに反映"), ("mail", "通知", "結果をメール")])
