from pptx_helpers import *

def build(prs):
    s = content_slide(prs, "表部品②", breadcrumb="部品カタログ：表",
                      lead="matrix_table / harvey_table")
    l, r = split_cols(BODY, [1.3, 1], 24)
    matrix_table(s, l, ["業務", "システム"], ["短期（〜1年）", "中長期（〜3年）"],
                 [[["手順書の整備"], ["業務の標準化"]], [["既存システムの延命"], ["クラウドへ移行"]]], corner="領域＼期間")
    harvey_table(s, r, ["売上管理", "在庫管理", "会計連携", "分析機能"], ["A社", "B社", "C社"],
                 [[4, 3, 2], [3, 4, 1], [2, 2, 4], [1, 3, 3]])
