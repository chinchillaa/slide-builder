from pptx_helpers import *

def build(prs):
    s = content_slide(prs, "表部品⑤", breadcrumb="部品カタログ：表",
                      lead="mark_matrix（列グループ・凡例付き） / bar_table（行内の横棒）")
    top, bot = split_rows(BODY, [1, 1], 14)
    mark_matrix(s, top, ["UI", "運営", "機器", "サービス"], ["A社", "B社", "C社", "D社", "E社"],
                [["★", "✓", "", "✓", ""], ["★", "★", "✓", "", "✓"], ["", "✓", "★", "", ""], ["✓", "", "", "★", "★"]],
                groups=[("米国", 2), ("中国", 2), ("欧州", 1)], legend={"★": "注力", "✓": "参入"})
    bar_table(s, bot, ["国・地域", "特徴", "充電器の台数（千台）"],
              [["米国", "民間主導で整備が進む", 160], ["中国", "国策で急拡大", 1800], ["欧州", "域内で規格を統一", 450]],
              bar_col=2, col_ratios=[1, 2.4, 2])
