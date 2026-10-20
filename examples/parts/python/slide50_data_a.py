from pptx_helpers import *

def build(prs):
    s = content_slide(prs, "数値部品①", breadcrumb="部品カタログ：数値",
                      lead="kpi_row / stat_compare / big_number")
    top, mid, bot = split_rows(BODY, [1, 1, 0.9], 14)
    kpi_row(s, top, [("40", "工数削減", "h/月"), ("98", "満足度", "%"), ("0", "転記ミス", "件/月"), ("3", "導入期間", "か月")])
    stat_compare(s, mid, ("導入前", "72h"), ("導入後", "32h"), delta="▲56%")
    big_number(s, bot, "3.2", unit="億円", desc="5年間の累計削減効果（試算）", label="期待効果")
