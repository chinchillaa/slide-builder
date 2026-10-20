from pptx_helpers import *

def build(prs):
    s = content_slide(prs, "構造部品⑥", breadcrumb="部品カタログ：構造",
                      lead="gap_analysis / brace_group")
    top, bot = split_rows(BODY, [1.2, 1], 16)
    gap_analysis(s, top, ["紙とExcelで管理", "集計に3日"], ["システムで一元管理", "集計は即時"],
                 ["業務の標準化", "システム導入", "教育"])
    brace_group(s, bot, ["売上データの統合", "在庫データの統合", "顧客データの統合"], "全社データ基盤の構築")
