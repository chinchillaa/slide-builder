from pptx_helpers import *

def build(prs):
    s = content_slide(prs, "グラフ部品②", breadcrumb="部品カタログ：グラフ",
                      lead="chart_with_insight / chart の追加種類（line・area・column_100・radar）")
    top, bot = split_rows(BODY, [1.1, 1], 14)
    chart_with_insight(s, top, "line", ["4月", "5月", "6月", "7月", "8月"], {"今年": [80, 92, 88, 105, 120], "昨年": [75, 78, 82, 85, 90]},
                       ["7月以降、前年を大きく上回る", "新商品の効果が出始めた"], title="月別売上", unit="百万円")
    a, b, c = split_cols(bot, [1, 1, 1], 16)
    chart(s, a, "area", ["Q1", "Q2", "Q3", "Q4"], {"利用者数": [120, 180, 260, 400]}, labels=False)
    chart(s, b, "column_100", ["2024", "2025", "2026"], {"紙": [60, 40, 15], "電子": [40, 60, 85]}, number_format="0")
    chart(s, c, "radar", ["費用", "品質", "速度", "拡張性", "保守性"], {"A案": [4, 3, 5, 2, 3], "B案": [3, 4, 3, 4, 4]}, labels=False)
