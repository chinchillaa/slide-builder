from pptx_helpers import *

def build(prs):
    s = content_slide(prs, "スケジュール部品", breadcrumb="部品カタログ：スケジュール",
                      lead="calendar_month（月間予定）")
    calendar_month(s, BODY, 2026, 11, {4: "キックオフ", 12: "ヒアリング①", 13: "ヒアリング②", 20: "中間報告", 27: "最終報告"})
