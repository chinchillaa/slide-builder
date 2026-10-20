from datetime import date
from pptx_helpers import *

def build(prs):
    s = content_slide(prs, "3.1 導入スケジュール", breadcrumb="3. 進め方とご判断事項",
                      lead="2027年1月にPoCを始め、4月の判定で基準を満たせば7月に全社展開する。")
    gantt(s, BODY, date(2026, 11, 1), date(2027, 9, 30),
          rows=["PoC準備", "PoC", "評価", "本番構築", "全社展開"],
          bars=[(0, date(2026, 11, 1), date(2026, 12, 25), "FAQ整備・契約"),
                (1, date(2027, 1, 5), date(2027, 3, 31), "情シス部門で試行", "mid"),
                (2, date(2027, 4, 1), date(2027, 4, 30), "効果測定", "pale"),
                (3, date(2027, 5, 6), date(2027, 6, 30), "全社FAQ連携"),
                (4, date(2027, 7, 1), date(2027, 9, 30), "展開・定着支援", "mid")],
          milestones=[(date(2027, 4, 30), "PoC判定"), (date(2027, 7, 1), "全社展開")])
    source(s, "注：期間は2026年11月〜2027年9月。PoC判定の基準：自動解決率40%以上かつ誤回答率5%未満")
