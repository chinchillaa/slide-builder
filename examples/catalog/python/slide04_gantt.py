from datetime import date
from pptx_helpers import *

def build(prs):
    s = content_slide(prs, "1.2. 実施スケジュール", breadcrumb="1. 本検討の概要",
                      lead="2026年9月から12月までの4か月で、現状調査・方針策定・ロードマップ策定を行う。")
    gantt(s, BODY, date(2026, 9, 1), date(2026, 12, 31),
          rows=["現状調査", "方針策定", "ロードマップ策定"],
          bars=[(0, date(2026, 9, 1), date(2026, 9, 30), "業務ヒアリング"),
                (0, date(2026, 9, 14), date(2026, 10, 16), "システム棚卸", "pale", 1),
                (1, date(2026, 10, 12), date(2026, 11, 20), "To-Be業務・機能の検討", "mid"),
                (2, date(2026, 11, 16), date(2026, 12, 25), "移行計画・概算費用")],
          milestones=[(date(2026, 10, 30), "中間報告"), (date(2026, 12, 25), "最終報告")])
