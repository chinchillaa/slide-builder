from pptx_helpers import *


def build(prs):
    cover_band(prs, "令和8年度 業務システム刷新に向けた\n現状調査および方針検討", subtitle="検討報告書",
               date_text="2026年10月20日", org="情報システム部")
    section_split(prs, "2. 現状分析", ["2.1 課題と対応方針", "2.2 定量分析", "2.3 ヒアリング結果"])
