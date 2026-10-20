from pptx_helpers import *

def build(prs):
    s = content_slide(prs, "テキスト部品②", breadcrumb="部品カタログ：テキスト",
                      lead="statement / qa_list / definition_list / checklist / ref_tag")
    ref_tag(s)
    top, bottom = split_rows(BODY, [1, 2.2], 16)
    statement(s, top, "データに基づく意思決定を、全社の標準にする", sub="2027年度 情報システム部 方針")
    a, b, c = split_cols(bottom, [1.2, 1.2, 1], 20)
    qa_list(s, a, [("移行中も旧システムは使えるか", "並行稼働期間を3か月設ける"),
                   ("費用はどの部署が負担するか", "全社共通費として計上する")])
    definition_list(s, b, [("To-Be", "あるべき業務の姿"), ("BPR", "業務プロセスの抜本的な見直し"),
                           ("PoC", "概念実証。小規模な試行")], term_ratio=0.3)
    checklist(s, c, ["要件定義書の承認", "予算の確保", "体制の確定", "ベンダー選定"], checked=[True, True, False, False])
