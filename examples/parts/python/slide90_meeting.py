from pptx_helpers import *


def build(prs):
    s = content_slide(prs, "会議・報告部品①", breadcrumb="部品カタログ：会議・報告",
                      lead="traffic_light / decision_log / meeting_agenda")
    l, r = split_cols(BODY, [1.2, 1], 24)
    traffic_light(s, l, [("スケジュール", "良好", "計画どおり進行"), ("品質", "注意", "テスト不具合が想定より多い"),
                         ("コスト", "良好", "予算内"), ("体制", "問題", "要員が1名不足。来週までに補充")])
    rt, rb = split_rows(r, [1, 1], 14)
    decision_log(s, rt, [("決定", "SaaS 型を採用する", "", ""), ("宿題", "見積りを再取得", "高橋", "10/31"),
                         ("確認", "データ移行の範囲", "鈴木", "11/7")])
    meeting_agenda(s, rb, [("10:00-10:10", "前回の振り返り", "PMO"), ("10:10-10:40", "要件の確認", "業務チーム"),
                           ("10:40-11:00", "今後の進め方", "PM")])
