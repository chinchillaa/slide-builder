from pptx_helpers import *

def build(prs):
    s = content_slide(prs, "カード部品③", breadcrumb="部品カタログ：カード",
                      lead="before_after / icon_rows / summary_box / exec_points / side_label / footnotes")
    l, r = split_cols(BODY, [1, 1], 24)
    lt, lb = split_rows(l, [1, 1], 14)
    before_after(s, lt, ["紙で申請", "押印のため出社"], ["Web で申請", "どこからでも承認"])
    icon_rows(s, lb, [("人", "体制", "専任3名を配置する"), ("¥", "予算", "初年度 2,000万円"), ("期", "期間", "12か月")])
    rt, rb = split_rows(r, [1, 1.1], 14)
    summary_box(s, rt, ["保守期限までに移行を完了する", "SaaS を基本とし、アドオンは最小限にする"])
    side_label(s, (rb[0], rb[1], 30, rb[3] - 36), "要点")
    exec_points(s, (rb[0] + 40, rb[1], rb[2] - 40, rb[3] - 36), [("移行を前倒し", "2027年度中に完了"),
                                                                  ("費用を抑制", "SaaS を採用")])
    footnotes(s, (rb[0], rb[1] + rb[3] - 30, rb[2], 30), ["金額は税抜", "2026年10月時点の見積り"])
