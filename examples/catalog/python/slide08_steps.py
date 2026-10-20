from pptx_helpers import *

def build(prs):
    s = content_slide(prs, "3. 今後の進め方", breadcrumb="3. 今後の進め方",
                      lead="次年度は3つのステップで進め、2027年度末までに新システムへ移行する。")
    x, y, w, h = BODY
    step_flow(s, (x, y, w, 250), [
        ("Step1：要件定義", ["To-Be業務を確定する", "機能・非機能要件を定義する"]),
        ("Step2：設計・開発", ["パッケージを選定する", "アドオン開発は最小限にする"]),
        ("Step3：移行・展開", ["データを移行する", "部署ごとに順次展開する"]),
    ], active=0)
    kw = (w - 2 * GAP) / 3
    for i, (v, lab, note) in enumerate([("40", "手作業工数の削減", "h/月"),
                                        ("0", "転記ミス", "件/月 を目標"),
                                        ("2028.3", "移行完了", "保守期限まで")]):
        kpi(s, (x + i * (kw + GAP), y + 262, kw, h - 262), v, lab, note=note)
