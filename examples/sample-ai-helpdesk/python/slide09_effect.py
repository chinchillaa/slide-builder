from pptx_helpers import *

def build(prs):
    s = content_slide(prs, "2.3 導入効果の試算", breadcrumb="2. 導入方針",
                      lead="一次回答の自動化で年間約1,700万円の工数を削減でき、2年目に投資を回収できる見込み。")
    l, r = split_cols(BODY, [1.7, 1], 24)
    roi_chart(s, l, ["2027年", "2028年", "2029年", "2030年", "2031年"],
              [1140, 480, 480, 480, 480], [850, 1700, 1700, 1700, 1700], unit="万円")
    a, b, c = split_rows(r, [1, 1, 1], 12)
    kpi(s, a, "45", "自動解決率", unit="%", note="PoCで検証")
    kpi(s, b, "3,400", "削減工数", unit="h/年", note="約1,700万円/年")
    kpi(s, c, "2", "投資回収", unit="年目", note="2028年")
    source(s, "出所：架空データ（サンプル）。2027年は7月の全社展開以降の半年分の効果")
