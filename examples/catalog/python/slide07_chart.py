from pptx_helpers import *

def build(prs):
    s = content_slide(prs, "2.2. 定量分析", breadcrumb="2. 現状分析",
                      lead="手作業工数は3年で**1.4倍**に増えており、売上集計が全体の約4割を占める。")
    areas, _ = split_columns(s, BODY, [1.3, 1])
    chart(s, areas[0], "column", ["2023年度", "2024年度", "2025年度", "2026年度(見込)"],
          {"手作業工数": [72, 81, 92, 100]}, title="手作業工数の推移", unit="h/月", highlight=3)
    chart(s, areas[1], "doughnut", ["売上集計", "請求書発行", "在庫照合", "その他"],
          {"構成比": [40, 24, 16, 20]}, title="業務別の内訳（2026年度）")
    source(s, "各部署の業務日報から集計")
