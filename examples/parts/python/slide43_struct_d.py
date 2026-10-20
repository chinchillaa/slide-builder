from pptx_helpers import *

def build(prs):
    s = content_slide(prs, "構造部品④", breadcrumb="部品カタログ：構造",
                      lead="venn（2円・3円） / hub_spoke / concentric")
    a, b, c, d = split_cols(BODY, [1, 1, 1.3, 1], 16)
    venn(s, (a[0], a[1] + 60, a[2], 200), ["業務知識", "IT スキル"], center="DX 人材")
    venn(s, b, ["技術", "事業", "組織"], center="成功")
    hub_spoke(s, c, "データ基盤", ["営業", "経理", "物流", "経営", "人事", "顧客"])
    concentric(s, (d[0], d[1] + 60, d[2], d[3] - 60), ["全社", "本部", "部門"])
