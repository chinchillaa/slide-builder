from pptx_helpers import *

def build(prs):
    s = content_slide(prs, "線・矢印部品", breadcrumb="部品カタログ：線・矢印",
                      lead="connect（種類別） / elbow_connect / curved_arrow / loop_arrow / block_arrow / flow_legend / span_bracket")
    x, y, w, h = BODY
    kinds = ["solid", "dashed", "dotted", "thick", "both"]
    for i, k in enumerate(kinds):
        a = (x, y + i * 46, 90, 32); b = (x + 260, y + i * 46, 90, 32)
        text(s, (a[0], a[1], a[2], a[3]), k, size=12, align="center", anchor="middle")
        connect(s, a, b, kind=k, label=FLOWS[k][4])
        _ = b
    a = (x + 420, y, 100, 40); b = (x + 620, y + 120, 100, 40)
    message_box(s, a, "A", style="pale", size=14); message_box(s, b, "B", style="pale", size=14)
    elbow_connect(s, a, b, first="h", label="カギ線")
    curved_arrow(s, (x + 420, y + 200), (x + 600, y + 250), kind="dashed", label="曲線")
    loop_arrow(s, (x + 760, y, 130, 130), "継続")
    block_arrow(s, (x + 420, y + 290, 220, 60), "太い矢印", "right")
    block_arrow(s, (x + 660, y + 290, 230, 60), "双方向", "both", style="mid")
    flow_legend(s, x, y + 250, ["service", "money", "goods", "info"])
    span_bracket(s, x, x + 350, y + 290, "対象範囲（2026年度）")
