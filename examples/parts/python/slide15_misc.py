import os
from pptx_helpers import *

A = os.path.join(os.path.dirname(__file__), "..", "assets")

def build(prs):
    s = content_slide(prs, "補助部品", breadcrumb="部品カタログ：補助",
                      lead="text / arrow / highlight_frame / titled_box / column_header / image / grid_areas / nav_tabs")
    nav_tabs(s, ["概要", "詳細"], 0)
    cells = grid_areas(BODY, 2, 3, 20)
    x, y, w, h = cells[0]
    used = column_header(s, x, y, w, "地の文（text）")
    text(s, (x, y + used, w, h - used), "箇条書きにしない説明文。\n改行も使える。")
    titled_box(s, cells[1], "titled_box", ["見出し帯付きの汎用カード"], style="mid")
    x, y, w, h = cells[2]
    arrow(s, (x + 20, y + 30, 40, 80), "right")
    arrow(s, (x + 120, y + 40, 80, 40), "down")
    message_box(s, (x, y + h - 60, w, 50), "矢印は右・下の2方向", style="pale")
    x, y, w, h = cells[3]
    bullets(s, (x + 10, y + 30, w - 20, h - 40), ["ここを強調したい", "注目枠で囲む"])
    highlight_frame(s, (x, y + 24, w, h - 26), label="重要")
    image(s, cells[4], os.path.join(A, "photo.png"))
    image_placeholder(s, cells[5], "image_placeholder")
