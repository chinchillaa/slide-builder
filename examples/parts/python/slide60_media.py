import os
from pptx_helpers import *

A = os.path.join(os.path.dirname(__file__), "..", "assets")

def build(prs):
    s = content_slide(prs, "画像部品", breadcrumb="部品カタログ：画像",
                      lead="screenshot_callouts / image_caption / image_grid / image_placeholder")
    top, bot = split_rows(BODY, [1.3, 1], 14)
    screenshot_callouts(s, top, os.path.join(A, "screen.png"),
                        [(0.1, 0.25, "メニューから機能を選ぶ"), (0.45, 0.2, "主要な指標を表示"), (0.6, 0.7, "月別の推移をグラフで確認")])
    a, b, c = split_cols(bot, [1, 1.6, 1], 16)
    image_caption(s, a, os.path.join(A, "photo.png"), "図1 現地の様子")
    image_grid(s, b, [(os.path.join(A, "photo.png"), "工場A"), (os.path.join(A, "photo.png"), "工場B")], cols=2)
    image_placeholder(s, c, "ここに組織図の画像")
