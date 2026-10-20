# -*- coding: utf-8 -*-
"""parts_text - レイアウト補助・テキスト・カード系の部品"""
from pptx.enum.shapes import MSO_SHAPE
from pptx.oxml.ns import qn

from pptx_helpers import *  # noqa: F401,F403
from pptx_helpers import (_CELL, _arrow_line, _box, _circle, _line, _plain, _put_text, _rect,
                          _split, _textbox, _triangle, estimate_height, part, text_width)


# ---------------------------------------------------------------- レイアウト
@part("レイアウト", "上下分割（描画なし。領域を返す）")
def split_rows(area, ratios, gap=GAP):
    x, y, w, h = area
    out, cy = [], y
    for rh in _split(h, ratios, gap):
        out.append((x, cy, w, rh))
        cy += rh + gap
    return out


@part("レイアウト", "格子状に分割（描画なし。行優先で領域を返す）")
def grid_areas(area, rows, cols, gap=GAP):
    out = []
    for rx, ry, rw, rh in split_rows(area, [1] * rows, gap):
        out += [(cx, ry, cw, rh) for cx, cw in _cols(rx, rw, [1] * cols, gap)]
    return out


def _cols(x, w, ratios, gap):
    """横方向の分割: [(x, w), ...]"""
    out, cx = [], x
    for cw in _split(w, ratios, gap):
        out.append((cx, cw))
        cx += cw + gap
    return out


@part("レイアウト", "左右に等分・比率分割（描画なし。領域を返す）")
def split_cols(area, ratios, gap=GAP):
    x, y, w, h = area
    return [(cx, y, cw, h) for cx, cw in _cols(x, w, ratios, gap)]


@part("レイアウト", "縦書きの帯ラベル（左端の分類など）")
def side_label(slide, area, label, style="label", size=14):
    x, y, w, h = area
    need = len(_plain(label)) * size * 1.05
    if need > h - 8:
        raise ValueError(f"[side_label] 縦書きラベル「{label}」が高さ {h:.0f}pt に収まりません（必要 {need:.0f}pt）")
    sp = _box(slide, x, y, w, h, label, style=style, size=size, where="side_label")
    sp.text_frame._txBody.bodyPr.set("vert", "eaVert")
    return sp


@part("レイアウト", "紺の帯見出し（セクション内の大見出し）。消費高さを返す")
def band_heading(slide, x, y, w, label, size=14):
    sp = _rect(slide, x, y, w, 26, fill=COLOR_PRIMARY)
    _put_text(sp, label, size=size, color=COLOR_WHITE, bold=True, anchor="middle", inset=(10, 1),
              where="band_heading")
    return 32


@part("レイアウト", "■付きの小見出し＋細罫線。消費高さを返す")
def sub_heading(slide, x, y, w, label, size=14):
    _rect(slide, x, y + 8, 10, 10, fill=COLOR_PRIMARY)
    _textbox(slide, x + 16, y, w - 16, 26, label, size=size, bold=True, color=COLOR_PRIMARY,
             anchor="middle", where="sub_heading")
    _line(slide, x, y + 27, x + w, y + 27, COLOR_GRAY_LIGHT, LINE_W)
    return 34


@part("レイアウト", "区切り線")
def divider(slide, x, y, w, dash=False):
    return _line(slide, x, y, x + w, y, COLOR_GRAY_DARK, LINE_W, dash=dash)


# ---------------------------------------------------------------- テキスト
@part("テキスト", "番号付きリスト（丸数字＋見出し＋説明）")
def numbered_list(slide, area, items, size=14):
    """items: ["文", ...] または [(見出し, 説明), ...]"""
    x, y, w, h = area
    rows = split_rows(area, [1] * len(items), 8)
    for i, (it, (rx, ry, rw, rh)) in enumerate(zip(items, rows)):
        d = min(30, rh - 4)
        _circle(slide, rx, ry + (rh - d) / 2, d, i + 1, size=min(14, d * 0.5))
        body = [f"**{it[0]}**", it[1]] if isinstance(it, tuple) else it
        _textbox(slide, rx + d + 10, ry, rw - d - 10, rh, body, size=size, anchor="middle",
                 where=f"numbered_list[{i}]")


@part("テキスト", "吹き出し（指し示す位置を tail で指定）")
def callout(slide, area, body, tail, size=12):
    """tail=(x, y) 吹き出しの先端の座標"""
    x, y, w, h = area
    sp = _rect(slide, x, y, w, h, fill=COLOR_PRIMARY_PALE, line=COLOR_PRIMARY,
               shape=MSO_SHAPE.RECTANGULAR_CALLOUT)
    sp.adjustments[0] = (tail[0] - (x + w / 2)) / w
    sp.adjustments[1] = (tail[1] - (y + h / 2)) / h
    return _put_text(sp, body, size=size, color=COLOR_PRIMARY, bold=True, anchor="middle",
                     align="center", where="callout")


@part("テキスト", "補足ボックス（グレー地＋「補足」タグ）")
def note_box(slide, area, items, label="補足", size=12):
    x, y, w, h = area
    _rect(slide, x, y, w, h, fill=COLOR_GRAY_PALE, line=COLOR_GRAY)
    tag(slide, x + 8, y + 8, label, style="primary")
    return _textbox(slide, x + 8, y + 32, w - 16, h - 38, items, size=size, bullets=isinstance(items, list),
                    where="note_box")


@part("テキスト", "引用・ヒアリングの声（発言者付き）")
def quote_box(slide, area, quote, speaker=None, size=14):
    x, y, w, h = area
    _rect(slide, x, y, w, h, fill=COLOR_WHITE, line=COLOR_GRAY)
    _rect(slide, x, y, 5, h, fill=COLOR_PRIMARY)
    _textbox(slide, x + 12, y + 2, 40, 40, "“", size=36, bold=True, color=COLOR_PRIMARY_LIGHT,
             where="quote_mark", check=False, name="deco:quote")
    sh = 20 if speaker else 0
    _textbox(slide, x + 44, y + 10, w - 56, h - 20 - sh, quote, size=size, anchor="middle", where="quote_box")
    if speaker:
        _textbox(slide, x + 44, y + h - 26, w - 56, 22, f"― {speaker}", size=12, color=COLOR_TEXT_SUB,
                 align="right", where="quote_speaker")


@part("テキスト", "大きな一文のステートメント（上下罫線）")
def statement(slide, area, body, sub=None, size=24):
    x, y, w, h = area
    _line(slide, x + w * 0.1, y, x + w * 0.9, y, COLOR_PRIMARY, 1.5)
    _line(slide, x + w * 0.1, y + h, x + w * 0.9, y + h, COLOR_PRIMARY, 1.5)
    sh = 30 if sub else 0
    _textbox(slide, x, y + 6, w, h - 12 - sh, body, size=size, bold=True, color=COLOR_PRIMARY,
             align="center", anchor="middle", where="statement")
    if sub:
        _textbox(slide, x, y + h - 36, w, 28, sub, size=14, align="center", color=COLOR_TEXT_SUB,
                 where="statement_sub")


_TAG = {"primary": (COLOR_PRIMARY, None, COLOR_WHITE), "mid": (COLOR_PRIMARY_MID, None, COLOR_WHITE),
        "pale": (COLOR_PRIMARY_PALE, None, COLOR_PRIMARY), "outline": (COLOR_WHITE, COLOR_PRIMARY, COLOR_PRIMARY),
        "gray": (COLOR_GRAY_PALE, COLOR_GRAY, COLOR_TEXT_SUB), "alert": (COLOR_WHITE, COLOR_ACCENT, COLOR_ACCENT)}


@part("テキスト", "小さなタグ・バッジ（幅は自動）。幅を返す")
def tag(slide, x, y, label, style="primary", size=11, h=20):
    fill, line, col = _TAG[style]
    w = text_width(label, size) + 16
    sp = _rect(slide, x, y, w, h, fill=fill, line=line, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    _put_text(sp, label, size=size, color=col, bold=True, align="center", anchor="middle", inset=(2, 0),
              dense=True, where="tag", check=False)
    return w


@part("テキスト", "Q&A の一覧")
def qa_list(slide, area, pairs, size=13):
    rows = split_rows(area, [1] * len(pairs), 10)
    for i, ((q, a), (x, y, w, h)) in enumerate(zip(pairs, rows)):
        qh = estimate_height(q, w - 40, size) + 8
        _circle(slide, x, y, 24, "Q", size=12)
        _textbox(slide, x + 32, y, w - 32, max(qh, 24), q, size=size, bold=True, anchor="middle",
                 where=f"qa_list Q[{i}]")
        _circle(slide, x, y + max(qh, 24) + 4, 24, "A", fill=COLOR_PRIMARY_MID, size=12)
        _textbox(slide, x + 32, y + max(qh, 24) + 4, w - 32, h - max(qh, 24) - 4, a, size=size,
                 where=f"qa_list A[{i}]")


@part("テキスト", "用語定義の一覧（用語＋説明）")
def definition_list(slide, area, pairs, term_ratio=0.25, size=13):
    x, y, w, h = area
    tw = w * term_ratio
    for i, ((term, desc), (rx, ry, rw, rh)) in enumerate(zip(pairs, split_rows(area, [1] * len(pairs), 4))):
        _box(slide, rx, ry, tw, rh, term, style="pale", size=size, where=f"definition_list用語[{i}]")
        _textbox(slide, rx + tw + 10, ry, rw - tw - 10, rh, desc, size=size, anchor="middle",
                 where=f"definition_list説明[{i}]")
        _line(slide, rx + tw, ry + rh + 2, rx + rw, ry + rh + 2, COLOR_GRAY_LIGHT, LINE_W)


@part("テキスト", "チェックリスト（☑/□）")
def checklist(slide, area, items, checked=None, size=14):
    x, y, w, h = area
    checked = checked or [False] * len(items)
    cy = y
    for i, it in enumerate(items):
        ih = max(estimate_height(it, w - 34, size) + 6, 22)
        _rect(slide, x, cy + 3, 16, 16, fill=COLOR_WHITE, line=COLOR_PRIMARY, line_w=1.25)
        if checked[i]:
            _textbox(slide, x - 2, cy - 1, 22, 22, "✓", size=16, bold=True, color=COLOR_PRIMARY_MID,
                     align="center", where="check", check=False)
        _textbox(slide, x + 26, cy, w - 26, ih, it, size=size, where=f"checklist[{i}]")
        cy += ih + 6
    if cy - 6 > y + h + 0.5:
        raise ValueError(f"[checklist] 項目が領域に収まりません（推定 {cy - 6 - y:.0f}pt > {h:.0f}pt）")


_HEAD = {"primary": (COLOR_PRIMARY, COLOR_WHITE, COLOR_PRIMARY),
         "mid": (COLOR_PRIMARY_MID, COLOR_WHITE, COLOR_PRIMARY_MID),
         "pale": (COLOR_PRIMARY_PALE, COLOR_PRIMARY, COLOR_PRIMARY_LIGHT),
         "gray": (COLOR_GRAY, COLOR_TEXT, COLOR_GRAY)}


@part("カード", "見出し帯付きのボックス（汎用カード）")
def titled_box(slide, area, title, items, style="primary", size=14, head_h=30):
    x, y, w, h = area
    fill, col, line = _HEAD[style]
    sp = _rect(slide, x, y, w, head_h, fill=fill)
    _put_text(sp, title, size=size, color=col, bold=True, align="center", anchor="middle",
              where="titled_box見出し")
    body = _rect(slide, x, y + head_h, w, h - head_h, fill=COLOR_WHITE, line=line)
    if items:
        _put_text(body, items, size=size, bullets=isinstance(items, list), inset=(8, 6),
                  where="titled_box本文")
    return body


@part("カード", "カードを格子状に並べる")
def card_grid(slide, area, cards, cols=3, style="primary", size=13):
    rows = -(-len(cards) // cols)
    areas = grid_areas(area, rows, cols)
    for (t, items), a in zip(cards, areas):
        titled_box(slide, a, t, items, style=style, size=size)
    return areas


@part("カード", "ポイントカード（POINT 番号＋見出し＋説明）")
def point_cards(slide, area, cards, label="POINT", size=13):
    for i, ((t, body), a) in enumerate(zip(cards, split_cols(area, [1] * len(cards)))):
        x, y, w, h = a
        _rect(slide, x, y, w, h, fill=COLOR_WHITE, line=COLOR_GRAY)
        _rect(slide, x, y, w, 4, fill=COLOR_PRIMARY)
        _textbox(slide, x + 10, y + 10, w - 20, 18, f"{label} {i + 1}", size=11, bold=True,
                 color=COLOR_PRIMARY_MID, dense=True, where="point_cards番号")
        th = estimate_height(t, w - 20, 15) + 6
        _textbox(slide, x + 10, y + 30, w - 20, th, t, size=15, bold=True, color=COLOR_PRIMARY,
                 where=f"point_cards見出し[{i}]")
        _line(slide, x + 10, y + 34 + th, x + w - 10, y + 34 + th, COLOR_GRAY_LIGHT, LINE_W)
        _textbox(slide, x + 10, y + 40 + th, w - 20, h - 48 - th, body, size=size,
                 bullets=isinstance(body, list), where=f"point_cards本文[{i}]")


@part("テキスト", "結論バー（左に「結論」ラベル＋淡青の本文）")
def conclusion_bar(slide, area, body, label="結論", size=16):
    x, y, w, h = area
    _box(slide, x, y, 90, h, label, style="label", size=size)
    sp = _rect(slide, x + 90, y, w - 90, h, fill=COLOR_PRIMARY_PALE)
    _put_text(sp, body, size=size, color=COLOR_PRIMARY, bold=True, anchor="middle", inset=(12, 4),
              where="conclusion_bar")


@part("カード", "メリット／デメリットの対比")
def pros_cons(slide, area, pros, cons, titles=("メリット", "デメリット"), size=14):
    a1, a2 = split_cols(area, [1, 1], GAP + 4)
    titled_box(slide, a1, "＋ " + titles[0], pros, style="primary", size=size)
    titled_box(slide, a2, "－ " + titles[1], cons, style="gray", size=size)


@part("カード", "Before → After の対比（矢印付き）")
def before_after(slide, area, before, after, labels=("Before", "After"), size=14):
    x, y, w, h = area
    cw = (w - 50) / 2
    titled_box(slide, (x, y, cw, h), labels[0], before, style="gray", size=size)
    _triangle(slide, x + cw + 12, y + h / 2 - 30, 26, 60, fill=COLOR_PRIMARY_MID)
    titled_box(slide, (x + cw + 50, y, cw, h), labels[1], after, style="primary", size=size)


@part("テキスト", "アイコン丸＋見出し＋説明の行リスト")
def icon_rows(slide, area, rows, size=13):
    """rows: [(アイコン文字1-2字, 見出し, 説明)]"""
    for i, ((icon, t, desc), (x, y, w, h)) in enumerate(zip(rows, split_rows(area, [1] * len(rows), 8))):
        d = min(44, h - 4)
        _circle(slide, x, y + (h - d) / 2, d, icon, fill=COLOR_PRIMARY_PALE, color=COLOR_PRIMARY,
                size=min(16, d * 0.42))
        _textbox(slide, x + d + 12, y, w - d - 12, h, [f"**{t}**", desc], size=size, anchor="middle",
                 where=f"icon_rows[{i}]")


@part("テキスト", "まとめボックス（紺枠＋左上タイトルタブ）")
def summary_box(slide, area, items, title="まとめ", size=14):
    x, y, w, h = area
    _rect(slide, x, y + 10, w, h - 10, fill=COLOR_WHITE, line=COLOR_PRIMARY, line_w=1.5)
    tw = text_width(title, size) + 30
    _box(slide, x + 12, y, tw, 22, title, style="label", size=size, inset=(2, 0))
    _textbox(slide, x + 12, y + 30, w - 24, h - 38, items, size=size, bullets=isinstance(items, list),
             where="summary_box")


@part("テキスト", "エグゼクティブサマリー（番号＋見出し＋説明の行）")
def exec_points(slide, area, points, size=14):
    for i, ((head, desc), (x, y, w, h)) in enumerate(zip(points, split_rows(area, [1] * len(points), 6))):
        _box(slide, x, y, 36, h, str(i + 1), style="label", size=16)
        _textbox(slide, x + 46, y, w - 46, h, [f"**{head}**", desc], size=size,
                 anchor="middle", where=f"exec_points[{i}]")
        _line(slide, x + 46, y + h + 3, x + w, y + h + 3, COLOR_GRAY_LIGHT, LINE_W)


@part("補助", "脚注リスト（注1, 注2…）")
def footnotes(slide, area, notes, size=10):
    items = [n if n.startswith(("注", "※")) else f"注{i + 1}：{n}" for i, n in enumerate(notes)]
    return _textbox(slide, *area, items, size=size, color=COLOR_TEXT_SUB, dense=True, where="footnotes")


@part("補助", "赤の点線で囲む注目枠（ラベル付き可）")
def highlight_frame(slide, area, label=None):
    x, y, w, h = area
    sp = _rect(slide, x, y, w, h, line=COLOR_ACCENT, line_w=2, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    sp.adjustments[0] = 0.05
    from pptx.enum.dml import MSO_LINE_DASH_STYLE
    sp.line.dash_style = MSO_LINE_DASH_STYLE.DASH
    sp.name = "frame"
    if label:
        lw = text_width(label, 11) + 16
        tag(slide, x + w - lw, y - 22, label, style="alert")
    return sp


@part("補助", "丸数字バッジ")
def number_badge(slide, x, y, n, d=24, style="primary"):
    fill = {"primary": COLOR_PRIMARY, "mid": COLOR_PRIMARY_MID, "alert": COLOR_ACCENT}[style]
    return _circle(slide, x, y, d, n, fill=fill, size=d * 0.5)


@part("補助", "右上の「参考」などの区分タグ")
def ref_tag(slide, label="参考"):
    sp = _rect(slide, SLIDE_W - 110, 32, 90, 22, fill=COLOR_WHITE, line=COLOR_GRAY_DARK)
    sp.name = "chrome:ref_tag"
    _put_text(sp, label, size=11, color=COLOR_TEXT_SUB, bold=True, align="center", anchor="middle",
              inset=(2, 0), dense=True, where="ref_tag")
    return sp


@part("補助", "凡例（色見本＋ラベル）。幅を返す")
def legend(slide, x, y, items, size=11):
    """items: [(ラベル, COLOR_*), ...]"""
    cx = x
    for lab, col in items:
        _rect(slide, cx, y + 4, 10, 10, fill=col)
        tw = text_width(lab, size) + 8
        _textbox(slide, cx + 13, y, tw, 18, lab, size=size, dense=True, anchor="middle", where="legend",
                 check=False)
        cx += 13 + tw + 10
    return cx - x
