# -*- coding: utf-8 -*-
"""parts_media - 画像・スライド枠の追加・カレンダー"""
import calendar
from datetime import date

from pptx.enum.dml import MSO_LINE_DASH_STYLE
from pptx.enum.shapes import MSO_SHAPE

from pptx_helpers import *  # noqa: F401,F403
from pptx_helpers import (_STATE, _blank, _box, _chrome_footer, _circle, _line, _put_text, _rect, _textbox,
                          content_slide, image, part)
from parts_text import grid_areas, split_cols, split_rows


@part("画像", "画像＋キャプション")
def image_caption(slide, area, path, caption, size=12):
    x, y, w, h = area
    image(slide, (x, y, w, h - 26), path)
    _textbox(slide, x, y + h - 24, w, 22, caption, size=size, color=COLOR_TEXT_SUB, align="center",
             dense=True, where="image_caption")


@part("画像", "画像を格子状に並べる（キャプション付き）")
def image_grid(slide, area, images, cols=3, size=11):
    """images: [(パス, キャプション)]"""
    rows = -(-len(images) // cols)
    for (p, cap), a in zip(images, grid_areas(area, rows, cols)):
        _rect(slide, *a, fill=COLOR_GRAY_PALE)
        image_caption(slide, (a[0] + 4, a[1] + 4, a[2] - 8, a[3] - 8), p, cap, size=size)


@part("画像", "画面キャプチャに番号付きの注釈")
def screenshot_callouts(slide, area, path, callouts, ratio=0.64, size=12):
    """callouts: [(x 0-1, y 0-1, 説明)] 画像上の相対位置に番号を置き、右に説明を並べる"""
    a, b = split_cols(area, [ratio, 1 - ratio], 20)
    pic = image(slide, a, path)
    px, py = pic.left / 12700, pic.top / 12700
    pw, ph = pic.width / 12700, pic.height / 12700
    _rect(slide, px, py, pw, ph, line=COLOR_GRAY)
    for i, (rx, ry, _) in enumerate(callouts):
        _circle(slide, px + rx * pw - 12, py + ry * ph - 12, 24, i + 1, fill=COLOR_ACCENT, size=12)
    for i, ((_, _, desc), (cx, cy, cw, ch)) in enumerate(zip(callouts, split_rows(b, [1] * len(callouts), 8))):
        _circle(slide, cx, cy + 4, 24, i + 1, fill=COLOR_ACCENT, size=12)
        _textbox(slide, cx + 32, cy, cw - 32, ch, desc, size=size, where=f"screenshot_callouts[{i}]")


@part("画像", "画像の差し込み予定枠（後で手で画像を貼る場所）")
def image_placeholder(slide, area, label="画像を挿入"):
    x, y, w, h = area
    sp = _rect(slide, x, y, w, h, fill=COLOR_GRAY_PALE, line=COLOR_GRAY_DARK)
    sp.line.dash_style = MSO_LINE_DASH_STYLE.DASH
    _put_text(sp, label, size=14, color=COLOR_TEXT_SUB, align="center", anchor="middle", where="image_placeholder")
    return sp


@part("スケジュール", "月間カレンダー（日付ごとの予定）")
def calendar_month(slide, area, year, month, events=None, size=10):
    """events: {日: "予定"}。土日は薄いグレー"""
    x, y, w, h = area
    events = events or {}
    heads = ["月", "火", "水", "木", "金", "土", "日"]
    cw = w / 7
    for j, t in enumerate(heads):
        _box(slide, x + j * cw + 1, y, cw - 2, 22, t, style="label" if j < 5 else "label2", size=11, dense=True)
    weeks = calendar.Calendar().monthdayscalendar(year, month)
    rh = (h - 26) / len(weeks)
    for i, wk in enumerate(weeks):
        for j, d in enumerate(wk):
            cx, cy = x + j * cw + 1, y + 26 + i * rh
            _rect(slide, cx, cy, cw - 2, rh - 2, fill=COLOR_GRAY_PALE if j >= 5 or d == 0 else COLOR_WHITE,
                  line=COLOR_GRAY_LIGHT)
            if d:
                _textbox(slide, cx + 2, cy + 1, 30, 16, str(d), size=10, bold=True, color=COLOR_TEXT_SUB, dense=True,
                         where="calendar日付", check=False)
                if d in events:
                    sp = _rect(slide, cx + 3, cy + 18, cw - 8, rh - 22, fill=COLOR_PRIMARY_PALE,
                               shape=MSO_SHAPE.ROUNDED_RECTANGLE)
                    _put_text(sp, events[d], size=size, color=COLOR_PRIMARY, bold=True, inset=(2, 1), dense=True,
                              where=f"calendar予定[{d}]")


@part("スライド枠", "アジェンダ（現在の項目を強調）")
def agenda_slide(prs, items, current=None, title="アジェンダ"):
    s = content_slide(prs, title)
    n = len(items)
    rh = min(44, 360 / n)
    for i, t in enumerate(items):
        on = current == i
        y = 100 + i * (rh + 8)
        sp = _rect(s, MARGIN_X + 60, y, CONTENT_W - 120, rh, fill=COLOR_PRIMARY if on else COLOR_WHITE,
                   line=None if on else COLOR_GRAY_LIGHT)
        _put_text(sp, f"{i + 1}.　{t}", size=18, bold=on, color=COLOR_WHITE if on else
                  (COLOR_TEXT if current is None else COLOR_GRAY_DARK), anchor="middle", inset=(20, 2),
                  where=f"agenda[{i}]")
    return s


@part("スライド枠", "締めのスライド（以上／問い合わせ先）")
def closing_slide(prs, message="以上", contact=None):
    s = _blank(prs)
    _textbox(s, 0, 200, SLIDE_W, 60, message, size=28, bold=True, color=COLOR_PRIMARY, align="center",
             anchor="middle", name="chrome:closing", where="closing")
    _line(s, SLIDE_W / 2 - 120, 268, SLIDE_W / 2 + 120, 268, COLOR_PRIMARY, 1.5)
    if contact:
        _textbox(s, 180, 290, SLIDE_W - 360, 80, contact, size=14, align="center", color=COLOR_TEXT_SUB,
                 where="closing連絡先")
    _chrome_footer(prs, s)
    return s


@part("スケジュール", "日程リスト（日付＋予定＋状態タグ）")
def schedule_list(slide, area, items, size=13):
    """items: [(日付ラベル, 予定, 状態 or None)]  状態は status_table と同じ（完了/進行中/未着手/遅延/保留）"""
    from parts_text import tag
    stmap = {"完了": "primary", "進行中": "mid", "未着手": "gray", "遅延": "alert", "保留": "gray"}
    rows = split_rows(area, [1] * len(items), 4)
    for i, ((d, t, st), (x, y, w, h)) in enumerate(zip(items, rows)):
        _box(slide, x, y, 120, h, d, style="pale", size=size, where=f"schedule_list日付[{i}]")
        _textbox(slide, x + 132, y, w - 132 - 90, h, t, size=size, anchor="middle", where=f"schedule_list予定[{i}]")
        if st:
            tag(slide, x + w - 80, y + h / 2 - 10, st, style=stmap.get(st, "gray"))
        _line(slide, x + 120, y + h + 2, x + w, y + h + 2, COLOR_GRAY_LIGHT, LINE_W)
