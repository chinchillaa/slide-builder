# -*- coding: utf-8 -*-
"""parts_flow - フロー・プロセス・時系列の部品"""
import math

from pptx.enum.shapes import MSO_SHAPE

from pptx_helpers import *  # noqa: F401,F403
from pptx_helpers import (_arrow_line, _box, _circle, _line, _poly, _put_text, _rect, _split, _textbox,
                          _triangle, estimate_height, part)
from parts_text import split_cols, split_rows, titled_box


def _chevrons(slide, x, y, w, h, labels, active=None, size=14, fills=None):
    n = len(labels)
    widths = _split(w, [1] * n, 4)
    cx, out = x, []
    for i, lab in enumerate(labels):
        on = active is None or i == active
        fill = (fills[i] if fills else COLOR_PRIMARY) if on else COLOR_GRAY
        sp = _rect(slide, cx, y, widths[i] + (h * 0.3 if i < n - 1 else 0), h, fill=fill,
                   shape=MSO_SHAPE.PENTAGON if i == 0 else MSO_SHAPE.CHEVRON)
        sp.adjustments[0] = 0.3 if i == 0 else 0.3
        _put_text(sp, lab, size=size, color=COLOR_WHITE, bold=True, align="center", anchor="middle",
                  inset=(h * 0.35 if i else 6, 1), where=f"chevron[{i}]")
        out.append((cx, widths[i]))
        cx += widths[i] + 4
    return out


@part("フロー", "シェブロンの横一列（プロセス概要・現在地の表示）")
def chevron_row(slide, area, labels, active=None, size=14):
    x, y, w, h = area
    return _chevrons(slide, x, y, w, h, labels, active, size)


@part("フロー", "縦のステップ（番号丸を線でつなぐ）")
def vertical_steps(slide, area, steps, size=13):
    """steps: [(見出し, 説明)]"""
    x, y, w, h = area
    rows = split_rows(area, [1] * len(steps), 6)
    d = min(32, rows[0][3] - 4)
    _line(slide, x + d / 2, rows[0][1] + rows[0][3] / 2, x + d / 2, rows[-1][1] + rows[-1][3] / 2,
          COLOR_PRIMARY_LIGHT, 2)
    for i, ((t, desc), (rx, ry, rw, rh)) in enumerate(zip(steps, rows)):
        _circle(slide, rx, ry + (rh - d) / 2, d, i + 1, size=d * 0.45)
        _textbox(slide, rx + d + 12, ry, rw - d - 12, rh, [f"**{t}**", desc] if desc else f"**{t}**",
                 size=size, anchor="middle", where=f"vertical_steps[{i}]")


@part("フロー", "箱を矢印でつなぐ横フロー")
def flow_boxes(slide, area, items, style="text", size=13, gap=34):
    """items: ["文", ...] または [(見出し, 説明), ...]。style は card_table と同じ"""
    x, y, w, h = area
    cols = split_cols(area, [1] * len(items), gap)
    for i, (it, (cx, cy, cw, ch)) in enumerate(zip(items, cols)):
        body = [f"**{it[0]}**", it[1]] if isinstance(it, tuple) else it
        _box(slide, cx, cy, cw, ch, body, style=style, size=size, align="center", bullets=False,
             where=f"flow_boxes[{i}]")
        if i < len(items) - 1:
            _arrow_line(slide, cx + cw + 4, cy + ch / 2, cx + cw + gap - 4, cy + ch / 2, COLOR_PRIMARY, 2)
    return cols


@part("フロー", "スイムレーン（担当別の業務フロー）")
def swimlane(slide, area, lanes, n_cols, steps, arrows=(), col_headers=None, lane_w=90, size=11):
    """steps: [(lane, col, テキスト)]  arrows: [(from_step_index, to_step_index)]"""
    x, y, w, h = area
    top = y
    gx, gw = x + lane_w + 4, w - lane_w - 4
    cw = gw / n_cols
    if col_headers:
        for j, lab in enumerate(col_headers):
            _box(slide, gx + j * cw + 2, y, cw - 4, 24, lab, style="pale", size=size, dense=True,
                 where=f"swimlane列[{j}]")
        top = y + 28
    lh = (y + h - top) / len(lanes)
    for i, lab in enumerate(lanes):
        ly = top + i * lh
        _box(slide, x, ly + 1, lane_w, lh - 2, lab, style="label", size=size + 1, where=f"swimlaneレーン[{i}]")
        _rect(slide, gx, ly + 1, gw, lh - 2, fill=COLOR_GRAY_PALE if i % 2 == 0 else COLOR_WHITE)
    for j in range(1, n_cols):
        _line(slide, gx + j * cw, top, gx + j * cw, y + h, COLOR_GRAY_LIGHT, 0.5, dash=True)
    bw, bh = cw * 0.8, min(lh * 0.62, 44)
    pos = []
    for k, (ln, col, t) in enumerate(steps):
        bx, by = gx + col * cw + (cw - bw) / 2, top + ln * lh + (lh - bh) / 2
        _box(slide, bx, by, bw, bh, t, style="text", size=size, align="center", dense=True,
             where=f"swimlaneステップ[{k}]", inset=(3, 2))
        pos.append((bx, by, bw, bh))
    for a, b in arrows:
        ax, ay, aw, ah = pos[a]
        bx, by, bw2, bh2 = pos[b]
        if abs(ay - by) < 1:
            _arrow_line(slide, ax + aw, ay + ah / 2, bx, by + bh2 / 2, COLOR_PRIMARY, 1.5)
        elif abs(ax - bx) < 1:
            down = by > ay
            _arrow_line(slide, ax + aw / 2, ay + (ah if down else 0), bx + bw2 / 2, by + (0 if down else bh2),
                        COLOR_PRIMARY, 1.5)
        else:
            _arrow_line(slide, ax + aw, ay + ah / 2, bx, by + bh2 / 2, COLOR_PRIMARY, 1.5, elbow=True)


@part("フロー", "循環サイクル（PDCA など）")
def cycle(slide, area, labels, center=None, size=13):
    x, y, w, h = area
    n = len(labels)
    d = min(110, min(w, h) * 0.34)
    r = min(w, h) / 2 - d / 2 - 2
    cx, cy = x + w / 2, y + h / 2
    pts = [(cx + r * math.cos(-math.pi / 2 + 2 * math.pi * i / n),
            cy + r * math.sin(-math.pi / 2 + 2 * math.pi * i / n)) for i in range(n)]
    for i in range(n):
        (x1, y1), (x2, y2) = pts[i], pts[(i + 1) % n]
        L = math.hypot(x2 - x1, y2 - y1)
        k = (d / 2 + 6) / L
        _arrow_line(slide, x1 + (x2 - x1) * k, y1 + (y2 - y1) * k, x2 - (x2 - x1) * k, y2 - (y2 - y1) * k,
                    COLOR_PRIMARY_LIGHT, 3)
    if center:
        cd = max(r * 2 - d - 20, 40)
        _circle(slide, cx - cd / 2, cy - cd / 2, cd, None, fill=COLOR_PRIMARY_PALE)
        _textbox(slide, cx - cd / 2 + 4, cy - cd / 2 + 4, cd - 8, cd - 8, center, size=size, bold=True,
                 color=COLOR_PRIMARY, align="center", anchor="middle", where="cycle中心")
    for i, (px, py) in enumerate(pts):
        sp = _circle(slide, px - d / 2, py - d / 2, d, None, fill=COLOR_PRIMARY if i % 2 == 0 else COLOR_PRIMARY_MID)
        _put_text(sp, labels[i], size=size, color=COLOR_WHITE, bold=True, align="center", anchor="middle",
                  inset=(2, 2), where=f"cycle[{i}]")


@part("フロー", "ファネル（段階的な絞り込み）")
def funnel(slide, area, stages, size=13, min_ratio=0.35):
    """stages: [(ラベル, 値テキスト)]"""
    x, y, w, h = area
    n = len(stages)
    rows = split_rows(area, [1] * n, 4)
    cols = [COLOR_PRIMARY, COLOR_PRIMARY_MID, COLOR_PRIMARY_LIGHT, COLOR_GRAY, COLOR_GRAY_DARK]
    for i, ((lab, val), (rx, ry, rw, rh)) in enumerate(zip(stages, rows)):
        w0 = w * (1 - (1 - min_ratio) * i / n)
        w1 = w * (1 - (1 - min_ratio) * (i + 1) / n)
        cx = x + w / 2
        sp = _poly(slide, [(cx - w0 / 2, ry), (cx + w0 / 2, ry), (cx + w1 / 2, ry + rh), (cx - w1 / 2, ry + rh)],
                   fill=cols[i % len(cols)])
        dark = i % len(cols) < 2 or i % len(cols) == 4
        _put_text(sp, f"{lab}　**{val}**" if val else lab, size=size, color=COLOR_WHITE if dark else COLOR_TEXT,
                  bold=False, align="center", anchor="middle", inset=(w1 * 0.1 + 6, 1), where=f"funnel[{i}]")


@part("フロー", "ピラミッド（階層構造＋右に説明）")
def pyramid(slide, area, levels, size=13, ratio=0.45):
    """levels: 上から [(ラベル, 説明)]"""
    x, y, w, h = area
    pw = w * ratio
    n = len(levels)
    rows = split_rows((x, y, pw, h), [1] * n, 4)
    cols = [COLOR_PRIMARY, COLOR_PRIMARY_MID, COLOR_PRIMARY_LIGHT, COLOR_PRIMARY_PALE, COLOR_GRAY_PALE]
    for i, ((lab, desc), (rx, ry, rw, rh)) in enumerate(zip(levels, rows)):
        t0, t1 = (ry - y) / h, (ry + rh - y) / h
        cx = x + pw / 2
        pts = [(cx - pw * t0 / 2, ry), (cx + pw * t0 / 2, ry), (cx + pw * t1 / 2, ry + rh), (cx - pw * t1 / 2, ry + rh)]
        if t0 < 0.01:
            pts = [(cx, ry), (cx + pw * t1 / 2, ry + rh), (cx - pw * t1 / 2, ry + rh)]
        sp = _poly(slide, pts, fill=cols[i % len(cols)])
        _put_text(sp, lab, size=size, bold=True, color=COLOR_WHITE if i < 2 else COLOR_PRIMARY, align="center",
                  anchor="bottom" if t0 < 0.01 else "middle", inset=(2, 2), where=f"pyramid[{i}]",
                  check=t0 >= 0.01)
        if desc:
            _line(slide, cx + pw * (t0 + t1) / 4 + 4, ry + rh / 2, x + pw + 16, ry + rh / 2, COLOR_GRAY, LINE_W, dash=True)
            _textbox(slide, x + pw + 20, ry, w - pw - 20, rh, desc, size=size, anchor="middle",
                     where=f"pyramid説明[{i}]")


@part("フロー", "階段（段階的な成長・ステップアップ）")
def stair_steps(slide, area, steps, size=12):
    """steps: [(見出し, 説明)] 左から右へ上がっていく"""
    x, y, w, h = area
    n = len(steps)
    cols = split_cols(area, [1] * n, 4)
    fills = [COLOR_PRIMARY_PALE, COLOR_PRIMARY_LIGHT, COLOR_PRIMARY_MID, COLOR_PRIMARY, COLOR_PRIMARY]
    for i, ((t, desc), (cx, cy, cw, ch)) in enumerate(zip(steps, cols)):
        bh = h * (0.3 + 0.45 * i / max(n - 1, 1))
        fill = fills[min(i * 4 // max(n - 1, 1), 4)] if n > 1 else COLOR_PRIMARY
        sp = _rect(slide, cx, y + h - bh, cw, bh, fill=fill)
        _put_text(sp, t, size=size + 2, bold=True, color=COLOR_PRIMARY if fill in (COLOR_PRIMARY_PALE, COLOR_PRIMARY_LIGHT)
                  else COLOR_WHITE, align="center", anchor="top", inset=(4, 8), where=f"stair_steps[{i}]")
        if desc:
            _textbox(slide, cx, y + h - bh - 70, cw, 66, desc, size=size, anchor="bottom", align="center",
                     where=f"stair_steps説明[{i}]")


@part("フロー", "タイムライン（横軸に出来事を上下交互に配置）")
def timeline(slide, area, events, size=12):
    """events: [(日付ラベル, 見出し, 説明)]"""
    x, y, w, h = area
    n = len(events)
    my = y + h / 2
    _arrow_line(slide, x, my, x + w, my, COLOR_PRIMARY, 3)
    step = (w - 30) / n
    for i, (dt, t, desc) in enumerate(events):
        cx = x + step * (i + 0.5)
        _circle(slide, cx - 7, my - 7, 14, None, fill=COLOR_WHITE, line=COLOR_PRIMARY)
        up = i % 2 == 0
        _line(slide, cx, my + (-8 if up else 8), cx, my + (-24 if up else 24), COLOR_PRIMARY_LIGHT, 1)
        tw = min(step * 1.8, 200)
        ty = y if up else my + 26
        th = h / 2 - 28
        _textbox(slide, cx - tw / 2, ty, tw, th, [f"**{dt}**", t] + ([desc] if desc else []), size=size,
                 align="center", anchor="bottom" if up else "top", where=f"timeline[{i}]")


@part("フロー", "ロードマップ（フェーズ×領域に施策を配置）")
def roadmap(slide, area, phases, lanes, items, lane_w=120, size=12):
    """items: [(lane, 開始phase, 終了phase, テキスト, style="primary"|"mid"|"pale", sub=0)]"""
    x, y, w, h = area
    gx, gw = x + lane_w + 6, w - lane_w - 6
    _chevrons(slide, gx, y, gw, 34, phases, size=size + 1)
    pw = (gw - 4 * (len(phases) - 1)) / len(phases)
    top = y + 42
    lh = (y + h - top) / len(lanes)
    subs = {}
    for it in items:
        subs[it[0]] = max(subs.get(it[0], 0), (it[5] if len(it) > 5 else 0) + 1)
    for i, lab in enumerate(lanes):
        ly = top + i * lh
        _box(slide, x, ly + 2, lane_w, lh - 4, lab, style="label", size=size + 1, where=f"roadmap領域[{i}]")
        if i:
            _line(slide, gx, ly, gx + gw, ly, COLOR_GRAY_LIGHT, LINE_W, dash=True)
    sty = {"primary": "label2", "mid": "label", "pale": "pale"}
    for k, it in enumerate(items):
        ln, p0, p1, t = it[:4]
        st = it[4] if len(it) > 4 else "primary"
        sub = it[5] if len(it) > 5 else 0
        sh = (lh - 8) / subs[ln]
        bh = min(sh - 4, 40)
        bx = gx + p0 * (pw + 4)
        bw = (p1 - p0 + 1) * (pw + 4) - 4
        sp = _box(slide, bx, top + ln * lh + 4 + sub * sh + (sh - bh) / 2, bw, bh, t, style=sty[st], size=size,
                  shape=MSO_SHAPE.ROUNDED_RECTANGLE, dense=True, where=f"roadmap施策[{k}]")
        sp.adjustments[0] = 0.2


@part("フロー", "集約（複数の要素 → 1つの結果）")
def converge(slide, area, inputs, output, size=13, out_ratio=0.35):
    x, y, w, h = area
    iw = w * (1 - out_ratio) - 60
    rows = split_rows((x, y, iw, h), [1] * len(inputs), 8)
    ow, oh = w * out_ratio, min(h * 0.5, 120)
    ox, oy = x + w - ow, y + (h - oh) / 2
    for i, (it, (rx, ry, rw, rh)) in enumerate(zip(inputs, rows)):
        _box(slide, rx, ry, rw, rh, it, style="text", size=size, align="center", where=f"converge入力[{i}]")
        _arrow_line(slide, rx + rw + 2, ry + rh / 2, ox - 4, oy + oh / 2, COLOR_PRIMARY_LIGHT, 1.75)
    _box(slide, ox, oy, ow, oh, output, style="label", size=size + 1, where="converge出力")


@part("フロー", "展開（1つの要素 → 複数の結果）")
def diverge(slide, area, source_text, outputs, size=13, src_ratio=0.35):
    x, y, w, h = area
    sw, sh = w * src_ratio, min(h * 0.5, 120)
    sy = y + (h - sh) / 2
    ox = x + sw + 60
    rows = split_rows((ox, y, w - sw - 60, h), [1] * len(outputs), 8)
    _box(slide, x, sy, sw, sh, source_text, style="label", size=size + 1, where="diverge起点")
    for i, (it, (rx, ry, rw, rh)) in enumerate(zip(outputs, rows)):
        _arrow_line(slide, x + sw + 4, sy + sh / 2, rx - 2, ry + rh / 2, COLOR_PRIMARY_LIGHT, 1.75)
        _box(slide, rx, ry, rw, rh, it, style="text", size=size, align="center", where=f"diverge出力[{i}]")


@part("フロー", "Input → Process → Output")
def ipo(slide, area, inputs, process, outputs, labels=("Input", "Process", "Output"), size=13):
    x, y, w, h = area
    cols = split_cols(area, [1, 1.2, 1], 40)
    for i, (lab, items, st) in enumerate(zip(labels, (inputs, process, outputs), ("pale", "primary", "pale"))):
        titled_box(slide, cols[i], lab, items, style=st, size=size)
        if i < 2:
            _triangle(slide, cols[i][0] + cols[i][2] + 10, y + h / 2 - 26, 20, 52, fill=COLOR_GRAY)


@part("フロー", "フェーズゲート（フェーズの間に判定ゲート）")
def phase_gate(slide, area, phases, gates, size=12):
    """phases: [(フェーズ名, items)]  gates: フェーズ間のゲート名（len = フェーズ数-1）"""
    x, y, w, h = area
    cols = split_cols(area, [1] * len(phases), 46)
    for i, ((t, items), a) in enumerate(zip(phases, cols)):
        titled_box(slide, a, t, items, style="primary", size=size)
        if i < len(phases) - 1:
            gx = a[0] + a[2] + 7
            _rect(slide, gx, y + 1, 32, 32, fill=COLOR_PRIMARY_MID, shape=MSO_SHAPE.DIAMOND)
            _textbox(slide, gx - 20, y + 36, 72, 40, gates[i], size=10, bold=True, color=COLOR_PRIMARY_MID,
                     align="center", dense=True, where=f"phase_gateゲート[{i}]")


@part("フロー", "バリューチェーン（支援活動の帯＋主活動のシェブロン）")
def value_chain(slide, area, primary, support=(), size=12):
    x, y, w, h = area
    sh = 30
    for i, lab in enumerate(support):
        _box(slide, x, y + i * (sh + 4), w - 30, sh, lab, style="pale", size=size, align="left",
             where=f"value_chain支援[{i}]", inset=(12, 2))
    top = y + len(support) * (sh + 4) + (8 if support else 0)
    _chevrons(slide, x, top, w, y + h - top, primary, size=size)


@part("フロー", "ステップ×観点のマトリクス（シェブロン見出し）")
def step_matrix(slide, area, steps, row_labels, cells, label_w=110, size=11):
    """cells[行][ステップ] は str か list"""
    x, y, w, h = area
    gx, gw = x + label_w + 6, w - label_w - 6
    cols = _chevrons(slide, gx, y, gw, 34, steps, size=size + 1)
    rh = _split(h - 42, [1] * len(row_labels), 4)
    ry = y + 42
    for i, lab in enumerate(row_labels):
        _box(slide, x, ry, label_w, rh[i], lab, style="label", size=size + 1, where=f"step_matrix行[{i}]")
        for j, (cx, cw) in enumerate(cols):
            _box(slide, cx, ry, cw, rh[i], cells[i][j], style="text", size=size, dense=True,
                 where=f"step_matrix[{i},{j}]")
        ry += rh[i] + 4
