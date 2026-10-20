# -*- coding: utf-8 -*-
"""parts_structure - 構造図・関係図の部品（ツリー・マトリクス・ベン図・体制図など）"""
import math

from pptx.enum.shapes import MSO_SHAPE

from pptx_helpers import *  # noqa: F401,F403
from pptx_helpers import (_alpha, _arrow_line, _box, _circle, _line, _put_text, _rect, _split, _textbox,
                          _triangle, part)
from parts_text import definition_list, grid_areas, split_cols, split_rows, titled_box

_DEPTH = ["label", "label2", "text", "text"]


def _tree_leaves(node):
    label, kids = (node, []) if isinstance(node, str) else node
    return max(1, sum(_tree_leaves(k) for k in kids)) if kids else 1


def _norm(node):
    return (node, []) if isinstance(node, str) else node


@part("構造", "ロジックツリー（左から右へ分解）")
def logic_tree(slide, area, root, size=12, col_ratios=None):
    """root = (ラベル, [子, ...])  子は str または (ラベル, [孫, ...])。4階層まで"""
    x, y, w, h = area
    def _d(n):
        k = _norm(n)[1]
        return 1 + (max(_d(c) for c in k) if k else 0)
    depth = _d(root)
    if depth > 4:
        raise ValueError("[logic_tree] 4階層までにしてください")
    cols = split_cols(area, col_ratios or [1, 1.2, 1.4, 1.4][:depth], 30 if depth == 4 else 36)
    total = _tree_leaves(root)
    slot = h / total

    def draw(node, lv, y0, n_leaves):
        label, kids = _norm(node)
        span = n_leaves * slot
        nh = min(span - 8, 44 if lv else 70)
        cx, _, cw, _ = cols[lv]
        ny = y0 + (span - nh) / 2
        _box(slide, cx, ny, cw, nh, label, style=_DEPTH[lv], size=size, dense=True, where=f"logic_tree[{label[:8]}]")
        cy = y0
        for k in kids:
            kl = _tree_leaves(k)
            ky, kh = draw(k, lv + 1, cy, kl)
            _arrow_line(slide, cx + cw, ny + nh / 2, cols[lv + 1][0], ky + kh / 2, COLOR_GRAY_DARK, 1, head=False,
                        elbow=True)
            cy += kl * slot
        return ny, nh
    draw(root, 0, y, total)


@part("構造", "組織図・体制図（上から下へ）")
def org_chart(slide, area, root, size=12):
    """root = (ラベル, [子, ...])  ラベルに "\\n" で役職と氏名を2行にできる"""
    x, y, w, h = area

    def depth(n):
        _, k = _norm(n)
        return 1 + (max(depth(c) for c in k) if k else 0)
    d = depth(root)
    rh = h / d
    total = _tree_leaves(root)
    slot = w / total

    def draw(node, lv, x0, n_leaves):
        label, kids = _norm(node)
        span = n_leaves * slot
        bw, bh = min(span - 10, 170), min(rh - 22, 48)
        bx, by = x0 + (span - bw) / 2, y + lv * rh
        _box(slide, bx, by, bw, bh, label, style=_DEPTH[min(lv, 3)], size=size, dense=True, align="center",
             where=f"org_chart[{label[:8]}]", inset=(3, 2))
        cx = x0
        tops = []
        for k in kids:
            kl = _tree_leaves(k)
            kx, ky, kw = draw(k, lv + 1, cx, kl)
            tops.append((kx + kw / 2, ky))
            cx += kl * slot
        if tops:   # 親の下 → 横の連結線 → 各子の上（縦優先の配線）
            my = by + bh + (tops[0][1] - by - bh) / 2
            _line(slide, bx + bw / 2, by + bh, bx + bw / 2, my, COLOR_GRAY_DARK, 1)
            _line(slide, min(t[0] for t in tops + [(bx + bw / 2, 0)]), my,
                  max(t[0] for t in tops + [(bx + bw / 2, 0)]), my, COLOR_GRAY_DARK, 1)
            for tx, ty in tops:
                _line(slide, tx, my, tx, ty, COLOR_GRAY_DARK, 1)
        return bx, by, bw
    draw(root, 0, x, total)


@part("構造", "2×2 マトリクス（軸付きの4象限）")
def matrix_2x2(slide, area, x_axis, y_axis, quadrants, highlight=None, size=13):
    """x_axis/y_axis = (軸名, 低側ラベル, 高側ラベル)。quadrants = [左上, 右上, 左下, 右下]（str か (見出し, 説明)）"""
    x, y, w, h = area
    ax, ay = x + 40, y
    aw, ah = w - 40, h - 36
    _arrow_line(slide, ax - 6, ay + ah + 6, ax - 6, ay, COLOR_PRIMARY, 2)
    _arrow_line(slide, ax - 6, ay + ah + 6, ax + aw, ay + ah + 6, COLOR_PRIMARY, 2)
    _textbox(slide, x, ay + ah * 0.25, 30, ah * 0.5, y_axis[0], size=12, bold=True, color=COLOR_PRIMARY, align="center",
             anchor="middle", where="matrix_2x2縦軸", check=False).text_frame._txBody.bodyPr.set("vert", "eaVert")
    _textbox(slide, ax + aw * 0.25, ay + ah + 12, aw * 0.5, 22, f"{x_axis[0]}", size=12, bold=True, color=COLOR_PRIMARY,
             align="center", where="matrix_2x2横軸")
    _textbox(slide, ax, ay + ah + 12, 80, 22, x_axis[1], size=11, color=COLOR_TEXT_SUB, dense=True, where="軸")
    _textbox(slide, ax + aw - 80, ay + ah + 12, 80, 22, x_axis[2], size=11, color=COLOR_TEXT_SUB, align="right",
             dense=True, where="軸")
    _textbox(slide, x, ay + ah - 18, 30, 20, y_axis[1], size=10, color=COLOR_TEXT_SUB, align="center", dense=True,
             where="軸", check=False)
    _textbox(slide, x, ay, 30, 20, y_axis[2], size=10, color=COLOR_TEXT_SUB, align="center", dense=True,
             where="軸", check=False)
    for i, (q, a) in enumerate(zip(quadrants, grid_areas((ax, ay, aw, ah), 2, 2, 6))):
        body = [f"**{q[0]}**", q[1]] if isinstance(q, tuple) else q
        _box(slide, *a, body, style="label" if i == highlight else "pale", size=size, align="center",
             bullets=False, where=f"matrix_2x2[{i}]")


@part("構造", "ポジショニングマップ（2軸に項目をプロット）")
def positioning_map(slide, area, x_axis, y_axis, points, size=11):
    """x_axis/y_axis = (軸名, 低側, 高側)。points = [(x 0-1, y 0-1, ラベル, 強調bool)]"""
    x, y, w, h = area
    cx, cy = x + w / 2, y + h / 2
    _arrow_line(slide, x, cy, x + w, cy, COLOR_GRAY_DARK, 1.25, both=True)
    _arrow_line(slide, cx, y + h, cx, y, COLOR_GRAY_DARK, 1.25, both=True)
    _textbox(slide, x + w - 160, cy + 4, 160, 18, f"{x_axis[2]}", size=11, color=COLOR_TEXT_SUB, align="right", dense=True, where="軸")
    _textbox(slide, x, cy + 4, 160, 18, f"{x_axis[1]}", size=11, color=COLOR_TEXT_SUB, dense=True, where="軸")
    _textbox(slide, cx + 6, y, 200, 18, f"{y_axis[2]}", size=11, color=COLOR_TEXT_SUB, dense=True, where="軸")
    _textbox(slide, cx + 6, y + h - 18, 200, 18, f"{y_axis[1]}", size=11, color=COLOR_TEXT_SUB, dense=True, where="軸")
    _textbox(slide, x + w - 200, cy - 22, 200, 18, x_axis[0], size=11, bold=True, color=COLOR_PRIMARY, align="right",
             dense=True, where="軸名")
    _textbox(slide, cx - 206, y, 200, 18, y_axis[0], size=11, bold=True, color=COLOR_PRIMARY, align="right",
             dense=True, where="軸名")
    for px, py, lab, *hl in points:
        on = bool(hl and hl[0])
        d = 16 if on else 12
        ex, ey = x + px * w, y + (1 - py) * h
        _circle(slide, ex - d / 2, ey - d / 2, d, None, fill=COLOR_PRIMARY if on else COLOR_PRIMARY_LIGHT)
        _textbox(slide, ex + d / 2 + 2, ey - 10, 150, 20, lab, size=size + (1 if on else 0), bold=on,
                 color=COLOR_PRIMARY if on else COLOR_TEXT, dense=True, anchor="middle", where="positioning_map")


@part("構造", "ベン図（2〜3円の重なり）")
def venn(slide, area, labels, center=None, size=13):
    x, y, w, h = area
    n = len(labels)
    cols = [COLOR_PRIMARY_MID, COLOR_PRIMARY_LIGHT, COLOR_GRAY]
    if n == 2:
        d = min(h, w / 1.6)
        cs = [(x + w / 2 - d * 0.3, y + h / 2), (x + w / 2 + d * 0.3, y + h / 2)]
        tpos = [(-0.32, 0), (0.32, 0)]
    else:
        d = min(h / 1.55, w / 1.65)
        cs = [(x + w / 2 - d * 0.3, y + d / 2), (x + w / 2 + d * 0.3, y + d / 2), (x + w / 2, y + d / 2 + d * 0.52)]
        tpos = [(-0.25, -0.18), (0.25, -0.18), (0, 0.25)]
    for i, (cx, cy) in enumerate(cs):
        sp = _rect(slide, cx - d / 2, cy - d / 2, d, d, fill=cols[i], shape=MSO_SHAPE.OVAL)
        _alpha(sp, 55)
    for i, (lab, (tx, ty)) in enumerate(zip(labels, tpos)):
        cx, cy = cs[i]
        _textbox(slide, cx + tx * d - d * 0.25, cy + ty * d - 20, d * 0.5, 40, lab, size=size, bold=True,
                 color=COLOR_PRIMARY, align="center", anchor="middle", where=f"venn[{i}]")
    if center:
        mx = sum(c[0] for c in cs) / n
        my = sum(c[1] for c in cs) / n
        _textbox(slide, mx - d * 0.24, my - 18, d * 0.48, 36, center, size=size - 1, bold=True, color=COLOR_WHITE,
                 align="center", anchor="middle", where="venn中心")


@part("構造", "ハブ＆スポーク（中心と周囲の関係者）")
def hub_spoke(slide, area, center, nodes, size=12):
    x, y, w, h = area
    cx, cy = x + w / 2, y + h / 2
    d = min(h * 0.36, 130)
    bw, bh = min(w * 0.24, 180), 44
    rx, ry = w / 2 - bw / 2 - 4, h / 2 - bh / 2 - 2
    n = len(nodes)
    for i, lab in enumerate(nodes):
        a = -math.pi / 2 + 2 * math.pi * i / n
        nx, ny = cx + rx * math.cos(a), cy + ry * math.sin(a)
        k = (d / 2) / max(math.hypot(nx - cx, ny - cy), 1)
        _line(slide, cx + (nx - cx) * k, cy + (ny - cy) * k, nx, ny, COLOR_PRIMARY_LIGHT, 2)
        _box(slide, nx - bw / 2, ny - bh / 2, bw, bh, lab, style="text", size=size, align="center",
             shape=MSO_SHAPE.ROUNDED_RECTANGLE, dense=True, where=f"hub_spoke[{i}]")
    sp = _circle(slide, cx - d / 2, cy - d / 2, d, None)
    _put_text(sp, center, size=size + 2, color=COLOR_WHITE, bold=True, align="center", anchor="middle",
              inset=(4, 4), where="hub_spoke中心")


@part("構造", "レイヤー構成図（アーキテクチャの階層）")
def layer_stack(slide, area, layers, label_w=130, size=12):
    """layers: 上から [(層の名前, [要素, ...])]"""
    x, y, w, h = area
    for i, ((lab, items), (rx, ry, rw, rh)) in enumerate(zip(layers, split_rows(area, [1] * len(layers), 6))):
        _box(slide, rx, ry, label_w, rh, lab, style="label", size=size + 1, where=f"layer_stack層[{i}]")
        _rect(slide, rx + label_w, ry, rw - label_w, rh, fill=COLOR_GRAY_PALE)
        if items:
            for j, (cx, cy, cw, ch) in enumerate(split_cols((rx + label_w + 8, ry + 6, rw - label_w - 16, rh - 12),
                                                            [1] * len(items), 8)):
                _box(slide, cx, cy, cw, ch, items[j], style="text", size=size, align="center", dense=True,
                     where=f"layer_stack要素[{i},{j}]")


@part("構造", "同心円（包含関係・範囲の広さ）")
def concentric(slide, area, labels, size=13):
    """labels: 外側から内側へ"""
    x, y, w, h = area
    n = len(labels)
    D = min(w, h)
    cols = [COLOR_PRIMARY_PALE, COLOR_PRIMARY_LIGHT, COLOR_PRIMARY_MID, COLOR_PRIMARY]
    by = y + h
    for i, lab in enumerate(labels):
        d = D * (1 - i / (n + 0.4))
        cx = x + w / 2
        _rect(slide, cx - d / 2, by - d, d, d, fill=cols[min(i, 3)], line=COLOR_WHITE, shape=MSO_SHAPE.OVAL)
        ring = D / (n + 0.4)
        last = i == n - 1
        _textbox(slide, cx - d * 0.35, by - d + (d / 2 - 18 if last else 4), d * 0.7, min(ring - 4, 36) if not last else 36,
                 lab, size=size, bold=True, color=COLOR_WHITE if i >= 2 else COLOR_PRIMARY, align="center",
                 anchor="middle", where=f"concentric[{i}]")


@part("構造", "式の形（A ＋ B ＝ C）")
def equation(slide, area, terms, result, op="+", size=14):
    x, y, w, h = area
    n = len(terms)
    cols = split_cols(area, [1] * n + [1.1], 44)
    for i, (t, (cx, cy, cw, ch)) in enumerate(zip(terms, cols)):
        _box(slide, cx, cy, cw, ch, t, style="pale", size=size, where=f"equation項[{i}]")
        sym = {"+": "＋", "x": "×", "-": "－"}.get(op, op) if i < n - 1 else "＝"
        _circle(slide, cx + cw + 8, cy + ch / 2 - 14, 28, sym, fill=COLOR_WHITE, color=COLOR_PRIMARY, size=18)
    rx, ry, rw, rh = cols[-1]
    _box(slide, rx, ry, rw, rh, result, style="label", size=size + 1, where="equation結果")


@part("構造", "VS 対比（2つの選択肢を向かい合わせる）")
def versus(slide, area, left, right, size=14):
    """left/right = (見出し, items)"""
    x, y, w, h = area
    a, b = split_cols(area, [1, 1], 50)
    titled_box(slide, a, left[0], left[1], style="primary", size=size)
    titled_box(slide, b, right[0], right[1], style="mid", size=size)
    _circle(slide, x + w / 2 - 26, y + h / 2 - 26, 52, "VS", fill=COLOR_WHITE, color=COLOR_PRIMARY, size=18,
            line=COLOR_PRIMARY)


@part("構造", "ギャップ分析（As-Is → 施策 → To-Be）")
def gap_analysis(slide, area, as_is, to_be, actions, labels=("現状（As-Is）", "あるべき姿（To-Be）"),
                 action_label="ギャップを埋める施策", size=13):
    x, y, w, h = area
    a, m, b = split_cols(area, [1, 1.05, 1], 14)
    titled_box(slide, a, labels[0], as_is, style="gray", size=size)
    titled_box(slide, b, labels[1], to_be, style="primary", size=size)
    sp = _rect(slide, m[0], m[1] + 30, m[2], m[3] - 60, fill=COLOR_PRIMARY_PALE, shape=MSO_SHAPE.RIGHT_ARROW)
    sp.adjustments[0] = 0.72
    sp.adjustments[1] = 0.28
    _textbox(slide, m[0] + 8, m[1] + 30 + (m[3] - 60) * 0.14 + 4, m[2] * 0.72 - 8, 22, action_label, size=size,
             bold=True, color=COLOR_PRIMARY, where="gap_analysis見出し")
    _textbox(slide, m[0] + 8, m[1] + 30 + (m[3] - 60) * 0.14 + 28, m[2] * 0.72 - 8, (m[3] - 60) * 0.72 - 32,
             actions, size=size, bullets=True, where="gap_analysis施策")


@part("構造", "SWOT 分析")
def swot(slide, area, s_items, w_items, o_items, t_items, size=12):
    heads = [("S：強み", s_items, "primary"), ("W：弱み", w_items, "gray"),
             ("O：機会", o_items, "mid"), ("T：脅威", t_items, "pale")]
    for (t, items, st), a in zip(heads, grid_areas(area, 2, 2, 10)):
        titled_box(slide, a, t, items, style=st, size=size)


@part("構造", "中かっこでまとめる（複数要素 → 1つの意味）")
def brace_group(slide, area, items, label, size=13, item_ratio=0.55):
    x, y, w, h = area
    iw = w * item_ratio
    for i, (rx, ry, rw, rh) in enumerate(split_rows((x, y, iw, h), [1] * len(items), 6)):
        _box(slide, rx, ry, rw, rh, items[i], style="text", size=size, align="center", where=f"brace_group[{i}]")
    br = _rect(slide, x + iw + 8, y, 22, h, line=COLOR_PRIMARY, line_w=1.5, shape=MSO_SHAPE.RIGHT_BRACE)
    br.name = "deco:brace"
    lh = min(h * 0.6, 100)
    _box(slide, x + iw + 40, y + (h - lh) / 2, w - iw - 40, lh, label, style="label", size=size + 1,
         where="brace_groupまとめ")


def _edge_point(cx, cy, bw, bh, tx, ty):
    """箱の中心 (cx,cy) から (tx,ty) へ向かう線が箱の縁と交わる点"""
    dx, dy = tx - cx, ty - cy
    if dx == 0 and dy == 0:
        return cx, cy
    k = min((bw / 2) / abs(dx) if dx else 1e9, (bh / 2) / abs(dy) if dy else 1e9)
    return cx + dx * k, cy + dy * k


@part("構造", "関係図（ノードを自由配置し矢印とラベルで結ぶ）")
def relation_map(slide, area, nodes, edges, box=(130, 44), size=11):
    """nodes: {id: (ラベル, x 0-1, y 0-1, style="text")}  edges: [(from_id, to_id, ラベル or "")]"""
    x, y, w, h = area
    bw, bh = box
    pos = {}
    for nid, spec in nodes.items():
        lab, rx, ry = spec[:3]
        pos[nid] = (x + bw / 2 + rx * (w - bw), y + bh / 2 + ry * (h - bh))
    for a, b, lab in edges:
        (ax, ay), (bx, by) = pos[a], pos[b]
        sx, sy = _edge_point(ax, ay, bw, bh, bx, by)
        ex, ey = _edge_point(bx, by, bw, bh, ax, ay)
        _arrow_line(slide, sx, sy, ex, ey, COLOR_PRIMARY_MID, 1.5)
        if lab:
            lw = min(140, max(60, len(lab) * 11 + 12))
            tb = _textbox(slide, (sx + ex) / 2 - lw / 2, (sy + ey) / 2 - 10, lw, 20, lab, size=10,
                          color=COLOR_PRIMARY_MID, align="center", anchor="middle", dense=True, where="relation_mapラベル")
            tb.fill.solid()
            tb.fill.fore_color.rgb = RGBColor.from_string(COLOR_WHITE)
    for nid, spec in nodes.items():
        lab = spec[0]
        st = spec[3] if len(spec) > 3 else "text"
        cx, cy = pos[nid]
        _box(slide, cx - bw / 2, cy - bh / 2, bw, bh, lab, style=st, size=size, align="center",
             shape=MSO_SHAPE.ROUNDED_RECTANGLE, dense=True, where=f"relation_map[{nid}]")


@part("構造", "ペルソナカード（人物像・属性・ニーズ）")
def persona_card(slide, area, name, attrs, needs, icon="人", size=12):
    x, y, w, h = area
    _rect(slide, x, y, w, h, fill=COLOR_WHITE, line=COLOR_GRAY)
    a, b, c = split_cols((x + 12, y + 12, w - 24, h - 24), [0.6, 1.2, 1.2], 16)
    d = min(a[2], 90)
    _circle(slide, a[0] + (a[2] - d) / 2, a[1] + 10, d, icon, fill=COLOR_PRIMARY_PALE, color=COLOR_PRIMARY, size=d * 0.4)
    _textbox(slide, a[0], a[1] + d + 18, a[2], 50, name, size=size + 2, bold=True, align="center", where="persona名前")
    definition_list(slide, b, attrs, term_ratio=0.35, size=size)
    titled_box(slide, c, "ニーズ・課題", needs, style="primary", size=size)


@part("構造", "役割カード（役割・担当者・責任）")
def role_cards(slide, area, roles, size=12):
    """roles: [(役割, 担当者, [責任, ...])]"""
    for i, ((role, who, duties), (x, y, w, h)) in enumerate(zip(roles, split_cols(area, [1] * len(roles)))):
        _box(slide, x, y, w, 30, role, style="label", size=size + 1, where=f"role_cards役割[{i}]")
        _rect(slide, x, y + 30, w, h - 30, fill=COLOR_WHITE, line=COLOR_PRIMARY)
        _textbox(slide, x + 8, y + 36, w - 16, 26, who, size=size + 1, bold=True, color=COLOR_PRIMARY,
                 align="center", where=f"role_cards担当[{i}]")
        _line(slide, x + 12, y + 64, x + w - 12, y + 64, COLOR_GRAY_LIGHT, LINE_W)
        _textbox(slide, x + 8, y + 70, w - 16, h - 78, duties, size=size, bullets=True, where=f"role_cards責任[{i}]")
