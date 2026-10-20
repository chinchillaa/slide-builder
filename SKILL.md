---
name: slide-builder
description: 業務用のPowerPoint資料（報告書・提案書・検討資料・会議資料などのスライド/パワポ/デッキ）を新規に作る・作り直すときに使う。python-pptx と固定デザインのヘルパー関数で、ユーザーのいつもの資料フォーマット（パンくず＋タイトル＋紺罫線＋リード文、紺ラベルのカード表、ガント、ネイティブグラフ）どおりのネイティブ pptx を生成する。構成案レビュー→人間の承認→1枚1スクリプトで実装→PowerPointで画像化→目視レビューの順に進める。「スライド作って」「資料にまとめて」「パワポで」「報告資料」などで起動する。既存 pptx の読み取りや軽微な編集だけなら汎用の pptx スキルを使う。
---

# slide-builder

このスキルの原則は「**デザインの判断を AI がしない**」こと。
配色・フォント・罫線・余白・座標の規則は `scripts/pptx_helpers.py` に固定されている。
AI がやることは **何を伝えるかを決め、どの関数をどのデータで呼ぶかを書くこと** だけ。

> **関数のデザインこそが承認済みデザインである。** 色・罫線・フォント・余白を自分の美的判断で上書きしてはならない。
> 関数が引数で受け付けない項目は意図的に固定されている。python-pptx を直接使って図形を描くのは、ヘルパーで表現できない場合に限り、その場合もパレット定数（`COLOR_*`）と `FONT` を使う。

`SKILL_DIR` はこの SKILL.md があるフォルダ（`%USERPROFILE%\.claude\skills\slide-builder`）。

## 0. 準備（初回のみ）

```
python -c "import pptx, win32com"   # 失敗したら: python -m pip install python-pptx pywin32 pillow
```

## 1. 企画のヒアリング

次の項目が分からなければユーザーに聞く。分かっていれば聞かない。
- 目的と読み手（誰が読んで、何を判断・行動するのか）
- 結論（いちばん伝えたいこと）
- 素材（議事録・データ・既存資料のパス）
- おおよその枚数、提出先（宛先）、作成者名・日付、フッター表記

## 2. 構成計画を作る（実装前）

デッキフォルダを作る。ユーザーの指定がなければ、作業ディレクトリに `slides/<テーマ名>/` を作る。

```
slides/<theme>/
  plan.md          構成計画
  deck.json        {"name": "<出力ファイル名>", "footer": "<右下表記 例: ©2026 ○○部>"}
  python/          slide00_cover.py, slide01_toc.py, ... （1枚1スクリプト）
  <name>.pptx      生成物（毎回作り直す）
  slides/          レビュー用 PNG
```

`plan.md` には全スライドを表で書く:

| No | 種別 | パンくず | タイトル | リード文（2行以内・結論を書く） | 部品 | 本文の骨子 |
|---|---|---|---|---|---|---|

- 種別: cover / toc / section / content
- 部品は `reference/patterns.md` の型から選び、必要に応じて `reference/parts.md` で個別の部品を探す（例: `card_table(label,label2,text→text)`、`split_columns 1:1 + bullets/message_box`）
- リード文はそのスライドの結論にする（「〜について説明する」ではなく「〜のため、〜とした」）
- 1スライドに1メッセージ。情報が多いときはスライドを分ける

## 3. 構成レビュー → 人間の承認（承認前に実装しない）

1. 作成の経緯を知らない **新しいサブエージェント**（Agent ツール、general-purpose）に、`SKILL_DIR/reviewers/plan_review.md` の指示と `plan.md` のパスを渡してレビューさせる
2. 指摘を反映したうえで、**plan.md の表と主な変更点をユーザーに示し、承認を待つ**
3. ユーザーが修正を求めたら直して再提示する。承認されるまで python/ を書き始めない

## 4. 実装（1スライド = 1スクリプト）

各ファイルに `build(prs)` を定義する。ファイル名の順にスライドが追加される。**書く前に、使う部品の実例（`reference/parts.md` の「実例」列のスクリプト）を読む**。デッキ全体の例は `SKILL_DIR/examples/catalog/`。

```python
from pptx_helpers import *

def build(prs):
    s = content_slide(prs, "2.1. 課題と対応方針", breadcrumb="2. 現状分析",
                      lead="ヒアリング結果から主に3つの課題が見つかったため、それぞれに対応方針を定めた。")
    areas, _ = split_columns(s, BODY, [1, 2], titles=["背景", "課題分析"], arrow=True)
    bullets(s, areas[0], ["…", (1, "サブ項目")], box=True)
    card_table(s, areas[1], ...)
    source(s, "○○調査（2026年）")
```

ビルド:
```
python "SKILL_DIR/scripts/build_deck.py" slides/<theme>
```
- `ValueError`（フォント下限・はみ出し）は **内容側で解決する**: 文字数を減らす / 枠や比率を変える / スライドを分ける。ヘルパーの改変やサイズ下げで回避しない
- `[ERROR]` は必ず直す。`[WARN]`（重なり・パレット外の色）は確認して、意図したものでなければ直す
- 「手で編集された可能性がある」と言われて止まったら、**ユーザーに確認してから** `--force` を付ける（手修正が消えるため）

## 5. 画像化して目視レビュー

```
python "SKILL_DIR/scripts/render.py" slides/<theme>/<name>.pptx            # 全枚
python "SKILL_DIR/scripts/render.py" slides/<theme>/<name>.pptx --slides 3,5  # 修正した分だけ
```
1. **新しいサブエージェント** に `SKILL_DIR/reviewers/visual_review.md` の指示と、PNG フォルダ・plan.md のパスを渡す（画像はサブエージェントに読ませ、メインには所見だけを返させる）
2. 返ってきた Must fix / Nice to fix をユーザーに示す。Must fix は直す。Nice to fix はユーザーに選んでもらう
3. 該当する slideNN の .py だけを直し、ビルドしてから、変更したスライドだけを再度画像化して確認する

## 6. 引き渡し

pptx のパスと枚数を伝える。この後はユーザーが PowerPoint で手修正する前提でよい。その後に AI でスクリプトを直して作り直すと手修正が消えることも、あわせて伝える。

---

## 部品（全 167 種）と型

部品の一覧・引数・実例は **`SKILL_DIR/reference/parts.md`**（`scripts/list_parts.py` で自動生成）を読む。
スライドの型（目的 → 部品の組み合わせ）は **`SKILL_DIR/reference/patterns.md`** を読む。
見た目は `SKILL_DIR/examples/parts/slides/*.png`（全部品のカタログ画像）で確認できる。

| カテゴリ | 数 | 主な部品 |
|---|---|---|
| スライド枠 | 9 | cover_slide, cover_band, toc_slide, agenda_slide, section_slide, section_split, content_slide, closing_slide |
| レイアウト | 10 | split_columns, split_cols, split_rows, grid_areas, band_heading, sub_heading, side_label, section_tab |
| テキスト | 18 | bullets, message_box, numbered_list, statement, callout, quote_box, qa_list, checklist, exec_points, spec_text, code_block, tag |
| カード | 6 | titled_box, card_grid, point_cards, pros_cons, before_after, case_card |
| 表 | 15 | card_table, make_table, comparison_table, score_table, harvey_table, raci_table, heatmap_table, status_table, mark_matrix, bar_table, action_items |
| フロー | 17 | step_flow, chevron_row, flow_boxes, swimlane, cycle, funnel, pyramid, timeline, roadmap, phase_gate, step_matrix |
| 構造 | 17 | logic_tree, count_tree, org_chart, matrix_2x2, positioning_map, venn, hub_spoke, layer_stack, gap_analysis, swot, relation_map |
| 図解 | 5 | zone, business_model, flow_before_after, issue_map, pipeline |
| 数値 | 10 | kpi, kpi_row, kpi_trend, stat_compare, big_number, progress_bars, rank_bars, percent_bar, donut_kpi, reduction_bars |
| グラフ | 5 | chart（column/bar/line/area/pie/doughnut/100%/radar…）, waterfall, scatter_chart, bubble_chart, chart_with_insight |
| スケジュール | 4 | gantt, calendar_month, schedule_list, chronology |
| 画像 | 5 | image, image_caption, image_grid, screenshot_callouts, image_placeholder |
| アイコン | 3 | icon（person/people/document/pc/server/database/cloud/building/factory/money/phone/mail/gear/lock/chart/check/warning/clock の18種）, icon_label, icon_steps |
| 線・矢印 | 8 | connect（solid/dashed/dotted/thick/both/service/money/goods/info）, elbow_connect, curved_arrow, loop_arrow, block_arrow, flow_legend, annotate_points, span_bracket |
| 補助 | 7 | source, arrow, legend, number_badge, highlight_frame, footnotes, ref_tag |
| 会議・報告 | 7 | traffic_light, decision_log, meeting_bodies, weekly_report, meeting_agenda, revision_history, notice_box |
| リスク・課題 | 3 | risk_matrix, risk_table, issue_tracker |
| 費用・効果 | 4 | cost_table, roi_chart, budget_actual_table, effect_breakdown |
| 要件・仕様 | 2 | requirements_table, grouped_table |
| 目標管理 | 3 | kpi_tree, okr, target_actual |
| 顧客・利用者 | 3 | customer_journey, likert_chart, empathy_map |
| 意思決定 | 3 | recommendation, assumptions, plan_cards |
| フレームワーク | 3 | framework_grid（3c/4p/4c/pest/kpt）, five_forces, business_model_canvas |

基本の寸法: 単位はすべて pt、スライドは 960×540、`area = (x, y, w, h)`、本文領域は `BODY = (34, 118, 892, 378)`、間隔は `GAP = 12`。
テキスト内のマークアップ: `**太字**`、`!!赤の強調!!`（1スライドに1〜2か所まで）、`__下線__`、改行は `\n`。箇条書きの階層は `(1, "サブ項目")`。
図の中の箱どうしは座標を計算して線を引かず、`connect(slide, 箱A, 箱B, kind=...)` でつなぐ（縁から縁へ自動で配線される）。
画像のパスは `os.path.join(os.path.dirname(__file__), "..", "assets", "xxx.png")` のようにスクリプトからの相対で書き、画像はデッキの `assets/` に置く。

### デザイン規則（コードで強制されているもの）
- フォントは Meiryo UI。パンくず 14pt 太字、タイトル 24pt 太字、リード 16pt、本文 14pt（下限 12pt）、表・グラフ・図の注記は下限 10pt
- 色はパレット定数（`COLOR_PRIMARY` 紺、`COLOR_PRIMARY_MID` 青、`COLOR_PRIMARY_LIGHT`、`COLOR_PRIMARY_PALE` 淡青、グレー系、`COLOR_ACCENT` 赤＝文字の強調のみ、`COLOR_STATUS_OK`/`COLOR_STATUS_WARN`＝信号表示のみ）だけを使う。hex は直接書かない
- 次の場合はビルドエラーになる: 枠から溢れるテキスト、英単語や数値が途中で改行される幅、スライドの外に出る図形
- 大きな数字（kpi・big_number など）は1行に収まるよう自動で縮小される

### 部品を増やすとき
既存の部品で表現できない型が繰り返し出てきたら、`scripts/parts_*.py` の該当カテゴリに `@part("カテゴリ", "説明")` を付けて関数を追加する。続けて `examples/parts/python/` に実例を足し、ビルド → 画像で確認したうえで `python scripts/list_parts.py` を実行して一覧を更新する。部品の中でも、色・フォントは定数を使い、文字は `_put_text` / `_textbox` を通す（溢れ検査を効かせるため）。
