# 部品一覧（自動生成: scripts/list_parts.py）

全 167 部品。`from pptx_helpers import *` で使える。単位は pt、area は (x, y, w, h)。
見た目は `examples/parts/slides/`（カタログ画像）、使い方の実例は「実例」列のスクリプトを読む。

## スライド枠（9）

| 関数 | 用途 | 引数 | 実例 |
|---|---|---|---|
| `cover_slide` | 表紙<br>表紙。title は最大2行（28pt） | `(prs, title, subtitle='', date_text='', org='', recipient='')` | examples/catalog/python/slide00_cover.py |
| `section_slide` | 章の区切り（紺の全面塗り）<br>セクション区切り（紺ベタ）。items は配下の見出し一覧 | `(prs, title, items=None)` | examples/catalog/python/slide02_section.py |
| `content_slide` | 本文スライドの枠（パンくず・タイトル・罫線・リード文）<br>本文スライド。パンくず(14pt太字) / タイトル(24pt太字・1行) / 紺罫線 / リード文(16pt・2行以内)。 nav=(tabs, active_index) で右上にナビタブ。本文は BODY 領域に置く。 | `(prs, title, lead='', breadcrumb='', nav=None)` | examples/catalog/python/slide03_split.py<br>examples/catalog/python/slide04_gantt.py |
| `nav_tabs` | 右上のナビタブ（現在地の表示）<br>右上のナビタブ（現在地表示）。占有幅を返す | `(slide, tabs, active)` | examples/parts/python/slide15_misc.py |
| `toc_slide` | 目次（2段組まで自動）<br>目次。items: [(level, text, page), ...]  level 0=章(太字) / 1=節。15行超で2段組 | `(prs, items, title='目次')` | examples/catalog/python/slide01_toc.py |
| `agenda_slide` | アジェンダ（現在の項目を強調） | `(prs, items, current=None, title='アジェンダ')` | examples/parts/python/slide70_slides.py |
| `closing_slide` | 締めのスライド（以上／問い合わせ先） | `(prs, message='以上', contact=None)` | examples/parts/python/slide70_slides.py |
| `cover_band` | 表紙（紺の帯に白抜きタイトル） | `(prs, title, subtitle='', date_text='', org='', recipient='')` | examples/parts/python/slide9i_frames.py |
| `section_split` | 章の区切り（左に紺のパネル、右に節の一覧） | `(prs, title, items=None)` | examples/parts/python/slide9i_frames.py |

## レイアウト（10）

| 関数 | 用途 | 引数 | 実例 |
|---|---|---|---|
| `column_header` | 列見出し（下線付き）<br>列見出し（中央寄せテキスト + 下線）。消費した高さを返す | `(slide, x, y, w, text, size=14)` | examples/parts/python/slide15_misc.py |
| `split_columns` | 左右分割（見出し・列間矢印は任意）<br>領域を左右に分割し、各列の本文領域を返す。 titles で列見出し、arrow=True で列間にグレー三角矢印。 戻り値: (areas, arrow_areas)  ※見出し位置は全列で必ず揃う | `(slide, area, ratios, titles=None, arrow=False)` | examples/catalog/python/slide03_split.py<br>examples/catalog/python/slide07_chart.py |
| `split_rows` | 上下分割（描画なし。領域を返す） | `(area, ratios, gap=12)` | examples/parts/python/slide11_text_b.py<br>examples/parts/python/slide13_cards_b.py |
| `grid_areas` | 格子状に分割（描画なし。行優先で領域を返す） | `(area, rows, cols, gap=12)` | examples/parts/python/slide15_misc.py<br>examples/parts/python/slide80_icons.py |
| `split_cols` | 左右に等分・比率分割（描画なし。領域を返す） | `(area, ratios, gap=12)` | examples/parts/python/slide10_text_a.py<br>examples/parts/python/slide11_text_b.py |
| `side_label` | 縦書きの帯ラベル（左端の分類など） | `(slide, area, label, style='label', size=14)` | examples/parts/python/slide14_cards_c.py |
| `band_heading` | 紺の帯見出し（セクション内の大見出し）。消費高さを返す | `(slide, x, y, w, label, size=14)` | examples/parts/python/slide12_cards.py |
| `sub_heading` | ■付きの小見出し＋細罫線。消費高さを返す | `(slide, x, y, w, label, size=14)` | examples/parts/python/slide12_cards.py |
| `divider` | 区切り線 | `(slide, x, y, w, dash=False)` | examples/parts/python/slide12_cards.py |
| `section_tab` | セクションタブ（本文の先頭に「4.1 システムの基本機能」の青タブ）。残りの本文領域を返す | `(slide, label, area=None, size=13)` | examples/parts/python/slide84_zone.py |

## テキスト（18）

| 関数 | 用途 | 引数 | 実例 |
|---|---|---|---|
| `bullets` | 箇条書き（枠・小見出しは任意）<br>箇条書き。title で紺太字の小見出し、box=True でグレー枠 | `(slide, area, items, size=14, title=None, box=False, anchor='top')` | examples/catalog/python/slide03_split.py<br>examples/parts/python/slide12_cards.py |
| `text` | 地の文<br>箇条書きにしない地の文 | `(slide, area, body, size=14, bold=False, align='left', anchor='top', color='000000')` | examples/parts/python/slide15_misc.py<br>examples/parts/python/slide81_lines.py |
| `message_box` | 結論・キーメッセージの強調ボックス<br>結論・キーメッセージの強調ボックス。style: primary(紺ベタ) / pale(淡青) / outline(紺枠) | `(slide, area, body, style='primary', size=16)` | examples/catalog/python/slide03_split.py<br>examples/parts/python/slide15_misc.py |
| `numbered_list` | 番号付きリスト（丸数字＋見出し＋説明）<br>items: ["文", ...] または [(見出し, 説明), ...] | `(slide, area, items, size=14)` | examples/parts/python/slide10_text_a.py |
| `callout` | 吹き出し（指し示す位置を tail で指定）<br>tail=(x, y) 吹き出しの先端の座標 | `(slide, area, body, tail, size=12)` | examples/parts/python/slide10_text_a.py |
| `note_box` | 補足ボックス（グレー地＋「補足」タグ） | `(slide, area, items, label='補足', size=12)` | examples/parts/python/slide10_text_a.py |
| `quote_box` | 引用・ヒアリングの声（発言者付き） | `(slide, area, quote, speaker=None, size=14)` | examples/parts/python/slide10_text_a.py |
| `statement` | 大きな一文のステートメント（上下罫線） | `(slide, area, body, sub=None, size=24)` | examples/parts/python/slide11_text_b.py |
| `tag` | 小さなタグ・バッジ（幅は自動）。幅を返す | `(slide, x, y, label, style='primary', size=11, h=20)` | examples/parts/python/slide10_text_a.py |
| `qa_list` | Q&A の一覧 | `(slide, area, pairs, size=13)` | examples/parts/python/slide11_text_b.py |
| `definition_list` | 用語定義の一覧（用語＋説明） | `(slide, area, pairs, term_ratio=0.25, size=13)` | examples/parts/python/slide11_text_b.py |
| `checklist` | チェックリスト（☑/□） | `(slide, area, items, checked=None, size=14)` | examples/parts/python/slide11_text_b.py |
| `conclusion_bar` | 結論バー（左に「結論」ラベル＋淡青の本文） | `(slide, area, body, label='結論', size=16)` | examples/parts/python/slide13_cards_b.py |
| `icon_rows` | アイコン丸＋見出し＋説明の行リスト<br>rows: [(アイコン文字1-2字, 見出し, 説明)] | `(slide, area, rows, size=13)` | examples/parts/python/slide14_cards_c.py |
| `summary_box` | まとめボックス（紺枠＋左上タイトルタブ） | `(slide, area, items, title='まとめ', size=14)` | examples/parts/python/slide14_cards_c.py |
| `exec_points` | エグゼクティブサマリー（番号＋見出し＋説明の行） | `(slide, area, points, size=14)` | examples/parts/python/slide14_cards_c.py |
| `spec_text` | 仕様書型の本文（①見出し＋段落、i./ii. と イ)ロ) の階層、__下線__で強調）<br>sections: [(見出し, 本文)]  本文は str（段落）か list（(1,"..") で i. ii.、(2,"..") で イ) ロ)） | `(slide, area, sections, size=13, numbered=True)` | examples/parts/python/slide86_spec.py |
| `code_block` | コード・データ例のブロック（等幅フォント。JSON・設定例など） | `(slide, area, code, title=None, size=11)` | examples/parts/python/slide86_spec.py |

## カード（6）

| 関数 | 用途 | 引数 | 実例 |
|---|---|---|---|
| `titled_box` | 見出し帯付きのボックス（汎用カード） | `(slide, area, title, items, style='primary', size=14, head_h=30)` | examples/parts/python/slide15_misc.py |
| `card_grid` | カードを格子状に並べる | `(slide, area, cards, cols=3, style='primary', size=13)` | examples/parts/python/slide12_cards.py |
| `point_cards` | ポイントカード（POINT 番号＋見出し＋説明） | `(slide, area, cards, label='POINT', size=13)` | examples/parts/python/slide13_cards_b.py |
| `pros_cons` | メリット／デメリットの対比 | `(slide, area, pros, cons, titles=('メリット', 'デメリット'), size=14)` | examples/parts/python/slide13_cards_b.py |
| `before_after` | Before → After の対比（矢印付き） | `(slide, area, before, after, labels=('Before', 'After'), size=14)` | examples/parts/python/slide14_cards_c.py |
| `case_card` | 事例紹介カード（画像・概要・ポイント・出所） | `(slide, area, title, overview, points, image_path=None, label=None, src=None, size=12)` | examples/parts/python/slide89_case.py |

## 表（15）

| 関数 | 用途 | 引数 | 実例 |
|---|---|---|---|
| `card_table` | カード型の表（紺ラベル＋本文、結合・矢印あり）<br>カード型の表（ラベル列 + 本文列）。資料で最も多用する部品。 rows: [[cell, ...], ...]  cell は str / list(箇条書き) / None(上のセルと結合: label 系のみ) styles: 列ごとのスタイル "label"(紺) "label2"(青) "pale"(淡青) "gray" "text"(枠+箇条書き) "plain" headers: 列見出し（None で省略、要素に "" で個別省略） arrow_before: この列番号の手前にグレー三角矢印を置く（例: (3,)） | `(slide, area, rows, col_ratios, styles, headers=None, arrow_before=(), row_ratios=None, size=14)` | examples/catalog/python/slide05_card.py |
| `make_table` | ネイティブの表（データの多い一覧）<br>PowerPoint ネイティブの表（編集可能）。データ量の多い一覧向け。 align: 列ごと "l"/"c"/"r"。cell は str か list(箇条書き)。 stretch=True で行高を広げて領域を埋める。収まらなければ ValueError | `(slide, area, header, rows, col_ratios=None, size=12, align=None, first_col_label=False, zebra=False, stretch=False)` | examples/catalog/python/slide06_table.py |
| `comparison_table` | 比較評価表（◎○△×を色付きで表示、推奨列を強調）<br>cells[行=評価項目][列=選択肢]。セル先頭の ◎○△× は自動で色付け（例: "◎ 低コスト"） | `(slide, area, criteria, options, cells, recommend=None, col_ratios=None, size=12, stretch=False)` | examples/parts/python/slide20_table_a.py |
| `matrix_table` | マトリクス表（行見出し×列見出し、カード型） | `(slide, area, row_labels, col_labels, cells, corner='', label_ratio=0.8, size=13)` | examples/parts/python/slide21_table_b.py |
| `score_table` | スコア表（●○の点数表示＋合計、重み付け可） | `(slide, area, criteria, options, scores, max_score=5, weights=None, size=12, stretch=False)` | examples/parts/python/slide20_table_a.py |
| `harvey_table` | ハーベイボール表（充足度を ○◔◑◕● で表示）<br>values[行][列] = 0..4 | `(slide, area, row_labels, col_labels, values, label_ratio=1.6, size=12)` | examples/parts/python/slide21_table_b.py |
| `raci_table` | RACI 表（役割分担 R/A/C/I を色分け） | `(slide, area, tasks, roles, matrix, size=12, stretch=False)` | examples/parts/python/slide22_table_c.py |
| `key_value_table` | 項目／内容の2列表（概要表・諸元表） | `(slide, area, pairs, key_ratio=0.25, size=13, stretch=False)` | examples/parts/python/slide22_table_c.py |
| `heatmap_table` | ヒートマップ表（値の大小を5段階の青で塗る） | `(slide, area, row_labels, col_labels, values, fmt='{:,.0f}', size=12, stretch=False)` | examples/parts/python/slide23_table_d.py |
| `status_table` | 進捗・ステータス表（完了/進行中/未着手/遅延/保留を色分け） | `(slide, area, header, rows, status_col, col_ratios=None, size=12, stretch=False)` | examples/parts/python/slide23_table_d.py |
| `check_table` | 対応表（✓ マーク。機能×対象など）<br>marks[行][列] は True/False または文字列 | `(slide, area, row_labels, col_labels, marks, size=12, stretch=False)` | examples/parts/python/slide23_table_d.py |
| `highlight_table` | 行を強調できる一覧表（強調行は淡青＋太字） | `(slide, area, header, rows, highlight_rows=(), col_ratios=None, align=None, size=12, stretch=False)` | examples/parts/python/slide23_table_d.py |
| `action_items` | 宿題・アクションアイテム一覧（内容・担当・期限・状況）<br>items: [(内容, 担当, 期限, 状況)]  状況は 完了/進行中/未着手/遅延/保留 | `(slide, area, items, size=12, stretch=False)` | examples/parts/python/slide62_schedule_b.py |
| `mark_matrix` | ★/✓ マトリクス（項目×企業など。列グループ見出し・凡例付き）<br>marks[行][列] = "★" / "✓" / "○" / "" など。groups: [(グループ名, 列数)] で上段に見出し。legend: {"★": "注力", ...} | `(slide, area, row_labels, col_labels, marks, groups=None, legend=None, size=12)` | examples/parts/python/slide87_tables2.py |
| `bar_table` | 表＋行内の横棒（数値列を棒で見せる。国別比較など）<br>rows の bar_col 列は数値。その列のセルに横棒と値を描く | `(slide, area, header, rows, bar_col, unit='', fmt='{:,.0f}', col_ratios=None, size=12)` | examples/parts/python/slide87_tables2.py |

## フロー（17）

| 関数 | 用途 | 引数 | 実例 |
|---|---|---|---|
| `step_flow` | シェブロン見出し＋説明枠のステップ<br>横向きのステップ（シェブロン見出し + 各ステップの説明枠）。 steps: [(見出し, items), ...]  active を指定するとそれ以外の見出しをグレーにする | `(slide, area, steps, size=14, active=None, header_h=40)` | examples/catalog/python/slide08_steps.py |
| `chevron_row` | シェブロンの横一列（プロセス概要・現在地の表示） | `(slide, area, labels, active=None, size=14)` | examples/parts/python/slide30_flow_a.py |
| `vertical_steps` | 縦のステップ（番号丸を線でつなぐ）<br>steps: [(見出し, 説明)] | `(slide, area, steps, size=13)` | examples/parts/python/slide30_flow_a.py |
| `flow_boxes` | 箱を矢印でつなぐ横フロー<br>items: ["文", ...] または [(見出し, 説明), ...]。style は card_table と同じ | `(slide, area, items, style='text', size=13, gap=34)` | examples/parts/python/slide30_flow_a.py |
| `swimlane` | スイムレーン（担当別の業務フロー）<br>steps: [(lane, col, テキスト)]  arrows: [(from_step_index, to_step_index)] | `(slide, area, lanes, n_cols, steps, arrows=(), col_headers=None, lane_w=90, size=11)` | examples/parts/python/slide31_flow_b.py |
| `cycle` | 循環サイクル（PDCA など） | `(slide, area, labels, center=None, size=13)` | examples/parts/python/slide32_flow_c.py |
| `funnel` | ファネル（段階的な絞り込み）<br>stages: [(ラベル, 値テキスト)] | `(slide, area, stages, size=13, min_ratio=0.35)` | examples/parts/python/slide32_flow_c.py |
| `pyramid` | ピラミッド（階層構造＋右に説明）<br>levels: 上から [(ラベル, 説明)] | `(slide, area, levels, size=13, ratio=0.45)` | examples/parts/python/slide32_flow_c.py |
| `stair_steps` | 階段（段階的な成長・ステップアップ）<br>steps: [(見出し, 説明)] 左から右へ上がっていく | `(slide, area, steps, size=12)` | examples/parts/python/slide33_flow_d.py |
| `timeline` | タイムライン（横軸に出来事を上下交互に配置）<br>events: [(日付ラベル, 見出し, 説明)] | `(slide, area, events, size=12)` | examples/parts/python/slide33_flow_d.py |
| `roadmap` | ロードマップ（フェーズ×領域に施策を配置）<br>items: [(lane, 開始phase, 終了phase, テキスト, style="primary"|"mid"|"pale", sub=0)] | `(slide, area, phases, lanes, items, lane_w=120, size=12)` | examples/parts/python/slide34_flow_e.py |
| `converge` | 集約（複数の要素 → 1つの結果） | `(slide, area, inputs, output, size=13, out_ratio=0.35)` | examples/parts/python/slide35_flow_f.py |
| `diverge` | 展開（1つの要素 → 複数の結果） | `(slide, area, source_text, outputs, size=13, src_ratio=0.35)` | examples/parts/python/slide35_flow_f.py |
| `ipo` | Input → Process → Output | `(slide, area, inputs, process, outputs, labels=('Input', 'Process', 'Output'), size=13)` | examples/parts/python/slide30_flow_a.py |
| `phase_gate` | フェーズゲート（フェーズの間に判定ゲート）<br>phases: [(フェーズ名, items)]  gates: フェーズ間のゲート名（len = フェーズ数-1） | `(slide, area, phases, gates, size=12)` | examples/parts/python/slide35_flow_f.py |
| `value_chain` | バリューチェーン（支援活動の帯＋主活動のシェブロン） | `(slide, area, primary, support=(), size=12)` | examples/parts/python/slide36_flow_g.py |
| `step_matrix` | ステップ×観点のマトリクス（シェブロン見出し）<br>cells[行][ステップ] は str か list | `(slide, area, steps, row_labels, cells, label_w=110, size=11)` | examples/parts/python/slide36_flow_g.py |

## 構造（17）

| 関数 | 用途 | 引数 | 実例 |
|---|---|---|---|
| `logic_tree` | ロジックツリー（左から右へ分解）<br>root = (ラベル, [子, ...])  子は str または (ラベル, [孫, ...])。4階層まで | `(slide, area, root, size=12, col_ratios=None)` | examples/parts/python/slide40_struct_a.py |
| `org_chart` | 組織図・体制図（上から下へ）<br>root = (ラベル, [子, ...])  ラベルに "\n" で役職と氏名を2行にできる | `(slide, area, root, size=12)` | examples/parts/python/slide41_struct_b.py |
| `matrix_2x2` | 2×2 マトリクス（軸付きの4象限）<br>x_axis/y_axis = (軸名, 低側ラベル, 高側ラベル)。quadrants = [左上, 右上, 左下, 右下]（str か (見出し, 説明)） | `(slide, area, x_axis, y_axis, quadrants, highlight=None, size=13)` | examples/parts/python/slide42_struct_c.py |
| `positioning_map` | ポジショニングマップ（2軸に項目をプロット）<br>x_axis/y_axis = (軸名, 低側, 高側)。points = [(x 0-1, y 0-1, ラベル, 強調bool)] | `(slide, area, x_axis, y_axis, points, size=11)` | examples/parts/python/slide42_struct_c.py |
| `venn` | ベン図（2〜3円の重なり） | `(slide, area, labels, center=None, size=13)` | examples/parts/python/slide43_struct_d.py |
| `hub_spoke` | ハブ＆スポーク（中心と周囲の関係者） | `(slide, area, center, nodes, size=12)` | examples/parts/python/slide43_struct_d.py |
| `layer_stack` | レイヤー構成図（アーキテクチャの階層）<br>layers: 上から [(層の名前, [要素, ...])] | `(slide, area, layers, label_w=130, size=12)` | examples/parts/python/slide44_struct_e.py<br>examples/parts/python/slide82_annotate.py |
| `concentric` | 同心円（包含関係・範囲の広さ）<br>labels: 外側から内側へ | `(slide, area, labels, size=13)` | examples/parts/python/slide43_struct_d.py |
| `equation` | 式の形（A ＋ B ＝ C） | `(slide, area, terms, result, op='+', size=14)` | examples/parts/python/slide44_struct_e.py |
| `versus` | VS 対比（2つの選択肢を向かい合わせる）<br>left/right = (見出し, items) | `(slide, area, left, right, size=14)` | examples/parts/python/slide44_struct_e.py |
| `gap_analysis` | ギャップ分析（As-Is → 施策 → To-Be） | `(slide, area, as_is, to_be, actions, labels=('現状（As-Is）', 'あるべき姿（To-Be）'), action_label='ギャップを埋める施策', size=13)` | examples/parts/python/slide45_struct_f.py |
| `swot` | SWOT 分析 | `(slide, area, s_items, w_items, o_items, t_items, size=12)` | examples/parts/python/slide46_struct_g.py |
| `brace_group` | 中かっこでまとめる（複数要素 → 1つの意味） | `(slide, area, items, label, size=13, item_ratio=0.55)` | examples/parts/python/slide45_struct_f.py |
| `relation_map` | 関係図（ノードを自由配置し矢印とラベルで結ぶ）<br>nodes: {id: (ラベル, x 0-1, y 0-1, style="text")}  edges: [(from_id, to_id, ラベル or "")] | `(slide, area, nodes, edges, box=(130, 44), size=11)` | examples/parts/python/slide46_struct_g.py |
| `persona_card` | ペルソナカード（人物像・属性・ニーズ） | `(slide, area, name, attrs, needs, icon='人', size=12)` | examples/parts/python/slide47_struct_h.py |
| `role_cards` | 役割カード（役割・担当者・責任）<br>roles: [(役割, 担当者, [責任, ...])] | `(slide, area, roles, size=12)` | examples/parts/python/slide47_struct_h.py |
| `count_tree` | 件数付きの分類ツリー（全体 → 類型ごとの件数）<br>root = (ラベル, 件数, [子, ...])  子も同じ形。4階層まで | `(slide, area, root, unit='件', size=12)` | examples/parts/python/slide88_count.py |

## 図解（5）

| 関数 | 用途 | 引数 | 実例 |
|---|---|---|---|
| `zone` | ゾーン（点線枠で領域をまとめ、左上にラベル）。内側の領域を返す<br>style: dashed（点線・塗りなし） / pale（淡青の塗り） / gray（薄グレーの塗り） / solid（紺の実線） | `(slide, area, label, style='dashed', size=12)` | examples/parts/python/slide84_zone.py |
| `business_model` | ビジネスモデル図（関係者をアイコンで置き、お金・サービス・モノ・情報の流れでつなぐ）<br>actors: {id: (ラベル, 列, 行, icon_kind)}  列・行は 0 始まりの格子位置 flows: [(from_id, to_id, 種類, ラベル)] 種類は service / money / goods / info / solid / dashed … columns: 列見出し（例: ["ユーザー", "事業者", "メーカー"]） | `(slide, area, actors, flows, columns=None, legend=True, size=11)` | examples/parts/python/slide83_zone_bm.py |
| `flow_before_after` | 業務フローの Before / After（縦の手順を左右に並べ、変わった手順を強調）<br>before/after: 手順の list。changed: after 側で変わった手順の index | `(slide, area, before, after, changed=(), labels=('これまで', 'これから'), size=12)` | examples/parts/python/slide85_diag.py |
| `issue_map` | 論点マップ（全論点を並べ、いま説明している論点だけを強調）<br>items: 論点の list（(番号, 文) でも可）。current: 強調する index。overview: 左に置く調査概要（任意） | `(slide, area, items, current, overview=None, size=12)` | examples/parts/python/slide85_diag.py |
| `pipeline` | パイプライン（アイコンのノードを矢印とラベルでつなぐ。データの流れなど）<br>nodes: [(icon_kind, ラベル)]  edge_labels: ノード間の矢印に添える文字（len = ノード数-1） | `(slide, area, nodes, edge_labels=None, size=12)` | examples/parts/python/slide9h_pipeline.py |

## 数値（10）

| 関数 | 用途 | 引数 | 実例 |
|---|---|---|---|
| `kpi` | 大きな数字のKPIカード<br>大きな数値で示す KPI カード | `(slide, area, value, label, note=None, unit='')` | examples/catalog/python/slide08_steps.py |
| `kpi_row` | KPI カードを横に並べる<br>items: [(値, ラベル, 注記 or None)] | `(slide, area, items, unit='')` | examples/parts/python/slide50_data_a.py |
| `stat_compare` | 前後比較の数値（Before → After と増減）<br>before/after = (ラベル, 値)。delta は "▲40%" などの表示文字列 | `(slide, area, before, after, delta=None, unit='', size=36)` | examples/parts/python/slide50_data_a.py |
| `progress_bars` | 進捗バー（％）と目標線<br>items: [(ラベル, 0-100)]  target: 目標値(%) | `(slide, area, items, target=None, label_ratio=0.3, size=12)` | examples/parts/python/slide51_data_b.py |
| `rank_bars` | ランキングの横棒（図形で描く。強調行あり）<br>items: [(ラベル, 値)] 上から順に表示 | `(slide, area, items, unit='', highlight=None, fmt='{:,.0f}', label_ratio=0.25, size=12)` | examples/parts/python/slide51_data_b.py |
| `percent_bar` | 100% 帯グラフ（構成比を1本の帯で）<br>segments: [(ラベル, 割合%)] | `(slide, area, segments, size=12)` | examples/parts/python/slide51_data_b.py |
| `big_number` | 大きな数字＋説明文（インパクトのある1つの数値） | `(slide, area, number, unit='', desc='', label=None, size=64)` | examples/parts/python/slide50_data_a.py |
| `donut_kpi` | ドーナツ型の達成率（中央に％） | `(slide, area, pct, label, size=12)` | examples/parts/python/slide51_data_b.py |
| `kpi_trend` | 推移線付きの KPI カード（小さな折れ線）<br>values: 推移の数値リスト（古い順） | `(slide, area, value, label, values, unit='', note=None)` | examples/parts/python/slide62_schedule_b.py |
| `reduction_bars` | 絞り込みの棒（407件 → 293件 → 3件 のように段階的に減る過程）<br>steps: [(ラベル, 値, 絞り込み条件 or None)]  条件は次の棒への矢印に添える | `(slide, area, steps, unit='件', size=12)` | examples/parts/python/slide88_count.py |

## グラフ（5）

| 関数 | 用途 | 引数 | 実例 |
|---|---|---|---|
| `chart` | ネイティブグラフ（棒/積上/横棒/折れ線/面/円/ドーナツ/100%積上）<br>PowerPoint ネイティブグラフ（編集可能）。 kind: column / column_stacked / bar / bar_stacked / line / pie / doughnut series: {"系列名": [値, ...]}（順序 = 色順 SERIES_COLORS） highlight: 単一系列の棒で強調するカテゴリ index（他は薄青） | `(slide, area, kind, categories, series, title=None, unit=None, number_format='#,##0', labels=None, legend=None, highlight=None, size=12)` | examples/catalog/python/slide07_chart.py<br>examples/parts/python/slide53_data_d.py |
| `waterfall` | ウォーターフォール（増減の積み上げ。ネイティブ）<br>start=(ラベル, 値)  steps=[(ラベル, 増減)] | `(slide, area, start, steps, end_label='合計', unit=None, title=None, size=12, number_format='#,##0')` | examples/parts/python/slide52_data_c.py |
| `scatter_chart` | 散布図（ネイティブ）<br>series: {"系列名": [(x, y), ...]} | `(slide, area, series, x_title=None, y_title=None, size=12)` | examples/parts/python/slide52_data_c.py |
| `bubble_chart` | バブルチャート（ネイティブ。3つ目の値を円の大きさで）<br>series: {"系列名": [(x, y, 大きさ), ...]} | `(slide, area, series, x_title=None, y_title=None, size=12)` | examples/parts/python/slide52_data_c.py |
| `chart_with_insight` | グラフ＋右に示唆ボックス<br>chart() の引数に加えて insights（箇条書き）を受け取る | `(slide, area, kind, categories, series, insights, insight_title='示唆', ratio=0.62, **kw)` | examples/parts/python/slide53_data_d.py |

## スケジュール（4）

| 関数 | 用途 | 引数 | 実例 |
|---|---|---|---|
| `gantt` | ガントチャート（月・週の目盛り、マイルストーン）<br>スケジュール（ガント）。 rows: 行ラベル list。 bars: [(row, 開始date, 終了date, ラベル, style="primary"|"mid"|"pale", lane=0)] milestones: [(date, ラベル)]  週数が26以下なら週の目盛りも表示 | `(slide, area, start, end, rows, bars, milestones=(), label_w=150, size=12)` | examples/catalog/python/slide04_gantt.py |
| `calendar_month` | 月間カレンダー（日付ごとの予定）<br>events: {日: "予定"}。土日は薄いグレー | `(slide, area, year, month, events=None, size=10)` | examples/parts/python/slide61_calendar.py |
| `schedule_list` | 日程リスト（日付＋予定＋状態タグ）<br>items: [(日付ラベル, 予定, 状態 or None)]  状態は status_table と同じ（完了/進行中/未着手/遅延/保留） | `(slide, area, items, size=13)` | examples/parts/python/slide62_schedule_b.py |
| `chronology` | 年表（期間×分野のマスに出来事を記入）<br>events: {(行index, 期間index): テキスト or list} | `(slide, area, periods, rows, events, label_w=110, size=11)` | examples/parts/python/slide9e_chrono.py |

## 画像（5）

| 関数 | 用途 | 引数 | 実例 |
|---|---|---|---|
| `image` | 画像を縦横比を保って配置<br>画像を領域内に縦横比を保って中央配置 | `(slide, area, path)` | examples/parts/python/slide15_misc.py |
| `image_caption` | 画像＋キャプション | `(slide, area, path, caption, size=12)` | examples/parts/python/slide60_media.py |
| `image_grid` | 画像を格子状に並べる（キャプション付き）<br>images: [(パス, キャプション)] | `(slide, area, images, cols=3, size=11)` | examples/parts/python/slide60_media.py |
| `screenshot_callouts` | 画面キャプチャに番号付きの注釈<br>callouts: [(x 0-1, y 0-1, 説明)] 画像上の相対位置に番号を置き、右に説明を並べる | `(slide, area, path, callouts, ratio=0.64, size=12)` | examples/parts/python/slide60_media.py |
| `image_placeholder` | 画像の差し込み予定枠（後で手で画像を貼る場所） | `(slide, area, label='画像を挿入')` | examples/parts/python/slide15_misc.py<br>examples/parts/python/slide60_media.py |

## アイコン（3）

| 関数 | 用途 | 引数 | 実例 |
|---|---|---|---|
| `icon` | アイコン（図形で描画・編集可能）。kind は ICONS の18種<br>kind: person / people / document / pc / server / database / cloud / building / factory / money / phone / mail / gear / lock / chart / check / warning / clock | `(slide, x, y, s, kind, color='001F5F')` | examples/parts/python/slide84_zone.py |
| `icon_label` | アイコン＋ラベル（＋補足）を縦に並べたノード。外枠の領域を返す | `(slide, area, kind, label, sub=None, color='001F5F', size=12)` | examples/parts/python/slide80_icons.py<br>examples/parts/python/slide84_zone.py |
| `icon_steps` | アイコン付きのステップ（アイコン → 矢印 → アイコン）<br>steps: [(icon_kind, 見出し, 説明)] | `(slide, area, steps, size=12)` | examples/parts/python/slide80_icons.py |

## 線・矢印（8）

| 関数 | 用途 | 引数 | 実例 |
|---|---|---|---|
| `connect` | 2つの箱を縁から縁へ自動でつなぐ（種類: solid/dashed/dotted/thick/both/service/money/goods/info）<br>a, b は (x,y,w,h) か図形。offset で平行線をずらす（往復の矢印を2本引くとき） | `(slide, a, b, kind='solid', label=None, offset=0)` | examples/parts/python/slide81_lines.py<br>examples/parts/python/slide84_zone.py |
| `elbow_connect` | カギ線で2つの箱をつなぐ（first='h' 横→縦→横 / 'v' 縦→横→縦） | `(slide, a, b, first='h', kind='solid', label=None)` | examples/parts/python/slide81_lines.py |
| `curved_arrow` | 曲線の矢印（2点間） | `(slide, p1, p2, kind='solid', label=None)` | examples/parts/python/slide81_lines.py |
| `loop_arrow` | 循環の矢印（ぐるっと回る矢印＋中央ラベル） | `(slide, area, label=None, size=13)` | examples/parts/python/slide81_lines.py |
| `block_arrow` | 文字入りの太い矢印（right/left/up/down/both） | `(slide, area, body, direction='right', style='pale', size=14)` | examples/parts/python/slide81_lines.py |
| `flow_legend` | 線の種類の凡例（connect の kind を並べる）。幅を返す | `(slide, x, y, kinds, size=10)` | examples/parts/python/slide81_lines.py |
| `annotate_points` | 図の上に番号を打ち、欄外に説明を並べる（任意の図に使える番号注釈）<br>points: [(x, y)] 絶対座標(pt)。notes: 説明（points と同数）。area: 説明を並べる領域 | `(slide, points, notes, area, size=12, style='alert')` | examples/parts/python/slide82_annotate.py |
| `span_bracket` | 範囲を示すかっこ線（期間・対象範囲にラベル） | `(slide, x1, x2, y, label, below=True, size=11)` | examples/parts/python/slide81_lines.py |

## 補助（7）

| 関数 | 用途 | 引数 | 実例 |
|---|---|---|---|
| `source` | 左下の出所・注記<br>左下の出所・注記（9pt） | `(slide, text)` | examples/catalog/python/slide03_split.py<br>examples/catalog/python/slide07_chart.py |
| `arrow` | グレーの三角矢印<br>グレーの三角矢印。direction: right / down | `(slide, area, direction='right')` | examples/parts/python/slide15_misc.py |
| `footnotes` | 脚注リスト（注1, 注2…） | `(slide, area, notes, size=10)` | examples/parts/python/slide14_cards_c.py |
| `highlight_frame` | 赤の点線で囲む注目枠（ラベル付き可） | `(slide, area, label=None)` | examples/parts/python/slide15_misc.py |
| `number_badge` | 丸数字バッジ | `(slide, x, y, n, d=24, style='primary')` | examples/parts/python/slide10_text_a.py |
| `ref_tag` | 右上の「参考」などの区分タグ | `(slide, label='参考')` | examples/parts/python/slide11_text_b.py |
| `legend` | 凡例（色見本＋ラベル）。幅を返す<br>items: [(ラベル, COLOR_*), ...] | `(slide, x, y, items, size=11)` | examples/parts/python/slide10_text_a.py |

## 会議・報告（7）

| 関数 | 用途 | 引数 | 実例 |
|---|---|---|---|
| `traffic_light` | 信号機式の状況報告（項目ごとに 良好/注意/問題 の信号とコメント）<br>rows: [(項目, 状態 "良好"|"注意"|"問題", コメント)] | `(slide, area, rows, size=12)` | examples/parts/python/slide90_meeting.py |
| `decision_log` | 決定事項・宿題・確認事項の一覧（種別タグ＋内容＋担当・期限）<br>items: [(種別 "決定"|"宿題"|"確認"|"保留", 内容, 担当 or "", 期限 or "")] | `(slide, area, items, size=12)` | examples/parts/python/slide90_meeting.py |
| `meeting_bodies` | 会議体一覧（会議名・頻度・参加者・目的） | `(slide, area, rows, size=12, stretch=False)` | examples/parts/python/slide92_meeting_c.py |
| `weekly_report` | 週次報告（全体状況の信号＋今週の実績／来週の予定／課題・リスク）<br>overall = (状態 "良好"|"注意"|"問題", コメント) | `(slide, area, overall, done, plan, issues, period=None, size=12)` | examples/parts/python/slide91_meeting_b.py |
| `meeting_agenda` | 会議のアジェンダ（時間・議題・担当）<br>items: [(時間, 議題, 担当)] | `(slide, area, items, size=13)` | examples/parts/python/slide90_meeting.py |
| `revision_history` | 変更履歴（版・日付・変更内容・担当） | `(slide, area, rows, size=12, stretch=False)` | examples/parts/python/slide92_meeting_c.py |
| `notice_box` | お知らせ・注意の枠（info / warn） | `(slide, area, title, items, kind='info', size=13)` | examples/parts/python/slide91_meeting_b.py |

## リスク・課題（3）

| 関数 | 用途 | 引数 | 実例 |
|---|---|---|---|
| `risk_matrix` | リスクマトリクス（影響度×発生可能性の3×3に番号を配置＋右に一覧）<br>risks: [(リスク名, 影響度 1-3, 発生可能性 1-3)] | `(slide, area, risks, size=12)` | examples/parts/python/slide93_risk.py |
| `risk_table` | リスク対応表（影響度・発生可能性を高中低で色分け）<br>rows: [(リスク, 影響度 高中低, 発生可能性 高中低, 対策, 担当)] | `(slide, area, rows, size=12, stretch=False)` | examples/parts/python/slide94_risk_b.py |
| `issue_tracker` | 課題管理表（優先度を色分け、状況をステータス色で）<br>rows: [(課題, 優先度 高中低, 担当, 期限, 状況)] | `(slide, area, rows, size=12, stretch=False)` | examples/parts/python/slide94_risk_b.py |

## 費用・効果（4）

| 関数 | 用途 | 引数 | 実例 |
|---|---|---|---|
| `cost_table` | 費用内訳表（小計・合計行、金額は右寄せ・カンマ区切り）<br>rows: [(費目, 金額, ...)]。費目が「小計」で始まる行は小計として強調。auto_total で最終行に合計を追加 | `(slide, area, header, rows, unit=None, auto_total=True, size=12)` | examples/parts/python/slide95_cost.py |
| `roi_chart` | 投資回収グラフ（累計投資と累計効果の折れ線＋回収時期）<br>investment/effect: 期ごとの金額。cumulative=True なら累計に変換して描く | `(slide, area, periods, investment, effect, unit=None, cumulative=True, size=12)` | examples/parts/python/slide95_cost.py |
| `budget_actual_table` | 予算実績対比表（計画・実績・差異・達成率。未達は赤）<br>rows: [(項目, 計画, 実績)] | `(slide, area, rows, unit=None, size=12)` | examples/parts/python/slide96_cost_b.py |
| `effect_breakdown` | 効果の内訳（定量効果の金額一覧＋合計／定性効果の箇条書き）<br>quantitative: [(項目, 金額)] | `(slide, area, quantitative, qualitative, unit='万円/年', size=12)` | examples/parts/python/slide96_cost_b.py |

## 要件・仕様（2）

| 関数 | 用途 | 引数 | 実例 |
|---|---|---|---|
| `requirements_table` | 要件一覧（ID・区分・要件・優先度 MoSCoW を色分け）<br>rows: [(ID, 区分, 要件, 優先度 "Must"|"Should"|"Could"|"Won't")] | `(slide, area, rows, size=11, stretch=False)` | examples/parts/python/slide97_req.py |
| `grouped_table` | グループ化した表（1列目の同じ値を縦に結合。機能一覧・非機能要件など） | `(slide, area, header, rows, col_ratios=None, size=12, stretch=False)` | examples/parts/python/slide97_req.py |

## 目標管理（3）

| 関数 | 用途 | 引数 | 実例 |
|---|---|---|---|
| `kpi_tree` | KGI/KPI ツリー（各ノードに 現状 → 目標）<br>root = (ラベル, 現状, 目標, [子, ...]) | `(slide, area, root, size=11)` | examples/parts/python/slide98_goal.py |
| `okr` | OKR（Objective と Key Results の進捗）<br>krs: [(KR の文, 進捗 0-100)] | `(slide, area, objective, krs, size=13)` | examples/parts/python/slide99_goal_b.py |
| `target_actual` | 目標と実績の対比（目標＝枠、実績＝塗りの横棒と達成率）<br>items: [(ラベル, 目標, 実績)] | `(slide, area, items, unit='', size=12)` | examples/parts/python/slide99_goal_b.py |

## 顧客・利用者（3）

| 関数 | 用途 | 引数 | 実例 |
|---|---|---|---|
| `customer_journey` | カスタマージャーニー（段階×行動・接点・課題＋感情の曲線）<br>rows: [(行ラベル, [各段階のテキスト])]  emotions: 各段階の感情 -2〜+2 | `(slide, area, stages, rows, emotions=None, label_w=90, size=11)` | examples/parts/python/slide9a_customer.py |
| `likert_chart` | アンケート結果（5段階評価の100%帯グラフ。ネイティブ）<br>data[設問] = [各段階の割合(%)] | `(slide, area, questions, data, labels=('そう思う', 'ややそう思う', 'どちらでもない', 'あまりそう思わない', 'そう思わない'), size=11)` | examples/parts/python/slide9b_customer_b.py |
| `empathy_map` | 共感マップ（言う・考える・行動・感じる＋中央に人物） | `(slide, area, persona, says, thinks, does, feels, size=12)` | examples/parts/python/slide9b_customer_b.py |

## 意思決定（3）

| 関数 | 用途 | 引数 | 実例 |
|---|---|---|---|
| `recommendation` | 推奨案（選択肢カードで推奨を明示＋推奨理由・次のアクション）<br>options: [(案の名前, 概要)] | `(slide, area, options, recommended, reasons, actions, size=12)` | examples/parts/python/slide9c_decision.py |
| `assumptions` | 前提条件と制約事項（番号付き） | `(slide, area, premises, constraints, size=13)` | examples/parts/python/slide9d_decision_b.py |
| `plan_cards` | 料金・プラン比較カード（おすすめを強調）<br>plans: [(プラン名, 価格表示, [特徴, ...])] | `(slide, area, plans, recommended=None, size=12)` | examples/parts/python/slide9d_decision_b.py |

## フレームワーク（3）

| 関数 | 用途 | 引数 | 実例 |
|---|---|---|---|
| `framework_grid` | 定番フレームワーク（kind: 3c / 4p / 4c / pest / kpt）<br>contents: 各ブロックの箇条書き list を順に（3c なら [顧客, 競合, 自社]） | `(slide, area, kind, contents, size=12)` | examples/parts/python/slide9f_framework.py |
| `five_forces` | 5フォース分析（中央の業界と4つの力を矢印で）<br>forces: {"新規参入の脅威": items, "代替品の脅威": items, "売り手の交渉力": items, "買い手の交渉力": items} | `(slide, area, center, forces, size=12)` | examples/parts/python/slide9g_framework_b.py |
| `business_model_canvas` | ビジネスモデルキャンバス（9ブロック）<br>blocks: {"KP": [], "KA": [], "KR": [], "VP": [], "CR": [], "CH": [], "CS": [], "C$": [], "R$": []} | `(slide, area, blocks, size=10)` | examples/parts/python/slide9g_framework_b.py |
