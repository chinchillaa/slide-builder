# スライドの型（部品の組み合わせレシピ）

構成計画の「部品」欄は、ここから型を選んで書く。どの型も `content_slide(...)` の BODY 領域に置く前提。
領域は `split_cols` / `split_rows` / `split_columns` / `grid_areas` で BODY を分けて作る。座標を勘で書かない。

## 導入・全体像
| 型 | 使う場面 | 組み合わせ |
|---|---|---|
| 表紙 | 冒頭 | `cover_slide` |
| 目次 / アジェンダ | 全体の見取り図。章の冒頭で現在地を示す | `toc_slide` / `agenda_slide(current=章番号)` |
| 章の区切り | 章の切り替え | `section_slide(title, items)` |
| エグゼクティブサマリー | 結論を先に示す | `exec_points`（3〜5点）＋下部に `conclusion_bar` |
| 背景 → 目的 | 検討の出発点 | `split_columns(arrow=True, titles=["背景","目的"])` ＋左 `bullets(box=True)` ＋右 `message_box`×2〜3 |
| 検討の観点 | 論点を並べる | `band_heading` ＋ `card_grid(cols=3)` |
| 実施概要 | 何をやったか | `key_value_table`（目的・期間・対象・方法） |
| 実施内容と該当頁 | 報告書の目次代わり | `card_table(styles=["label","text","gray"])` |

## 現状・分析
| 型 | 使う場面 | 組み合わせ |
|---|---|---|
| 現状 → 課題 | 問題の提示 | `split_columns(arrow=True)` ＋ `bullets` / `card_table` |
| 課題と対応方針 | 分類ごとの課題 | `card_table(styles=["label","label2","text","text"], arrow_before=(3,))` |
| 原因の分解 | なぜを掘る | `logic_tree` |
| 業務フロー（As-Is） | 現行業務の可視化 | `swimlane`（担当×工程） |
| 工程ごとの課題 | フローの各段の問題 | `step_matrix`（行=現状/課題） |
| ヒアリング結果 | 現場の声 | `quote_box`×2〜3 ＋ `summary_box` |
| 定量分析 | 数字の推移・内訳 | `chart_with_insight` または `split_cols` ＋ `chart`×2 |
| 増減の要因 | 前年との差の内訳 | `waterfall` ＋右に `bullets` |
| ランキング | 部署別・項目別 | `rank_bars(highlight=0)` または `highlight_table` |
| 構成比 | 費用・時間の内訳 | `percent_bar` / `chart("doughnut")` |
| 分布・相関 | 2変数の関係 | `scatter_chart` / `bubble_chart` / `positioning_map` |
| 地域×期間 | 大小を色で見る | `heatmap_table` |
| 外部環境 | 強み弱み | `swot` |
| 関係者 | 利害関係・情報の流れ | `relation_map` / `hub_spoke` |
| 利用者像 | 誰の課題か | `persona_card` |

## 方針・比較・意思決定
| 型 | 使う場面 | 組み合わせ |
|---|---|---|
| あるべき姿 | To-Be の提示 | `gap_analysis` |
| Before / After | 変化を見せる | `before_after` ＋ `stat_compare` |
| 選択肢の比較 | 案の評価 | `comparison_table(recommend=推奨列)` ＋ `conclusion_bar` |
| 定量評価 | 重み付きの採点 | `score_table(weights=...)` |
| 機能の充足度 | ベンダー比較 | `harvey_table` / `check_table` |
| 2案の対比 | A か B か | `versus` / `pros_cons` |
| 優先順位付け | 施策の選別 | `matrix_2x2(highlight=右上)` |
| 方針の宣言 | 1メッセージを強く | `statement` |
| 効果の試算 | 投資対効果 | `big_number` ＋ `kpi_row` |
| 構成要素の統合 | A＋B で C | `equation` / `converge` / `brace_group` |
| 施策の展開 | 1方針から複数施策へ | `diverge` / `logic_tree` |

## 計画・体制・実行
| 型 | 使う場面 | 組み合わせ |
|---|---|---|
| 進め方 | フェーズと内容 | `step_flow(active=...)` ＋下に `kpi_row` |
| ロードマップ | 年度×領域 | `roadmap` |
| 詳細スケジュール | 週単位の計画 | `gantt(milestones=...)` |
| 直近の日程 | 会議・締切 | `schedule_list` / `calendar_month` / `timeline` |
| 判定ポイント | 稟議・ゲート | `phase_gate` |
| 体制図 | 誰が何をするか | `org_chart` ＋ `role_cards` |
| 役割分担 | 責任の所在 | `raci_table` |
| システム構成 | アーキテクチャ | `layer_stack` |
| 進捗報告 | 状況の共有 | `progress_bars(target=...)` ＋ `status_table` |
| 宿題・ToDo | 会議の締め | `action_items` |
| 成熟度・段階 | 目指すレベル | `stair_steps` / `pyramid` |
| 改善サイクル | 継続運用 | `cycle` |
| 締め | 最後 | `closing_slide` |

## 図解・システム説明
| 型 | 使う場面 | 組み合わせ |
|---|---|---|
| システム概念図 | 範囲と構成要素 | `zone` で領域を分け、中に `icon_label` と `message_box`、間を `connect` |
| 事業スキーム・ビジネスモデル | お金・サービス・モノの流れ | `business_model`（凡例は自動） |
| データの流れ | 収集 → 蓄積 → 活用 | `pipeline` / `icon_steps` |
| 業務の変化 | オンライン化の前後 | `flow_before_after(changed=...)` |
| 章内の現在地 | 長い章の中の位置 | `section_tab` で本文領域を作る / `issue_map(current=...)` |
| 仕様・要件の本文 | 文章が主体の説明 | `spec_text`（①見出し・i./イ) の階層・__下線__） |
| データ仕様の例 | JSON・設定例 | `code_block` ＋右に `definition_list` |
| 画面・図の説明 | 番号で指し示す | `screenshot_callouts` / 任意の図に `annotate_points` |
| 対象の絞り込み | 母数から対象へ | `reduction_bars` / `count_tree` |
| 事例紹介 | 先進事例・海外事例 | `case_card`×2（上下） |
| プレイヤー比較 | 企業×事業領域 | `mark_matrix(groups=..., legend=...)` / `bar_table` |

## 会議・管理
| 型 | 使う場面 | 組み合わせ |
|---|---|---|
| 状況報告 | 定例・ステコミ | `traffic_light` ＋ `issue_tracker` |
| 週次報告 | 毎週の報告 | `weekly_report` |
| 会議の冒頭／締め | 議題と決定事項 | `meeting_agenda` / `decision_log` |
| 会議体の設計 | 体制の説明 | `meeting_bodies` ＋ `org_chart` |
| リスク管理 | リスクの評価と対策 | `risk_matrix` → `risk_table` |
| 費用の説明 | 予算・見積り | `cost_table(auto_total=True)` ＋ `roi_chart` |
| 予実管理 | 計画と実績の差 | `budget_actual_table` / `target_actual` |
| 効果の説明 | 定量・定性 | `effect_breakdown` ＋ `big_number` |
| 要件の整理 | 要件定義 | `requirements_table` / `grouped_table`（機能一覧・非機能要件） |
| 目標の分解 | KGI → KPI | `kpi_tree` / `okr` |
| 顧客理解 | 体験・声・人物像 | `customer_journey` / `likert_chart` / `empathy_map` / `persona_card` |
| 意思決定を求める | 稟議・提案 | `recommendation` ＋ `assumptions` |
| 料金・プラン | 提案・比較 | `plan_cards(recommended=...)` |
| 経緯・沿革 | 背景の説明 | `chronology` |
| 環境分析 | 戦略検討 | `framework_grid("3c"/"pest"/"4p"/"4c")` / `five_forces` / `swot` |
| 事業の全体像 | 新規事業 | `business_model_canvas` |
| 振り返り | プロジェクト終了時 | `framework_grid("kpt")` |
| 表紙・章の別デザイン | 雰囲気を変えたい | `cover_band` / `section_split` |

## 組み合わせの原則
- 1スライドに主役の部品は1つ。補助（`message_box`・`conclusion_bar`・`source`・`note_box`）は1〜2個まで
- 結論は「下に `conclusion_bar`」か「右に `titled_box("示唆")`」のどちらかに統一する
- 同じデッキで同じ意味の型は同じ部品で表す（課題一覧はすべて `card_table` にする、など）
- 数字の出典は必ず `source` で書く
