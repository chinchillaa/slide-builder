# slide-builder 開発の引継ぎメモ

最終更新: 2026-10-21。このファイルは **スキルを開発する人（AI）向け** で、スキルを使う人向けの説明は `../SKILL.md` にある。
新しいセッションでは、まずこのファイルを読み、続けて `TODO.md` と `../SKILL.md` を読む。

## 1. これは何か
- Insight Edge の記事「AIにデザインさせるのをやめた」（https://techblog.insightedge.jp/entry/powerpoint-auto-generation）の考え方で作った、業務用 pptx を生成する Claude Code スキル
- 考え方: デザインの判断は AI にさせず、ヘルパー関数に固定する。規則はコードで強制する（溢れ・下限フォントは ValueError）。1枚1スクリプトで書く。構成案は人間が承認する。完成後は別のサブエージェントが画像で目視レビューする。描画は PowerPoint COM で行う
- デザインの手本: `C:\Users\2gawa\temp\slides` にある官公庁向けコンサル報告書の PDF 10本。パンくず 14pt、タイトル 24pt、紺の罫線、リード 16pt、Meiryo UI、紺 #001F5F
- ユーザーの方針: 普段の業務用。グラフは PowerPoint 標準のネイティブグラフ。テンプレートは不要。型（部品）はできるだけ多様にする
- GitHub: https://github.com/chinchillaa/slide-builder（**Private**、ブランチ main）。リポジトリのルートはスキルのフォルダそのもの

## 2. ユーザーとのやりとりの決まり
- **返答は日本語**で書く
- **危険な操作は避ける**。過去に `rm -f` を拒否された。ファイルは削除せず、新しいフォルダに書き出す
- 外部への push は、そのつど確認してから行う（初回のリポジトリ作成と push は承認済み）
- `--force` でビルドすると手で直した部分が消える。付ける前にユーザーに確認する

## 3. 環境
- Windows 11、Python 3.13。python-pptx 1.0.2、pywin32、PyMuPDF はユーザーのサイトパッケージに入っているので、`python -I` では import できない
- 出力の文字化けを防ぐため `PYTHONIOENCODING=utf-8` を付ける（cp932 でエラーになる）
- PowerPoint はインストール済み。render.py は COM 経由で、ウィンドウを出さず読み取り専用で開く。閉じるのは自分で開いたファイルだけで、ほかに何も開いていないときだけ終了する
- git の user.name と user.email は、このリポジトリだけに設定してある（chinchillaa / noreply アドレス）

## 4. 構成（コアの場所）
| パス | 役割 |
|---|---|
| `scripts/pptx_helpers.py` | コア。定数（色・寸法・フォント）、`REGISTRY` と `@part`、`_put_text`（溢れ・単語途中の改行・下限の検査）、スライド枠、基本部品、`check_presentation`（lint）。末尾で parts_*.py を import している |
| `scripts/parts_text/table/flow/structure/data/media.py` | 初期の 112 部品 |
| `scripts/parts_connect.py` | アイコン18種と線・矢印（`connect` は箱の縁から縁へ自動で配線） |
| `scripts/parts_diagram.py` | 参考資料に多い図解（zone, business_model, spec_text など） |
| `scripts/parts_biz.py` | 業務でよく使う型（会議・リスク・費用・目標・顧客・意思決定・フレームワーク） |
| `scripts/build_deck.py` | `python build_deck.py <deck_dir> [--force]`。失敗したスライドはエラースライドに置き換える。lint を実行する。手で編集されていたら上書きしない |
| `scripts/render.py` | `python render.py <pptx> [out_dir] [--slides 3,5]`。**`--slides` は 1 始まり**（slide_NN.png と同じ番号） |
| `scripts/list_parts.py` | `reference/parts.md` を生成する（`--count` で件数だけ出す） |
| `reference/patterns.md` | 用途ごとの部品の組み合わせ（手で書いている） |
| `reviewers/plan_review.md`, `visual_review.md` | レビュー用サブエージェントへの指示文 |
| `examples/catalog/` | 9枚の基本デッキ |
| `examples/parts/` | 全部品のカタログデッキ（64枚、全167部品を少なくとも1回使う） |
| `examples/sample-ai-helpdesk/` | スキルの手順どおりに作った実案件風のサンプル（16枚、数値は架空） |

部品の数: **167**（23カテゴリ）。カテゴリの並び順は `list_parts.py` の ORDER で決まる。

## 5. 守っている設計の決まり
- 単位は pt で、スライドは 960×540。`BODY=(34,118,892,378)`、`GAP=12`
- 文字サイズはパンくず 14、タイトル 24、リード 16、本文 14。下限は本文 12（`MIN_BODY_PT`）、表・図など密な部品は 10（`MIN_DENSE_PT`）。パンくずなどの枠の部品は `exempt=True` で検査を外す
- 色はパレット定数だけを使う。`COLOR_ACCENT`（赤）は文字だけ。`COLOR_STATUS_OK/WARN` は信号の部品だけ
- 文字幅の見積もり: 全角 1.0、そのほか 0.6、行の高さ 1.25。楕円は 0.7 倍、ひし形は 0.5 倍で計算する。`check_tokens` は英単語や数字が途中で改行されるとエラーにする。`fit_one_line` は大きな数字を1行に収める
- マークアップ: `**太字**`、`!!赤太字!!`、`__下線__`、`\n`、`(level, text)`
- コネクタは `_no_effect` でテーマの影を消す。アイコンは図形をグループにしたもの
- lint で重なりを調べるとき、名前が `deco` / `arrow` / `chrome:` で始まる図形は対象外
- `source()` は文頭が「出所/出典/注/※」でなければ「出所：」を自動で付ける

## 6. 部品を追加・変更するときの手順
1. 該当する `scripts/parts_*.py` に `@part("カテゴリ", "説明")` を付けた関数を書く。文字は `_put_text` / `_textbox` を通す
2. `examples/parts/python/` に実例を書く（新しいファイルか、既存のファイルに追記）
3. 次を実行する
   ```
   cd <scratchpad>   # スキルのフォルダ以外で実行する
   PYTHONIOENCODING=utf-8 python ~/.claude/skills/slide-builder/scripts/build_deck.py ~/.claude/skills/slide-builder/examples/parts
   PYTHONIOENCODING=utf-8 python ~/.claude/skills/slide-builder/scripts/build_deck.py ~/.claude/skills/slide-builder/examples/catalog
   PYTHONIOENCODING=utf-8 python ~/.claude/skills/slide-builder/scripts/build_deck.py ~/.claude/skills/slide-builder/examples/sample-ai-helpdesk
   PYTHONIOENCODING=utf-8 python ~/.claude/skills/slide-builder/scripts/render.py <pptx> --slides N
   ```
   - 3つのデッキがすべて **ERROR 0 件** になることを確かめる
   - WARN は、sample-ai-helpdesk の slide13 に出る ▲ マイルストーンの重なりが既知の誤検知なので、それ以外が増えていないことを確かめる
4. 画像を Read で開き、目で確認する
5. `python scripts/list_parts.py` を実行して parts.md を更新する。必要なら SKILL.md の部品表（件数）と patterns.md も直す
6. commit する（push はユーザーに確認してから）

## 7. ハマりどころ
- bash の heredoc で例のスクリプトを書くと、`Won't` のクォートで崩れたり、`\\n` が実際の改行になったりする。**Write ツールか Python スクリプトで書く**。Bash ツールの中で Python の文字列に `\\n` を書いた場合も、実際の改行になって SyntaxError を起こしたことがある
- `pptx.enum.connector` は存在しない。`pptx.enum.shapes` を使う
- グラフは `has_title=False` にしないとタイトルが自動で付く。100% グラフの軸は書式を 0% にする
- 「◎ 文字」の形のときだけ、評価記号として色を付ける（「×2」などを誤って色付けしないため）
- ガントのバーのラベルは、バーの幅に収まらないとエラーになる。ラベルを短くするか、期間を延ばす

## 8. これまでの経緯
- 2026-10-20: 112部品に拡張（ユーザーの要望は「100種類以上」）→ さらに A（参考資料の型）・B（業務の型）・C（図の基盤）を追加して167部品にした。D（1回の呼び出しでスライド1枚を作るテンプレート）は提案しただけで、ユーザーは未依頼
- 2026-10-21: サンプル資料 `sample-ai-helpdesk` をスキルの手順どおりに作成した。構成レビューで Must 6件を反映し、承認を得てから実装した。目視レビューの Must も反映済み。その後 GitHub に Private で公開した
- 次にやる候補は `TODO.md` を参照
