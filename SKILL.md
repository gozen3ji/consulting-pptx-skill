---
name: consulting-pptx-skill
description: コンサル型スライド38型の「型システム」から経営会議品質のPowerPoint資料を作成するスキル。型カタログから型を選び、SlideSpec（JSON）を書き、生成パイプラインでHTMLプレビュー→編集可能PPTXを自動生成し、機械チェックで規約を担保する。トリガー例:「スーパーテンプレで」「38型から選んで」「コンサル型のPPTX作って」。
---

# コンサル型スライド作成スキル（38型・SlideSpec生成パイプライン）

戦略系コンサルティングファームの公開資料を大量に実測して抽出したスライドの「型」38種類を、
**SlideSpec（JSONによるスライド定義）→ HTMLプレビュー → 編集可能PPTX** の自動生成パイプラインで組み立てる。

テンプレPPTXを手でコピーするのではなく、**型のデータ定義（SlideSpec）を書いて機械に組ませる**のが本来の使い方。
スライドの設計規約は `references/slide-rules.md` が正典。**作成前に必ず読み、出力後は `scripts/check_deck.py` で機械チェックする。**

## 同梱物

| もの | パス |
| --- | --- |
| 生成パイプライン（Node製） | `pipeline/`（validate / render / qa / export） |
| 38型のSlideSpec正本 | `pipeline/slide-spec/super_template.json` |
| SlideSpecスキーマ | `pipeline/slide-spec/schema.json` |
| スライド設計規約（正典） | `references/slide-rules.md` |
| 機械チェックスクリプト | `scripts/check_deck.py`（要 `pip3 install python-pptx`） |
| おまけ: テンプレPPTX見本帳 | `assets/SuperTemplate_38type.pptx`（全38型を1枚ずつ収録。手動利用・一覧確認用） |

## 38型カタログ（型ID / 使いどころ）

各スライド共通フィールド: `kicker`（左上の小見出し）/ `title`（**12字以上・主張を書く**）/ `source`（出典行・必須）。
フィールドの実例値（数値の形・series構造など）は `super_template.json` の該当スライドを**その型だけ**読んで確認する。

| # | 型ID | 使いどころ |
| --- | --- | --- |
| 0 | `cover` | 表紙 |
| 1 | `executive_summary` | 冒頭で結論と論点を一望させる場合 |
| 2 | `evidence_basis` | この資料が何に基づくかを冒頭で示す場合 |
| 3 | `big_stat_pair` | 大型数値2つで規模やインパクトを対比する場合 |
| 4 | `kpi_dashboard` | 主要KPIを一覧で示す場合 |
| 5 | `chart_insight` | 1つのチャートで主張を証明し、含意を添える場合 |
| 6 | `stacked_bar` | 構成の変化を積み上げ棒で示す場合 |
| 7 | `waterfall` | 増減の寄与をブリッジで示す場合 |
| 8 | `true_waterfall` | 起点から着地までの増減を厳密なブリッジで示す場合 |
| 9 | `small_multiples` | 同じ図法を並べて切り口違いで比較する場合 |
| 10 | `comparison_table` | 複数の選択肢を評価軸で比較する場合 |
| 11 | `scenario_table` | シナリオ別の前提と結果を並べる場合 |
| 12 | `risk_table` | リスク・兆候・打ち手を整理する場合 |
| 13 | `horizontal_axis_table` | 横軸に項目を並べて評価する場合 |
| 14 | `heatmap_table` | 濃淡で強弱を一覧表示する場合 |
| 15 | `matrix_2x2` | 2つの軸で位置づけを整理する場合 |
| 16 | `process_matrix` | プロセスと観点の掛け合わせで整理する場合 |
| 17 | `nested_row_matrix` | 入れ子の行構造で階層を示す場合 |
| 18 | `timeline_matrix` | 時系列と項目のマトリクスで示す場合 |
| 19 | `theme_card_grid` | 複数のテーマをカードで並べる場合 |
| 20 | `recommendation_pillars` | 複数の提言を柱立てで示す場合 |
| 21 | `numbered_imperatives` | やるべきことを番号付きで示す場合 |
| 22 | `question_framework` | 検討すべき問いを構造化する場合 |
| 23 | `scr` | 状況・複雑化・解決の3段で語る場合 |
| 24 | `issue_to_solution_map` | 課題と解決策を対応付ける場合 |
| 25 | `issue_cause_solution` | 課題から原因、解決策へ流れで示す場合 |
| 26 | `issue_tree` | 課題をツリーで分解する場合 |
| 27 | `cause_effect` | 原因と結果の連鎖を示す場合 |
| 28 | `current_target_state` | 現状と目指す姿を対比する場合 |
| 29 | `calc_flow` | 計算ロジックを式の流れで示す場合 |
| 30 | `process_flow` | プロセスの流れを段階で示す場合 |
| 31 | `cycle` | 循環するサイクル構造を示す場合 |
| 32 | `chevron_rail` | 段階の進行を矢羽（シェブロン）で示す場合 |
| 33 | `chevron_value_chain` | バリューチェーン全体を示す場合 |
| 34 | `decision_fork` | 分岐する選択肢と判断を示す場合 |
| 35 | `roadmap` | 実行ロードマップを示す場合 |
| 36 | `gantt` | スケジュールをガントチャートで示す場合 |
| 37 | `decision_page` | 意思決定を求める場合 — 決めること・前提・依頼 |

## デッキ作成の手順（実運用フロー）

1. **作る前に定義する（Define-before-Produce）**: 目的・成果物の定義・スコープIN/OUTを3〜5行で先に合意する。前提が薄いまま豪華な体裁で出すのが最悪の失敗。
2. ストーリーに合わせてカタログから型を選ぶ。**枠組み → 証拠 → 比較 → 構造 → 計画** の順が基本形。
3. `pipeline/slide-spec/super_template.json` から該当する型のスライド定義をコピーして新しいSlideSpec（JSON）を作り、文言・データを実物に差し替える。
   - `title` は12字以上・主張文（体言止め）。`source` 行は必須。プレースホルダー（Text N / ラベルN）を1つも残さない。
4. **パイプラインでビルド**（初回のみ `cd pipeline && npm run setup`）:
   ```bash
   cd pipeline
   node scripts/validate_spec.mjs slide-spec/mydeck.json          # スキーマ検証
   node scripts/render_spec_to_html.mjs slide-spec/mydeck.json generated/mydeck.html   # HTMLプレビュー
   node scripts/qa_html_deck.mjs generated/mydeck.html            # 構造QA
   node scripts/export_spec_to_editable_pptx.mjs slide-spec/mydeck.json generated/mydeck.pptx  # 編集可能PPTX
   ```
5. **機械チェック**: `python3 scripts/check_deck.py generated/mydeck.pptx` を実行し、FAIL 0 にする。出力されるタイトル一覧を上から通し読みして、1本のストーリーになっているか確認する。
6. **目視QA**: PDF化して全ページを確認する。文字だけでなく余白・版面バランス・孤立折返し・はみ出し・左右カラムの下端揃いも見る。

書き出したPPTXは全図形がネイティブ編集可能（`editable: true`）。納品後の微修正はPowerPoint上でそのままできる。

## おまけ: テンプレPPTX見本帳の使い方

`assets/SuperTemplate_38type.pptx` は同じ38型を1型=1スライドで収めた見本帳。
- 型選定時の**目視カタログ**として眺める
- Node環境が無い場合の**手動フォールバック**として、該当スライドをコピーして文言を差し替える
のどちらでも使える。ただし品質の再現性はSlideSpecパイプライン経由のほうが高い。

## カラーカスタマイズ

デフォルトはニュートラルなネイビー系。SlideSpecのルートに `palette` オブジェクトを入れると全スライドのブランドカラーを一括で差し替えられる（`schema.json` 参照）。その際も以下は守る:

- 意味を持つ色（✕の赤・追加の緑など）は変換しない
- 2系列の区別が要る図では、メインカラー×グレー系の2色に抑える
- 製品UIのスクリーンショットは無加工
