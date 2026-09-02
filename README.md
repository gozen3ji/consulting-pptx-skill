# consulting-pptx-skill

**AIに「経営会議レベル」のPowerPointを作らせるためのClaude Codeスキル。**
コンサル型スライド38型のカタログ＋SlideSpec（JSONによるスライド定義）＋生成パイプライン（HTMLプレビュー→編集可能PPTX）＋スライド作成規約＋機械チェックの一式です。

A Claude Code skill for generating boardroom-quality PowerPoint decks: a catalog of 38 consulting slide archetypes, a JSON SlideSpec format, a render pipeline (HTML preview → natively editable PPTX), a slide-design rulebook, and an automated rule checker.

私たちが実際に毎週の提案書・報告書づくりで回している資料作成システムの公開版です。解説記事はこちら → [note（【公式】Jinba）](https://note.com/jinbaflow)

## 仕組み

核心は「**AIにレイアウトを描かせない**」こと。38型それぞれのレイアウト（タイトル位置・余白・フォントサイズ・作図ルール）は実測値としてレンダラーに焼き込んであり、AIが書くのは中身（主張・数値・ラベル）だけです。

このJSON（SlideSpec）を書くと——

```json
{
  "template": "waterfall",
  "kicker": "07｜収益ブリッジ",
  "title": "営業利益は価格改定と歩留まり改善で+42を積み、為替の逆風を吸収して140に着地する",
  "chart": {
    "unit": "億円",
    "series": [
      { "label": "FY24実績", "value": 100, "kind": "base" },
      { "label": "価格改定", "value": 18, "kind": "up" },
      { "label": "歩留まり改善", "value": 24, "kind": "up" },
      { "label": "為替影響", "value": -9, "kind": "down" },
      { "label": "その他", "value": 7, "kind": "up" },
      { "label": "FY25見込", "value": 140, "kind": "total" }
    ]
  },
  "sections": [ { "title": "含意", "copy": "寄与の6割は価格改定。為替の逆風は想定レンジ内で吸収できる" } ],
  "note": "注：FY25は10月時点の着地見込",
  "source": "出典：社内管理会計"
}
```

——この1枚が出てきます。

![waterfall example](docs/example_waterfall.png)

ほかの例（`pipeline/slide-spec/example_deck.json` に3枚分を同梱）:

| kpi_dashboard | comparison_table |
| --- | --- |
| ![kpi](docs/example_kpi_dashboard.png) | ![comparison](docs/example_comparison_table.png) |

## セットアップ

```bash
# Claude Codeのスキルフォルダにcloneするだけ
git clone https://github.com/gozen3ji/consulting-pptx-skill.git ~/.claude/skills/consulting-pptx-skill

# パイプラインの初期化（Node.jsが必要。playwright chromiumも入ります）
cd ~/.claude/skills/consulting-pptx-skill/pipeline
npm run setup
```

あとはClaude Codeにこう頼みます:

> 新規事業の投資判断資料を作りたい。38型から型を選んで10枚構成のSlideSpecを書き、パイプラインで編集可能なPPTXまで出して。作成前に slide-rules.md を読み、出力後は check_deck.py でFAIL 0にすること。

## 手動で使う場合

```bash
cd pipeline
node scripts/validate_spec.mjs slide-spec/example_deck.json                              # スキーマ検証
node scripts/render_spec_to_html.mjs slide-spec/example_deck.json generated/deck.html    # HTMLプレビュー
node scripts/qa_html_deck.mjs generated/deck.html                                        # 構造QA
node scripts/export_spec_to_editable_pptx.mjs slide-spec/example_deck.json generated/deck.pptx  # 編集可能PPTX
python3 ../scripts/check_deck.py generated/deck.pptx                                     # 規約の機械チェック
```

書き出されるPPTXは画像貼り付けではなく、全図形がPowerPointで編集できるネイティブなテキストボックス・表・シェイプです。

## 中身

| パス | 内容 |
| --- | --- |
| `SKILL.md` | 38型カタログ＋実運用フロー（AIへの指示書。これがスキルの本体） |
| `pipeline/` | 生成スクリプト（validate / render / qa / export / shots）＋CSS |
| `pipeline/slide-spec/super_template.json` | 38型すべての完成SlideSpec定義（コピーして使う正本） |
| `pipeline/slide-spec/example_deck.json` | 記入例3枚（上のプレビュー画像の元データ） |
| `pipeline/slide-spec/schema.json` | SlideSpecスキーマ |
| `references/slide-rules.md` | スライド作成ルール正典（約80項目） |
| `scripts/check_deck.py` | 規約の機械チェック（PPTX / HTML両対応。要 `pip3 install python-pptx`） |
| `assets/SuperTemplate_38type.pptx` | おまけ: 38型を1枚ずつ収めた見本帳PPTX（[PDF版](assets/SuperTemplate_38type.pdf)） |

Node.jsが無い環境でも、見本帳PPTXの手動コピペとルール正典・機械チェックはそのまま使えます。

## カスタマイズ

- `references/slide-rules.md` に自社の規約を1行ずつ追記していくと、フィードバックが蓄積されて品質が上がっていきます
- SlideSpecルートの `palette` でブランドカラーを一括差し替えできます（`schema.json` 参照）

## About

Made by [Carnot AI](https://jinba.io) — AIエージェント基盤「Jinba」を開発・提供しています。
このスキルと同じ仕組みを、ブラウザのチャットだけで使える形（Jinba App Neo）でも提供しています。

## License

MIT
