# TICKET-008: 統合テスト・最終調整

| 項目 | 内容 |
|------|------|
| チケット番号 | TICKET-008 |
| 作成日 | 2026-02-22 |
| 更新日 | 2026-02-22 |
| 担当者 | Claude |
| ブランチ | feature/TICKET-008-integration-test |

## 概要

slide-generatorスキルの全15レイアウトを含むE2Eテストを実施し、SKILL.mdのワークフローが正しく機能することを確認する。サンプル入力からの完全なPPTX生成、エラーハンドリング、ワークフローの手動実行を検証する。

## 背景・目的

TICKET-003〜007で段階的に構築したスキルの全体動作を検証し、品質を担保する。個別チケットの受け入れ条件は各レイアウト単体の動作確認だが、本チケットでは全レイアウトを1つのプレゼンテーションにまとめた実運用に近いテストを行い、レイアウト間の競合やパーサーの境界条件なども検出する。

## 要件定義

### E2Eテスト用サンプルslides.md

- [ ] 全15レイアウトを含むサンプル `skills/slide-generator/tests/e2e_sample.md` を作成する
- [ ] 各レイアウトの記法が仕様通りに記述されていること
- [ ] チャートレイアウトは4種（horizontal_bar, bar, line, pie）をそれぞれ含むこと（計4スライド）
- [ ] 日本語コンテンツを使用すること（全角文字カウントの検証を兼ねる）
- [ ] 太字ラベル（`**太字**`）のあるチャートデータを含むこと
- [ ] 画像レイアウトはplaceholder.pngを参照すること

サンプルslides.mdの具体的な内容:

```markdown
<!-- layout: title -->
## スライド生成スキル E2Eテスト
### 全15レイアウト動作検証

---

# セクション1: 基本レイアウト

---

<!-- layout: toc -->
## 目次
### 本テストの構成
- 基本レイアウト（title, section, toc, bullet_points, numbered_list, cta）
- カラムレイアウト（two_column, three_column, four_column）
- テーブル・特殊レイアウト（metrics, comparison_table, faq, quote, image_with_text）
- チャートレイアウト（horizontal_bar, bar, line, pie）

---

## 箇条書きレイアウト
### デフォルトレイアウトの動作確認
- 箇条書き項目1: テキストが正しく表示される
- 箇条書き項目2: インデントが適用される
  - サブ項目2-1
  - サブ項目2-2
- 箇条書き項目3: 日本語テキストの表示確認

---

<!-- layout: numbered_list -->
## 手順リスト
### 番号付きリストの表示確認
1. **ステップ1**
   最初の手順の説明文。太字タイトルと説明文の組み合わせ。
2. **ステップ2**
   2番目の手順の説明文。
3. **ステップ3**
   3番目の手順の説明文。

---

<!-- layout: cta -->
## お問い合わせはこちら
### 詳しい情報をご希望の方へ
[資料をダウンロード]

---

# セクション2: カラムレイアウト

---

<!-- layout: two_column -->
## 2カラム比較
### 左右に分割して情報を整理
#### メリット
- コスト削減
- 効率化
- 品質向上
#### デメリット
- 初期投資
- 学習コスト

---

<!-- layout: three_column -->
## 3カラム構成
### 3つの視点から解説
#### 現状
現在の課題と背景を説明します。
#### 提案
解決策の概要を提示します。
#### 効果
期待される成果を示します。

---

<!-- layout: four_column -->
## 4カラム構成
### プロセスの全体像
#### 企画
アイデア出しと要件定義
#### 設計
技術設計とUI設計
#### 開発
実装とテスト
#### 運用
デプロイと監視

---

# セクション3: テーブル・特殊レイアウト

---

<!-- layout: metrics -->
## 主要KPI
### 2025年度の成果
- **120%** 売上成長率
- **4.8** 顧客満足度
- **98%** サービス稼働率
- **50%** コスト削減率

---

<!-- layout: comparison_table -->
## プラン比較
### 最適なプランをお選びください
| 項目 | スタンダード | プロフェッショナル | エンタープライズ |
|------|-------------|-------------------|-----------------|
| ユーザー数 | 10名まで | 100名まで | 無制限 |
| ストレージ | 10GB | 100GB | 1TB |
| サポート | メール | メール+チャット | 24/7専任 |
| 価格 | 月額1,000円 | 月額5,000円 | 要相談 |

---

<!-- layout: faq -->
## よくある質問
**Q: 導入にどのくらい時間がかかりますか？**
A: 標準的な導入は2〜4週間です。カスタマイズの程度により変動します。

**Q: 既存システムとの連携は可能ですか？**
A: REST APIを提供しており、主要なシステムとの連携が可能です。

**Q: データのセキュリティは？**
A: SOC2 Type2認証取得済み。データは暗号化して保存されます。

---

<!-- layout: quote -->
## お客様の声
> このツールを導入してから、資料作成の時間が半分になりました。
> チーム全体の生産性が大幅に向上しています。
> — 株式会社サンプル 田中太郎 様

---

<!-- layout: image_with_text -->
## システム構成図
### クラウドネイティブアーキテクチャ
![システム構成図](../../assets/placeholder.png)
マイクロサービスアーキテクチャを採用し、各サービスが独立してスケール可能な構成です。

---

# セクション4: チャートレイアウト

---

<!-- layout: chart -->
## 顧客満足度ランキング
### カテゴリ別の評価スコア
| カテゴリ | スコア |
|----------|--------|
| 使いやすさ | 85 |
| **サポート品質** | 92 |
| 機能充実度 | 78 |
| コスパ | 88 |
| **総合評価** | 90 |

---

<!-- layout: chart -->
## 売上推移
### 過去5年で3倍に成長
| 年度 | 売上（億円） |
|------|-------------|
| 2021 | 100 |
| 2022 | 150 |
| 2023 | 200 |
| 2024 | 250 |
| **2025** | 300 |

---

<!-- layout: chart -->
<!-- chart_type: line -->
## 月次トレンド
### 売上・利益・コストの推移
| 月 | 売上 | 利益 | コスト |
|----|------|------|--------|
| 1月 | 100 | 20 | 80 |
| 2月 | 110 | 25 | 85 |
| 3月 | 130 | 35 | 95 |
| 4月 | 120 | 30 | 90 |
| 5月 | 140 | 40 | 100 |
| 6月 | 150 | 45 | 105 |

---

<!-- layout: chart -->
## 市場シェア
### 当社が業界トップシェア
| 企業 | シェア |
|------|--------|
| **当社** | 35% |
| A社 | 25% |
| B社 | 20% |
| その他 | 20% |
```

### E2Eテストの実行

- [ ] サンプルslides.mdからPPTXが正常に生成されること

```bash
# テンプレートなしで実行
python skills/slide-generator/scripts/slide_generator_pptx.py \
  --markdown-file skills/slide-generator/tests/e2e_sample.md \
  --config skills/slide-generator/config.json \
  --title "E2Eテスト" \
  --output-dir output/e2e-test

# テンプレートありで実行（テンプレートが存在する場合）
python skills/slide-generator/scripts/slide_generator_pptx.py \
  --markdown-file skills/slide-generator/tests/e2e_sample.md \
  --config skills/slide-generator/config.json \
  --title "E2Eテスト_テンプレート" \
  --output-dir output/e2e-test \
  --template skills/slide-generator/assets/template.pptx

# 特定レイアウトのみで実行
python skills/slide-generator/scripts/slide_generator_pptx.py \
  --markdown-file skills/slide-generator/tests/e2e_sample.md \
  --config skills/slide-generator/config.json \
  --title "E2Eテスト_chart_only" \
  --output-dir output/e2e-test \
  --layout chart
```

- [ ] 生成されたPPTXファイルが破損していないこと（python-pptxで再読み込みして検証）

```bash
# PPTX検証スクリプト（テスト内で実行）
python -c "
from pptx import Presentation
prs = Presentation('output/e2e-test/E2Eテスト.pptx')
print(f'スライド数: {len(prs.slides)}')
for i, slide in enumerate(prs.slides):
    print(f'  スライド{i+1}: シェイプ数={len(slide.shapes)}')
print('検証OK: PPTXファイルは正常に読み込めます')
"
```

- [ ] スライド数が期待値と一致すること（18スライド: title=1, section=4, toc=1, bullet_points=1, numbered_list=1, cta=1, two_column=1, three_column=1, four_column=1, metrics=1, comparison_table=1, faq=1, quote=1, image_with_text=1, chart=4 の合計で18+セクション4 = サンプル内容に準じた枚数）

### パーサー単体テスト

- [ ] slides_markdown.pyのCLIデバッグモードでパース結果を検証する

```bash
# パース結果をJSON出力
python skills/slide-generator/scripts/slides_markdown.py \
  --input skills/slide-generator/tests/e2e_sample.md

# 期待される出力の確認ポイント:
# - 全スライドのlayoutが正しく判定されている
# - chartスライドのchart_typeが正しく判定されている
#   - 顧客満足度 → horizontal_bar（時系列でなく%でもない）
#   - 売上推移 → bar（年度=時系列ラベル）
#   - 月次トレンド → line（chart_typeコメント指定）
#   - 市場シェア → pie（%値）
# - 太字ラベルのbold_labelsインデックスが正しい
# - テーブルデータが正しくパースされている
```

### エラーハンドリングテスト

各ケースで適切なエラーメッセージが表示され、スクリプトがクラッシュしないことを確認する。

- [ ] **空のslides.md**
  - 入力: 空ファイル
  - 期待動作: エラーメッセージ「スライドが見つかりません」を表示して終了（exit code 1）

- [ ] **不正なレイアウト名**
  - 入力: `<!-- layout: unknown_layout -->`
  - 期待動作: 警告メッセージを出力し、デフォルトの `bullet_points` レイアウトにフォールバック

- [ ] **チャートデータなしのchartレイアウト**
  - 入力: `<!-- layout: chart -->` の後にテーブルがない
  - 期待動作: 警告メッセージを出力し、タイトルとキーメッセージのみのスライドを生成

- [ ] **存在しない画像パス**
  - 入力: `![alt](nonexistent.png)`
  - 期待動作: 警告メッセージを出力し、placeholder.pngで代替

- [ ] **存在しないconfig.jsonパス**
  - 入力: `--config nonexistent.json`
  - 期待動作: エラーメッセージを表示して終了（exit code 1）

- [ ] **存在しないslides.mdパス**
  - 入力: `--markdown-file nonexistent.md`
  - 期待動作: エラーメッセージを表示して終了（exit code 1）

- [ ] **文字数制限超過**
  - 入力: タイトルに100文字以上の文字列
  - 期待動作: 警告メッセージ「タイトルが文字数制限を超えています」を標準エラーに出力し、処理は継続する

- [ ] **テーブルの数値パースエラー**
  - 入力: チャートのテーブルで数値列に文字列が含まれる（例: `| 2025 | N/A |`）
  - 期待動作: 警告メッセージを出力し、該当値を0として処理を継続

- [ ] **--layoutに存在しないレイアウト名を指定**
  - 入力: `--layout nonexistent`
  - 期待動作: 警告メッセージを出力し、該当レイアウトのスライドがないPPTXを生成

- [ ] **slides.mdにスライド区切り（---）がない**
  - 入力: 区切りなしのMarkdown
  - 期待動作: 全体を1枚のスライドとしてパースする

### SKILL.mdワークフロー手動テスト

SKILL.mdに記載された6ステップのワークフローを実際に実行し、手順通りに動作することを確認する。

- [ ] ステップ1: input/に元資料を配置
- [ ] ステップ2: 元資料を分析してslides.mdを生成（Claude手動操作）
- [ ] ステップ3: slides.mdの内容を確認・修正
- [ ] ステップ4: slides_markdown.pyでパース・検証
- [ ] ステップ5: slide_generator_pptx.pyでPPTX生成
- [ ] ステップ6: 出力ファイルの確認

### バグ修正・調整

- [ ] E2Eテストで発見されたバグをその場で修正する
- [ ] SKILL.mdの手順に不明瞭な箇所があれば改善する
- [ ] パフォーマンス上の問題（15レイアウト全体の生成時間が極端に遅い等）があれば対処する

## 設計

### 変更対象

| ファイル | 変更内容 |
|----------|----------|
| `skills/slide-generator/tests/e2e_sample.md` | 新規作成。全15レイアウトのサンプルslides.md |
| `skills/slide-generator/scripts/slides_markdown.py` | バグ修正・調整（テスト結果に応じて） |
| `skills/slide-generator/scripts/slide_generator_pptx.py` | バグ修正・調整（テスト結果に応じて） |
| `skills/slide-generator/SKILL.md` | ワークフロー手順の微調整（テスト結果に応じて） |

### 設計方針

- テストは手動実行（自動テストフレームワークは使用しない）
- テスト結果はチケットのレビューメモに記録する
- 発見されたバグはその場で修正し、修正内容をコミットに含める
- PPTX生成結果はoutput/e2e-test/に保存する
- テスト用サンプルファイルは skills/slide-generator/tests/ ディレクトリに配置する

### テスト実行の手順

```bash
# 1. testsディレクトリの作成
mkdir -p skills/slide-generator/tests

# 2. サンプルslides.mdの配置（上記の内容）
# → skills/slide-generator/tests/e2e_sample.md

# 3. パーサー単体テスト
python skills/slide-generator/scripts/slides_markdown.py \
  --input skills/slide-generator/tests/e2e_sample.md

# 4. PPTX生成テスト（テンプレートなし）
python skills/slide-generator/scripts/slide_generator_pptx.py \
  --markdown-file skills/slide-generator/tests/e2e_sample.md \
  --config skills/slide-generator/config.json \
  --title "E2Eテスト" \
  --output-dir output/e2e-test

# 5. PPTX生成テスト（レイアウト絞り込み）
python skills/slide-generator/scripts/slide_generator_pptx.py \
  --markdown-file skills/slide-generator/tests/e2e_sample.md \
  --config skills/slide-generator/config.json \
  --title "E2Eテスト_chart" \
  --output-dir output/e2e-test \
  --layout chart

# 6. PPTX検証
python -c "
from pptx import Presentation
prs = Presentation('output/e2e-test/E2Eテスト.pptx')
print(f'スライド数: {len(prs.slides)}')
for i, slide in enumerate(prs.slides):
    shapes = slide.shapes
    print(f'  スライド{i+1}: シェイプ数={len(shapes)}')
    for shape in shapes:
        if shape.has_text_frame:
            text = shape.text_frame.text[:50]
            print(f'    テキスト: {text}')
        if shape.has_chart:
            print(f'    チャート: {shape.chart.chart_type}')
print('検証完了')
"

# 7. エラーハンドリングテスト
echo "" > /tmp/empty.md
python skills/slide-generator/scripts/slide_generator_pptx.py \
  --markdown-file /tmp/empty.md \
  --config skills/slide-generator/config.json \
  --title "空テスト" \
  --output-dir output/e2e-test
# → エラーメッセージが表示されることを確認

python skills/slide-generator/scripts/slide_generator_pptx.py \
  --markdown-file nonexistent.md \
  --config skills/slide-generator/config.json \
  --title "存在しないファイル" \
  --output-dir output/e2e-test
# → エラーメッセージが表示されることを確認

python skills/slide-generator/scripts/slide_generator_pptx.py \
  --markdown-file skills/slide-generator/tests/e2e_sample.md \
  --config nonexistent.json \
  --title "存在しないconfig" \
  --output-dir output/e2e-test
# → エラーメッセージが表示されることを確認
```

### 品質チェックリスト

テスト完了時に以下の項目を全て確認する:

**機能面:**
- [ ] 全15レイアウトがPPTX内に正しく生成されている
- [ ] チャートの4種（horizontal_bar, bar, line, pie）が正しいタイプで描画されている
- [ ] 太字ラベルにアクセント色（#3AA899）が適用されている
- [ ] テンプレートなし/ありの両方で動作する
- [ ] --layout オプションでレイアウトの絞り込みが動作する
- [ ] パーサーのJSON出力が正しい構造を持つ

**デザイン面:**
- [ ] 色がconfig.jsonのカラーパレットに準拠している
- [ ] フォントがconfig.jsonの指定通り（デフォルト: Meiryo UI）
- [ ] スライドサイズが16:9
- [ ] タイトルスライドの背景が紺色（#1E3A5F）
- [ ] セクションスライドの背景が薄グレー
- [ ] 棒グラフが角丸
- [ ] グリッド線が破線・#E5E5E5
- [ ] 折れ線グラフのマーカーが白塗り+枠線付き
- [ ] メトリクスの数値が36pt bold

**エラーハンドリング面:**
- [ ] 10件のエラーケースが全て適切に処理される
- [ ] エラー時にスタックトレースがユーザーに見えない
- [ ] 警告は標準エラーに出力される
- [ ] 致命的エラー時は exit code 1 で終了する

**互換性面:**
- [ ] 生成されたPPTXがpython-pptxで再読み込みできる
- [ ] OOXML直接操作がPPTXの構造を破損させていない

## 受け入れ条件

- [ ] 全15レイアウトを含むPPTXが正常に生成される
- [ ] パーサーのJSON出力でチャートタイプ判定が4パターン全て正しい
- [ ] 10件のエラーハンドリングテストケースが全て合格する
- [ ] SKILL.mdの6ステップワークフローが手動で実行可能であることを確認済み
- [ ] 品質チェックリストの全項目を確認済み
- [ ] 発見されたバグが全て修正済み
- [ ] PRが作成され、mainにマージされている

## レビューメモ

<!-- レビュー時のコメント・指摘事項を記録 -->
