---
name: slide-generator
description: input/の元資料からPowerPointスライドを自動生成する。プレゼン資料の作成を依頼されたときに使用する。
---

# slide-generator

## 概要

input/配下の元資料（テキスト、PDF、画像等）を読み取り、構成設計・レイアウト選択・テキスト調整を経て、slides.md（中間Markdown）を作成し、最終的にPowerPoint（.pptx）ファイルを自動生成するスキル。

**入出力フロー**: `input/元資料` → `output/slides.md` → `output/<タイトル>.pptx`

## 前提条件

- Python 3.10以上
- python-pptx ライブラリ
- Meiryo UI フォント（システムにインストール済みであること）
- input/ ディレクトリに元資料が配置されていること

## ワークフロー概要

6ステップで元資料からPowerPointを生成する。詳細は [references/workflow.md](references/workflow.md) を参照。

| ステップ | 内容 | インプット | アウトプット |
|---------|------|-----------|-------------|
| 1. 目的確認 | ターゲット・ゴール・期待アクションを抽出 | input/元資料 | 目的定義 |
| 2. 構成設計 | キーメッセージ抽出・論理構造選択 | 目的定義+元資料 | スライド構成案 |
| 3. レイアウト選択 | 15パターンから最適レイアウトを割り当て | 構成案 | レイアウト割り当て |
| 4. テキスト調整 | 文字数制限に合わせてテキストを調整 | 構成案+レイアウト | 調整済みテキスト |
| 5. slides.md出力 | Markdown形式でスライドを記述 | 調整済みテキスト | output/slides.md |
| 6. PowerPoint生成 | CLIコマンドでPPTXを生成 | slides.md+config.json | output/*.pptx |

論理構造は以下の4パターンから選択する:
- **PREP**: 結論→理由→具体例→結論（汎用的）
- **問題解決**: 課題提示→原因分析→解決策→効果
- **比較提案**: 現状→選択肢比較→推奨案→期待効果
- **時系列**: 過去→現在→未来（変遷・ロードマップ向け）

## レイアウト一覧

15種類のレイアウトから各スライドに最適なものを選択する。
選択基準の詳細は [references/layout-selection-guide.md](references/layout-selection-guide.md) を参照。
各レイアウトの完全な仕様は [references/layout-rules.md](references/layout-rules.md) を参照。

| レイアウト | 用途 | 背景色 |
|-----------|------|--------|
| title | 表紙 | 紺色 `#1E3A5F` |
| section | セクション区切り | 薄グレー `#F5F5F5` |
| toc | 目次・サマリー | 白 |
| bullet_points | 箇条書き（デフォルト） | 白 |
| numbered_list | 手順・ステップ | 白 |
| two_column | 2つの対比 | 白 |
| three_column | 3つの並列 | 白 |
| four_column | 4つの並列 | 白 |
| metrics | 数値・KPI（2-4個） | 白 |
| quote | 引用 | 白 |
| faq | Q&A | 白 |
| comparison_table | 比較表 | 白 |
| image_with_text | 画像＋テキスト | 白 |
| chart | グラフ（bar/horizontal_bar/pie/line） | 白 |
| cta | 行動喚起 | 薄グレー `#F5F5F5` |

## slides.md記法

### 共通ルール

- スライド間は `---` で区切る
- レイアウト指定は `<!-- layout: レイアウト名 -->` で行う（未指定時は bullet_points）
- 見出しレベルの役割:
  - `#`（h1）: セクション名（自動的にsectionレイアウト）
  - `##`（h2）: スライドタイトル
  - `###`（h3）: キーメッセージ
  - `####`（h4）: カラムヘッダー（two/three/four_column用）

### title - 表紙

```markdown
<!-- layout: title -->
## プレゼンテーションタイトル
### サブタイトルや日付・発表者名
```

### section - セクション区切り

```markdown
# セクション名
```

または明示的に指定:

```markdown
<!-- layout: section -->
## セクション名
```

### toc - 目次・サマリー

```markdown
<!-- layout: toc -->
## アジェンダ
### 本日お話しする内容をご紹介します

- 現状分析
- 課題の整理
- 解決策の提案
- ロードマップ
- まとめ
```

### bullet_points - 箇条書き

```markdown
<!-- layout: bullet_points -->
## 主要な成果
### 今期の取り組みにより3つの成果が得られた

- 売上が前年比120%に成長した
- 新規顧客獲得数が目標の150%を達成した
- 顧客満足度が4.5点に向上した
```

### numbered_list - 手順・ステップ

```markdown
<!-- layout: numbered_list -->
## 導入ステップ
### 3ステップで簡単に導入できます

1. 要件定義: ビジネス要件をヒアリングし要件を整理する
2. 環境構築: クラウド環境をセットアップし接続テストを実施する
3. 本番移行: データ移行と動作確認を経て本番稼働を開始する
```

### two_column - 2つの対比

```markdown
<!-- layout: two_column -->
## 導入前後の比較
### システム刷新により業務効率が大幅に改善された

#### Before
手作業でのデータ入力に毎日2時間を費やしていた。

#### After
自動化により入力作業がゼロになった。
```

### three_column - 3つの並列

```markdown
<!-- layout: three_column -->
## サービスプラン
### ニーズに合わせて3つのプランをご用意しています

#### ベーシック
月額1万円で基本機能を利用可能。

#### スタンダード
月額3万円で高度な分析機能を追加。

#### プレミアム
月額10万円で専任サポート付き。
```

### four_column - 4つの並列

```markdown
<!-- layout: four_column -->
## 開発フェーズ
### 4フェーズで段階的にリリースします

#### Phase 1
基盤構築と認証機能の実装

#### Phase 2
コア機能の開発とAPI連携

#### Phase 3
UI改善と性能最適化

#### Phase 4
本番リリースと運用体制確立
```

### metrics - 数値・KPI

```markdown
<!-- layout: metrics -->
## 今期の実績
### 全てのKPIで目標を上回る成果を達成した

| 数値 | ラベル |
|------|--------|
| 120% | 売上達成率 |
| 450件 | 新規契約数 |
| 4.5 | 顧客満足度 |
```

### quote - 引用

```markdown
<!-- layout: quote -->
## お客様の声

> このシステムを導入してから業務効率が劇的に改善しました。

-- 株式会社サンプル 田中太郎氏
```

### faq - Q&A

```markdown
<!-- layout: faq -->
## よくある質問

**Q. 導入にどれくらいの期間がかかりますか？**
A. 標準的な導入期間は約2ヶ月です。

**Q. 既存システムとの連携は可能ですか？**
A. REST APIを提供しており、主要なシステムとの連携実績があります。
```

### comparison_table - 比較表

```markdown
<!-- layout: comparison_table -->
## プラン比較
### 3つのプランから最適なものをお選びいただけます

| 項目 | ベーシック | スタンダード | プレミアム |
|------|-----------|-------------|-----------|
| 月額 | 1万円 | 3万円 | 10万円 |
| ユーザー数 | 5名 | 20名 | 無制限 |
| サポート | メール | チャット | 専任担当 |
```

### image_with_text - 画像＋テキスト

```markdown
<!-- layout: image_with_text -->
## システム構成図
### マイクロサービスアーキテクチャを採用しています

![システム構成図](images/architecture.png)

フロントエンドからAPIゲートウェイを経由して各サービスにアクセスする構成です。
```

### chart - グラフ

横棒グラフ（デフォルト）:

```markdown
<!-- layout: chart -->
## 顧客満足度調査
### 全カテゴリで高い評価を獲得しています

| カテゴリ | スコア |
|---------|-------|
| 使いやすさ | 4.5 |
| **サポート** | 4.8 |
| 機能性 | 4.2 |
```

縦棒グラフ（時系列ラベルで自動判定）:

```markdown
<!-- layout: chart -->
## 売上推移

| 四半期 | 売上(億円) |
|-------|-----------|
| Q1 | 2.5 |
| Q2 | 3.0 |
| **Q4** | 4.5 |
```

折れ線グラフ（明示指定、複数系列対応）:

```markdown
<!-- layout: chart -->
<!-- chart_type: line -->
## ユーザー数推移

| 月 | 登録者 | アクティブ |
|----|-------|----------|
| 1月 | 1000 | 500 |
| 2月 | 1500 | 800 |
```

100%積み上げ棒（%値で自動判定）:

```markdown
<!-- layout: chart -->
## 売上構成比

| カテゴリ | 割合 |
|---------|------|
| サービスA | 45% |
| **サービスB** | 35% |
| サービスC | 20% |
```

### cta - 行動喚起

```markdown
<!-- layout: cta -->
## お問い合わせはこちら
### 導入に関するご質問やデモのご依頼を承っております

- 無料デモを申し込む
```

## チャート自動判定

chartレイアウトではデータの特徴に応じてチャートタイプを自動判定する:

| 条件 | チャートタイプ | 説明 |
|------|--------------|------|
| 値が%表記 or 合計が約100 | pie | 100%積み上げ棒グラフ |
| ラベルが時系列（年/月/Q1等） | bar | 縦棒グラフ |
| `<!-- chart_type: line -->` | line | 折れ線グラフ（複数系列対応） |
| それ以外 | horizontal_bar | 横棒グラフ（デフォルト） |

- テーブルのラベルを `**太字**` にすると、その棒がアクセント色 `#3AA899` で強調表示される
- 折れ線グラフでは3系列まで対応（系列3は `#EDB120`）

## 文字数制限

テキストの表示幅: 全角 = 2、半角 = 1。超過時は末尾を「...」で切り詰める。
レイアウト別の詳細な制限値は [references/character-limits.md](references/character-limits.md) を参照。

キーメッセージ（`###`見出し）のルール:
- 体言止め禁止（文末を述語で終える）
- 20-40文字目安
- 具体的な数値・固有名詞を含める

超過時の対処:
- **要約**: 文言を簡潔にまとめて制限内に収める
- **分割**: 情報量が多い場合は複数スライドに分割する

## デザイン仕様

デザイン設定は [config.json](config.json) で管理する。

### カラーパレット

| 用途 | 色 | コード |
|------|-----|-------|
| プライマリ | 紺色 | `#1E3A5F` |
| セカンダリ | 青灰 | `#4A6FA5` |
| アクセント | ティール | `#3AA899` |
| グレー | 灰色 | `#999999` |
| テキスト（主） | 濃灰 | `#333333` |
| テキスト（副） | 中灰 | `#666666` |
| テキスト（白） | 白 | `#FFFFFF` |
| 背景（主） | 白 | `#FFFFFF` |
| 背景（副） | 薄灰 | `#F5F5F5` |
| 背景（暗） | 紺色 | `#1E3A5F` |
| チャート3色目 | 黄色 | `#EDB120` |

### フォント

- ファミリー: Meiryo UI
- サイズ一覧:

| 用途 | サイズ |
|------|-------|
| タイトル（表紙/セクション） | 44pt |
| セクション見出し | 36pt |
| スライドタイトル | 24pt |
| キーメッセージ | 16pt |
| 本文 | 12pt |

図形描画の詳細仕様は [references/shape-helpers.md](references/shape-helpers.md) を参照。

## CLI引数

### slides_markdown.py（パーサー・シリアライザー）

```
python scripts/slides_markdown.py --input output/slides.md
python scripts/slides_markdown.py --output output/slides.md
```

| 引数 | 必須 | デフォルト | 説明 |
|------|------|-----------|------|
| `--input` | - | - | slides.mdファイルのパス。パース実行し結果をJSON出力 |
| `--output` | - | - | 出力先ファイルパス。stdinのJSONをシリアライズ |

### slide_generator_pptx.py（PPTX生成）

```
python scripts/slide_generator_pptx.py --markdown-file output/slides.md --config config.json
python scripts/slide_generator_pptx.py --markdown-file output/slides.md --config config.json --template assets/template.pptx
```

| 引数 | 必須 | デフォルト | 説明 |
|------|------|-----------|------|
| `--markdown-file` | 必須 | - | slides.mdファイルのパス |
| `--config` | - | `config.json` | config.jsonファイルのパス |
| `--title` | - | - | プレゼンテーションのタイトル |
| `--output-dir` | - | config.jsonのoutput.dir | 出力ディレクトリ |
| `--template` | - | - | テンプレートPPTXのパス |
| `--layout` | - | - | 単一スライドのレイアウト指定 |

## リファレンス

- [ワークフロー詳細](references/workflow.md) - 6ステップの詳細手順
- [レイアウト選択ガイド](references/layout-selection-guide.md) - コンテンツに応じたレイアウト選択
- [レイアウト詳細仕様](references/layout-rules.md) - 15レイアウトの完全な記法と仕様
- [文字数制限ルール](references/character-limits.md) - レイアウト別の文字数制限テーブル
- [図形描画ヘルパー仕様](references/shape-helpers.md) - python-pptx図形描画の技術仕様
