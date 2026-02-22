# TICKET-003: slide-generator スキル基盤構築

| 項目 | 内容 |
|------|------|
| チケット番号 | TICKET-003 |
| 作成日 | 2026-02-22 |
| 更新日 | 2026-02-22 |
| 担当者 | Claude |
| ブランチ | feature/TICKET-003-slide-generator-foundation |

## 概要

PowerPointスライド自動生成スキル（slide-generator）のスキル定義・設定・リファレンスファイルを構築する。スクリプト実装（TICKET-004以降）の土台となる。

## 背景・目的

input/の元資料からpython-pptxでPPTXを自動生成するスキルを段階的に構築する。本チケットはその第1段階として、Claudeがワークフローを実行するためのスキル定義と参照資料を整備する。

## 要件定義

### SKILL.md

- [ ] 500行以下で作成する
- [ ] 以下のセクションを含む:
  - **概要**: スキルの目的と入出力の説明（input/ → slides.md → PPTX）
  - **前提条件**: Python 3.10+、python-pptx、Meiryo UIフォント
  - **ワークフロー概要**: 6ステップの簡潔な説明（詳細はreferences/workflow.mdへリンク）
    1. 目的確認
    2. 構成設計
    3. レイアウト選択
    4. テキスト調整
    5. slides.md出力
    6. PowerPoint生成
  - **レイアウト一覧**: 15パターンの名前・用途・簡潔な説明テーブル（詳細はreferences/layout-rules.mdへリンク）
  - **slides.md記法**: 共通ルール（`---`区切り、`<!-- layout: xxx -->`、見出しレベルの役割）と各レイアウトの記法例（詳細はreferences/layout-rules.mdへリンク）
  - **チャート自動判定**: 4パターンの判定ルール概要
  - **文字数制限**: 概要と計算ルール（全角=2、半角=1）。レイアウト別制限テーブルはreferences/character-limits.mdへリンク
  - **デザイン仕様**: カラーパレット・フォント・サイズの概要。config.jsonへの参照
  - **CLI引数**: 2スクリプトの引数一覧
  - **リファレンス**: 各references/ファイルへのリンク一覧
- [ ] 各セクションの行数配分目安:
  - 概要・前提条件: 約20行
  - ワークフロー概要: 約30行
  - レイアウト一覧: 約40行
  - slides.md記法（共通ルール＋全レイアウト例）: 約200行
  - チャート自動判定: 約20行
  - 文字数制限概要: 約20行
  - デザイン仕様: 約40行
  - CLI引数: 約40行
  - リファレンス: 約10行

### config.json

- [ ] 以下の構造で作成する:
```json
{
  "output": {
    "dir": "output"
  },
  "font": {
    "family": "Meiryo UI"
  },
  "palette": {
    "primary": "#1E3A5F",
    "secondary": "#4A6FA5",
    "accent": "#3AA899",
    "gray": "#999999",
    "text": {
      "primary": "#333333",
      "secondary": "#666666",
      "light": "#FFFFFF"
    },
    "background": {
      "primary": "#FFFFFF",
      "secondary": "#F5F5F5",
      "dark": "#1E3A5F"
    },
    "chart": {
      "line_color_3": "#EDB120"
    }
  },
  "template": {
    "pptx_path": "assets/template.pptx"
  }
}
```

### リファレンスファイル

#### references/workflow.md

- [ ] 6ステップワークフローの詳細手順を記載する:
  1. **目的確認**: input/配下の全ファイルを読み込み、以下を抽出する。不明な場合はユーザーに質問する
     - ターゲット（誰に向けたプレゼンか）
     - ゴール（プレゼンのゴール）
     - 期待アクション（聞いた人に何をしてほしいか）
  2. **構成設計**: 以下の手順で構成を決定する
     - キーメッセージの抽出
     - 情報のグルーピング
     - 論理構造の選択（PREP / 問題解決 / 比較提案 / 時系列）
     - 各スライドの役割決定
  3. **レイアウト選択**: 各スライドに15パターンから最適レイアウトを割り当てる。layout-selection-guide.mdを参照
  4. **テキスト調整**: character-limits.mdに基づき文字数制限に合わせて調整する。超過時は要約または分割で対応
  5. **slides.md出力**: layout-rules.mdの記法に従いoutput/slides.mdに出力する
  6. **PowerPoint生成**: CLIコマンドでPPTXを生成する。具体的なコマンド例を記載
     ```
     python scripts/slide_generator_pptx.py --markdown-file output/slides.md --config config.json
     ```
- [ ] 各ステップのインプット・アウトプットを明記する
- [ ] ステップ間の依存関係を明記する

#### references/layout-selection-guide.md

- [ ] レイアウト選択のフローチャートを記載する:
  - コンテンツの種類による分岐:
    - 表紙 → title
    - セクション区切り → section
    - 全体の流れを示す → toc
    - 数値・KPIを強調 → metrics
    - 引用を紹介 → quote
    - Q&A形式 → faq
    - 行動を促す → cta
    - 画像が主役 → image_with_text
    - データをグラフ化 → chart
    - 比較表 → comparison_table
    - 手順・ステップ → numbered_list
    - 2つの対比 → two_column
    - 3つの並列 → three_column
    - 4つの並列 → four_column
    - 上記以外 → bullet_points（デフォルト）
- [ ] 各レイアウトの「いつ使うか」「使わない方がよい場面」を記載する
- [ ] レイアウト選択のベストプラクティスを記載する:
  - 同じレイアウトの連続は2回まで
  - title/section/ctaは1プレゼンで各1回が目安
  - 全体のリズムを意識してレイアウトを変化させる

#### references/layout-rules.md

- [ ] 15レイアウトそれぞれについて以下を記載する:
  - レイアウト名と用途
  - slides.md記法の完全な例
  - 使用可能な要素（タイトル、キーメッセージ、箇条書き、テーブル等）
  - 背景色ルール（title: 紺色背景`#1E3A5F`、section: 薄グレー`#F5F5F5`、cta: 薄グレー`#F5F5F5`、その他: 白`#FFFFFF`）
  - デザイン上の注意点
- [ ] 以下の15レイアウトを全て網羅する:
  1. title - 表紙（紺色背景）
  2. section - セクション区切り（薄グレー背景）
  3. toc - 目次・サマリー
  4. bullet_points - 箇条書き（デフォルト）
  5. numbered_list - 手順・ステップ
  6. two_column - 2つの対比
  7. three_column - 3つの並列
  8. four_column - 4つの並列
  9. metrics - 数値・KPI（2-4個）
  10. quote - 引用
  11. faq - Q&A
  12. comparison_table - 比較表
  13. image_with_text - 画像＋テキスト
  14. chart - グラフ（bar/horizontal_bar/pie/line）
  15. cta - 行動喚起（薄グレー背景）
- [ ] 見出しレベルの共通ルールを冒頭に記載:
  - `#`（h1）: セクション名。自動的にsectionレイアウトとして扱う
  - `##`（h2）: スライドタイトル
  - `###`（h3）: キーメッセージ
  - `####`（h4）: カラムヘッダー（two/three/four_column用）
- [ ] chartレイアウトではチャートタイプ自動判定ルールも記載:
  - 値が%表記 or 合計≒100 → pie（100%積み上げ棒）
  - ラベルが時系列（年/月/Q1等）→ bar（縦棒）
  - `<!-- chart_type: line -->` → line（折れ線、複数系列対応）
  - それ以外 → horizontal_bar（横棒）
  - テーブルのラベルを`**太字**`にすると強調色で表示
- [ ] 100行超になるため冒頭に目次を入れる

#### references/character-limits.md

- [ ] 文字数計算ルールを記載する:
  - 全角 = 2、半角 = 1
  - 超過時は末尾を「...」で切り詰め
- [ ] キーメッセージのルールを記載する:
  - 体言止め禁止
  - 20-40文字目安
  - 具体的な数値・固有名詞を含める
- [ ] レイアウト別文字数制限テーブルを記載する:

| レイアウト | タイトル | キーメッセージ | 項目/セル | 推奨数 |
|-----------|---------|-------------|----------|-------|
| title | 30 | 50(subtitle) | - | - |
| section | 20 | - | - | - |
| toc | 25 | 40 | 1項目20 | 3-7 |
| bullet_points | 25 | 40 | 1項目50 | 3-5 |
| numbered_list | 25 | 40 | タイトル15/説明50 | 3-5 |
| two_column | 25 | 40 | 見出し10/本文80 | - |
| three_column | 25 | 40 | 見出し8/本文60 | - |
| four_column | 25 | 40 | 見出し6/本文40 | - |
| metrics | 25 | 40 | 数値10/ラベル15 | 2-4 |
| quote | 25 | - | 引用100/著者20 | - |
| faq | 25 | - | Q30/A80 | 2-3 |
| comparison_table | 25 | 40 | ヘッダ10/データ15 | 3-5行×3-5列 |
| image_with_text | 25 | 40 | 説明100 | - |
| chart | 25 | 40 | ラベル10 | 3-7 |
| cta | 20(msg) | 40(desc) | ボタン15 | - |

- [ ] 超過時の対処方法（要約 or スライド分割）を記載する

#### references/shape-helpers.md

- [ ] python-pptxでの図形描画ヘルパーの仕様を記載する:
  - 角丸矩形（ROUNDED_RECTANGLE）の生成方法
  - テキストフレームの設定（フォント、サイズ、行間1.5倍、配置）
  - カラーパレットからの色取得方法（config.json参照）
  - グリッド線描画（破線、色`#E5E5E5`）
  - グラフ描画仕様:
    - 棒グラフ: 角丸ROUNDED_RECTANGLE、通常=グレー(`#999999`)、強調=アクセント色(`#3AA899`)
    - 縦棒グラフ(bar): 時系列ラベル用
    - 横棒グラフ(horizontal_bar): デフォルト
    - 折れ線グラフ(line): 線幅2.5pt、円形マーカー8pt白枠付き、3系列目は`#EDB120`
    - 100%積み上げ棒(pie): %値・合計≒100のデータ用
  - スライドサイズ: 16:9
  - フォントサイズ一覧:
    - タイトル: 44pt
    - セクション: 36pt
    - スライドタイトル: 24pt
    - キーメッセージ: 16pt
    - 本文: 12pt
  - テンプレート対応ルール:
    - テンプレート指定の優先順位: CLIの`--template` > config.jsonの`template.pptx_path` > 指定なし（白背景で自動生成）
    - テンプレート使用時: スライドマスタ・レイアウトを保持、既存スライドは削除、プレースホルダは除去、背景色はテンプレートを尊重
    - テンプレート未使用時: 白紙プレゼンを新規作成、背景色はコードで設定

### assets/placeholder.png

- [ ] 1280x720px（16:9）の灰色プレースホルダー画像を配置する
- [ ] image_with_textレイアウトでの代替画像として使用する
- [ ] ファイルサイズは最小限に抑える（単色PNG）

### スクリプトスケルトン

#### scripts/slides_markdown.py

- [ ] 以下のCLI引数を定義する:
  - `--input`: slides.mdファイルのパス（パース実行、結果をJSON出力）
  - `--output`: 出力先ファイルパス（シリアライズ時）
- [ ] 以下のクラス・関数のスケルトン（docstring + `pass` or `raise NotImplementedError`）を定義する:
  - `SlideContent` データクラス
  - `ChartData` データクラス
  - `Presentation` データクラス
  - `SlidesMarkdownParser` クラス（`parse()` メソッド）
  - `SlidesMarkdownSerializer` クラス（`serialize()` メソッド）
  - `count_display_width(text: str) -> int` 関数
  - `validate_character_limits(presentation: Presentation) -> list[str]` 関数
- [ ] main関数でargparseを使用したCLIエントリポイントを定義する
- [ ] 冒頭にモジュールdocstringを記載する

#### scripts/slide_generator_pptx.py

- [ ] 以下のCLI引数を定義する:
  - `--markdown-file`: slides.mdファイルのパス（必須）
  - `--config`: config.jsonファイルのパス（デフォルト: `config.json`）
  - `--title`: プレゼンテーションのタイトル（オプション）
  - `--output-dir`: 出力ディレクトリ（デフォルト: config.jsonのoutput.dir）
  - `--template`: テンプレートPPTXのパス（オプション）
  - `--layout`: 単一スライドのレイアウト指定（オプション）
- [ ] 以下のクラス・関数のスケルトンを定義する:
  - `SlideGeneratorPptx` クラス
    - `__init__(self, config: dict)` メソッド
    - `generate(self, presentation: Presentation, output_path: str)` メソッド
    - 15レイアウトに対応する `_render_xxx(self, slide, content: SlideContent)` メソッド群
  - `load_config(config_path: str) -> dict` 関数
- [ ] main関数でargparseを使用したCLIエントリポイントを定義する
- [ ] 冒頭にモジュールdocstringを記載する

## 設計

### 変更対象

| ファイル | 変更内容 |
|----------|----------|
| `skills/slide-generator/SKILL.md` | 新規作成。スキル定義（500行以下）。6ステップワークフロー概要、15レイアウト一覧とslides.md記法例、チャート自動判定、文字数制限概要、デザイン仕様、CLI引数を記載 |
| `skills/slide-generator/config.json` | 新規作成。カラーパレット（primary/secondary/accent/gray/text/background/chart）、フォント（Meiryo UI）、出力ディレクトリ、テンプレートパスを定義 |
| `skills/slide-generator/references/workflow.md` | 新規作成。6ステップワークフローの詳細手順。各ステップのインプット・アウトプット・具体的な作業内容を記載 |
| `skills/slide-generator/references/layout-selection-guide.md` | 新規作成。コンテンツ種類に応じたレイアウト選択フローチャート。各レイアウトの適用場面・ベストプラクティスを記載 |
| `skills/slide-generator/references/layout-rules.md` | 新規作成。15レイアウト全ての詳細仕様。slides.md記法の完全な例、使用可能要素、背景色ルール、デザイン注意点を記載。目次付き |
| `skills/slide-generator/references/character-limits.md` | 新規作成。文字数計算ルール（全角=2、半角=1）、キーメッセージルール、レイアウト別文字数制限テーブル、超過時対処法を記載 |
| `skills/slide-generator/references/shape-helpers.md` | 新規作成。python-pptx図形描画ヘルパー仕様。角丸矩形、テキストフレーム、グラフ描画（4種）、フォントサイズ、テンプレート対応ルールを記載 |
| `skills/slide-generator/assets/placeholder.png` | 新規作成。1280x720px灰色プレースホルダー画像 |
| `skills/slide-generator/scripts/slides_markdown.py` | 新規作成。CLIスケルトン。SlideContent/ChartData/Presentationデータクラス、Parser/Serializerクラス、文字数関連関数のスケルトン |
| `skills/slide-generator/scripts/slide_generator_pptx.py` | 新規作成。CLIスケルトン。SlideGeneratorPptxクラスと15レイアウト描画メソッド、config読み込み関数のスケルトン |

### 設計方針

- **SKILL.md**: Agent Skills標準準拠。500行以下を維持するため、詳細仕様はreferences/に分離する。SKILL.mdには概要と各記法の使用例を記載し、references/への明示的なリンクを設置する
- **config.json**: ユーザー要件の仕様に完全準拠。トップレベルキーは `output`, `font`, `palette`, `template` の4つ。パレットはネスト構造（text.primary等）を使用する
- **references/**: SKILL.mdから1階層のみの参照関係。100行超のファイルには冒頭に目次（`## 目次`セクション）を入れる。各ファイルは独立して読めるよう、必要な前提情報を冒頭に記載する
- **スクリプトスケルトン**: 本チケットではCLI引数定義、クラス・関数のdocstring、型ヒント付きシグネチャのみを実装する。関数本体は `raise NotImplementedError("TICKET-004で実装")` 等とし、実装はTICKET-004以降で行う
- **placeholder.png**: Pythonの `Pillow` または最小限のPNGバイナリで1280x720の単色灰色画像を生成する。手動配置でもよい

### ディレクトリ構造

```
skills/slide-generator/
├── SKILL.md                          # スキル定義（500行以下）
├── config.json                       # デザイン設定
├── assets/
│   ├── placeholder.png               # プレースホルダー画像（1280x720）
│   └── template.pptx                 # PPTXテンプレート（本チケット対象外）
├── references/
│   ├── workflow.md                   # 6ステップワークフロー詳細
│   ├── layout-selection-guide.md     # レイアウト選択ガイド
│   ├── layout-rules.md              # 15レイアウト詳細仕様（目次付き）
│   ├── character-limits.md           # 文字数制限ルール
│   └── shape-helpers.md             # 図形描画ヘルパー仕様
└── scripts/
    ├── slides_markdown.py            # パーサースケルトン
    └── slide_generator_pptx.py       # ジェネレータースケルトン
```

## 受け入れ条件

- [ ] SKILL.mdが500行以下で全セクション（概要、前提条件、ワークフロー概要、レイアウト一覧、slides.md記法、チャート判定、文字数制限、デザイン仕様、CLI引数、リファレンス）を網羅している
- [ ] config.jsonが正しいJSON形式であり、要件定義のスキーマ通りである
- [ ] 5つのリファレンスファイルが全て作成されている
- [ ] 100行超のリファレンスファイル（layout-rules.md）に目次がある
- [ ] SKILL.mdから各リファレンスへのリンクが正しく、相対パスで参照できる
- [ ] スクリプトスケルトン2ファイルがPython構文エラーなくインポート可能である（`python -c "import slides_markdown"` が成功する）
- [ ] placeholder.pngが1280x720pxの画像である
- [ ] PRが作成され、mainにマージされている

## レビューメモ

<!-- レビュー時のコメント・指摘事項を記録 -->
