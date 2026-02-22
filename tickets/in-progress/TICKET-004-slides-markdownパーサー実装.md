# TICKET-004: slides_markdown.py パーサー実装

| 項目 | 内容 |
|------|------|
| チケット番号 | TICKET-004 |
| 作成日 | 2026-02-22 |
| 更新日 | 2026-02-22 |
| 担当者 | Claude |
| ブランチ | feature/TICKET-004-slides-markdown-parser |

## 概要

slides.md（独自Markdown記法）をPythonデータ構造に変換するパーサーを実装する。TICKET-003で作成したスケルトンに中身を実装する。

## 背景・目的

slide-generatorスキルのワークフローでは、Claudeが出力したslides.mdをPythonスクリプトでパースしてPPTX生成に渡す。本チケットではその変換処理を実装する。PPTX生成（TICKET-005以降）が依存するため、データ構造の設計が重要である。

## 要件定義

### データクラス定義

- [ ] `ChartData` データクラスを定義する:
  ```python
  @dataclass
  class ChartData:
      chart_type: str           # "bar" | "horizontal_bar" | "pie" | "line"
      labels: list[str]         # 行ラベル（例: ["2021", "2022", "2023"]）
      series: list[dict]        # 系列データ（例: [{"name": "売上", "values": [100, 200, 300]}]）
      highlight_indices: list[int]  # **太字**ラベルのインデックス（強調色表示用）
  ```

- [ ] `SlideContent` データクラスを定義する:
  ```python
  @dataclass
  class SlideContent:
      layout: str               # 15レイアウトのいずれか（デフォルト: "bullet_points"）
      title: str                # ##見出しから抽出。titleレイアウトではプレゼンタイトル
      subtitle: str             # ###見出しから抽出。キーメッセージ。titleレイアウトではサブタイトル
      body: str                 # 本文テキスト（箇条書き・番号リスト等のMarkdown文字列）
      columns: list[dict]       # カラムデータ [{"heading": str, "body": str}, ...]
                                #   two_column: 2要素、three_column: 3要素、four_column: 4要素
      items: list[dict]         # リスト項目 [{"title": str, "description": str}, ...]
                                #   numbered_list: タイトル+説明
                                #   bullet_points: タイトルのみ（descriptionは空文字）
                                #   toc: タイトルのみ
                                #   metrics: {"value": str, "label": str}
      table: list[list[str]]    # テーブルデータ [["ヘッダ1", "ヘッダ2"], ["値1", "値2"], ...]
                                #   comparison_table: 比較表
                                #   chart: グラフ元データ
      chart: ChartData | None   # chartレイアウト時のみ。テーブルから自動変換
      image_path: str           # image_with_textレイアウトの画像パス
      quote: str                # quoteレイアウトの引用文
      quote_author: str         # quoteレイアウトの発言者
      faq_items: list[dict]     # faqレイアウト [{"question": str, "answer": str}, ...]
      cta_button: str           # ctaレイアウトのボタンテキスト
  ```

- [ ] `Presentation` データクラスを定義する:
  ```python
  @dataclass
  class Presentation:
      slides: list[SlideContent]  # スライドのリスト
      title: str                  # プレゼンテーション全体のタイトル（最初のtitleスライドから取得）
  ```

### スライド分割処理

- [ ] `---`（3つ以上のハイフンのみの行）でスライドを分割する
- [ ] 先頭・末尾の空白スライド（内容が空のもの）は除外する
- [ ] `---`の前後の空行は無視する

### レイアウト抽出

- [ ] `<!-- layout: xxx -->` コメントからレイアウト名を抽出する
  - 正規表現: `<!--\s*layout:\s*(\w+)\s*-->`
  - 大文字小文字を区別しない
- [ ] レイアウトコメントがない場合のデフォルト: `bullet_points`
- [ ] `#`（h1）見出しで始まるスライドは自動的に `section` レイアウトとして扱う（layoutコメントより優先）
- [ ] 不明なレイアウト名はログに警告を出し、`bullet_points` にフォールバックする

### 見出し抽出

- [ ] `# 見出し`（h1）: sectionレイアウトのタイトルとして抽出
- [ ] `## 見出し`（h2）: スライドタイトル（`SlideContent.title`）として抽出
- [ ] `### 見出し`（h3）: キーメッセージ（`SlideContent.subtitle`）として抽出
- [ ] `#### 見出し`（h4）: カラムヘッダー（`SlideContent.columns[].heading`）として抽出
- [ ] 各見出しの`#`とテキストの間の空白を正しく処理する

### 15レイアウトのパース仕様

#### 1. title（表紙）
- [ ] `## タイトル` → `title`
- [ ] `### サブタイトル` → `subtitle`
- [ ] その他の本文は `body` に格納

#### 2. section（セクション区切り）
- [ ] `# セクション名` → `title`（h1で自動判定）
- [ ] `<!-- layout: section -->` + `## セクション名` でも可

#### 3. toc（目次）
- [ ] `## タイトル` → `title`
- [ ] `### キーメッセージ` → `subtitle`
- [ ] `- 項目` のリスト → `items` に `{"title": "項目", "description": ""}` として格納

#### 4. bullet_points（箇条書き）
- [ ] `## タイトル` → `title`
- [ ] `### キーメッセージ` → `subtitle`
- [ ] `- 項目` のリスト → `items` に `{"title": "項目", "description": ""}` として格納
- [ ] ネストした箇条書き（`  - サブ項目`）はサブ項目として `description` に格納

#### 5. numbered_list（手順・ステップ）
- [ ] `## タイトル` → `title`
- [ ] `### キーメッセージ` → `subtitle`
- [ ] `1. **項目タイトル**\n   説明文` → `items` に `{"title": "項目タイトル", "description": "説明文"}` として格納
- [ ] `**太字**` からタイトルを抽出し、続く行を説明として扱う

#### 6. two_column（2カラム）
- [ ] `## タイトル` → `title`
- [ ] `### キーメッセージ` → `subtitle`
- [ ] `#### 左見出し` + 本文 → `columns[0]` = `{"heading": "左見出し", "body": "本文"}`
- [ ] `#### 右見出し` + 本文 → `columns[1]` = `{"heading": "右見出し", "body": "本文"}`

#### 7. three_column（3カラム）
- [ ] two_columnと同様だが `columns` が3要素
- [ ] `####` 見出しが3つ必要

#### 8. four_column（4カラム）
- [ ] two_columnと同様だが `columns` が4要素
- [ ] `####` 見出しが4つ必要

#### 9. metrics（数値・KPI）
- [ ] `## タイトル` → `title`
- [ ] `### キーメッセージ` → `subtitle`
- [ ] `- **数値** ラベル` → `items` に `{"value": "数値", "label": "ラベル"}` として格納
- [ ] `**`で囲まれた部分を数値として抽出、残りをラベルとする

#### 10. quote（引用）
- [ ] `## タイトル` → `title`
- [ ] `> 引用文` → `quote` に格納
- [ ] `> — 発言者` → `quote_author` に格納（`—` または `--` で始まる行を発言者として判定）

#### 11. faq（Q&A）
- [ ] `## タイトル` → `title`
- [ ] `**Q: 質問文**` → `faq_items[].question`
- [ ] `A: 回答文` → `faq_items[].answer`
- [ ] Q/Aのペアを順に `faq_items` リストに格納

#### 12. comparison_table（比較表）
- [ ] `## タイトル` → `title`
- [ ] `### キーメッセージ` → `subtitle`
- [ ] Markdownテーブル → `table` に2次元リストとして格納
  - 1行目: ヘッダー行
  - 2行目: 区切り行（`|---|---|`）は除外
  - 3行目以降: データ行
- [ ] セル内の `**太字**` はそのまま保持（PPTX生成時に太字として描画）

#### 13. image_with_text（画像＋テキスト）
- [ ] `## タイトル` → `title`
- [ ] `### キーメッセージ` → `subtitle`
- [ ] `![alt](パス)` → `image_path` にパスを格納
- [ ] 画像の後の本文 → `body` に格納

#### 14. chart（グラフ）
- [ ] `## タイトル` → `title`
- [ ] `### キーメッセージ` → `subtitle`
- [ ] `<!-- chart_type: xxx -->` コメントからチャートタイプを抽出（オプション）
- [ ] Markdownテーブルをパースし `table` に格納
- [ ] テーブルデータからチャートタイプを自動判定し `chart` (ChartData) を生成:
  - **判定アルゴリズム**:
    1. `<!-- chart_type: line -->` が指定されている場合 → `line`
    2. データ値に `%` が含まれる、または全値の合計が95-105の範囲 → `pie`
    3. ラベルが時系列パターンに一致 → `bar`
       - 時系列パターン: `\d{4}年?`, `\d{1,2}月`, `Q[1-4]`, `\d{4}/\d{1,2}`, `FY\d{2,4}` 等
    4. 上記いずれにも該当しない → `horizontal_bar`
  - **系列データの変換**:
    - テーブルの1列目をラベル（`labels`）として使用
    - テーブルの2列目以降を系列（`series`）として使用
    - 各系列の `name` はヘッダー行の該当列名
    - 各系列の `values` はデータ行の該当列の数値（カンマ・単位を除去してfloatに変換）
  - **太字ラベルの検出**:
    - ラベル列で `**太字**` のものを検出
    - そのインデックスを `highlight_indices` に格納

#### 15. cta（行動喚起）
- [ ] `## メッセージ` → `title`
- [ ] `### 補足説明` → `subtitle`
- [ ] `[ボタンテキスト]` → `cta_button` に格納（`[`と`]`で囲まれたテキスト。リンク `[text](url)` とは区別する）

### 文字数カウント・バリデーション

- [ ] `count_display_width(text: str) -> int` 関数を実装する:
  - 全角文字（East Asian Width が W, F のもの）= 2
  - 半角文字 = 1
  - Pythonの `unicodedata.east_asian_width()` を使用
  - Markdown記法（`**`, `*`, `#` 等）は除去してからカウントする

- [ ] `validate_character_limits(presentation: Presentation) -> list[str]` 関数を実装する:
  - references/character-limits.md の制限テーブルに基づきチェック
  - 超過時は警告メッセージのリストを返す
  - 警告メッセージ形式: `"スライド{n} ({layout}): {フィールド}が制限{limit}に対して{actual}です"`
  - 全スライドを走査し、全ての超過を検出する

### SlidesMarkdownSerializer（逆変換）

- [ ] `Presentation` → slides.md形式のMarkdown文字列に変換する
- [ ] 全15レイアウトの逆変換に対応する
- [ ] パース → シリアライズ → パースのラウンドトリップで情報が失われないこと

### CLIインターフェース

- [ ] `--input <path>`: slides.mdファイルをパースし、結果をJSON形式で標準出力に出力する
  - JSON出力は `dataclasses.asdict()` を使用
  - インデント付き（`indent=2`）で読みやすく出力
  - エンコーディング: UTF-8
- [ ] `--validate <path>`: slides.mdファイルの文字数制限をチェックし、警告を標準エラー出力に出力する
  - 警告がある場合は終了コード1、ない場合は終了コード0
- [ ] エラー処理:
  - ファイルが存在しない場合: エラーメッセージを表示して終了コード1
  - パースエラー: エラー箇所（行番号）を含むメッセージを表示して終了コード1

## 設計

### 変更対象

| ファイル | 変更内容 |
|----------|----------|
| `skills/slide-generator/scripts/slides_markdown.py` | TICKET-003で作成したスケルトンに中身を実装。データクラス定義、パーサー、シリアライザー、文字数関連関数、CLIの全てを実装する |

### 設計方針

#### モジュール構成

`slides_markdown.py` は単一ファイルとし、以下の順序で定義する:

1. **モジュールdocstring・import文**
2. **定数定義**: レイアウト名一覧、正規表現パターン、文字数制限テーブル
3. **データクラス**: `ChartData`, `SlideContent`, `Presentation`
4. **ユーティリティ関数**: `count_display_width()`, `validate_character_limits()`
5. **SlidesMarkdownParser クラス**
6. **SlidesMarkdownSerializer クラス**
7. **CLI（main関数）**

#### パーサーの処理フロー

```
入力: slides.md (文字列)
  ↓
1. "---" でスライドブロックに分割
  ↓
2. 各ブロックについて:
   a. <!-- layout: xxx --> を抽出（なければデフォルト判定）
   b. # (h1) で始まる場合は layout = "section" に上書き
   c. ## (h2) からタイトルを抽出
   d. ### (h3) からキーメッセージを抽出
   e. レイアウト別のパース処理を呼び出し:
      - _parse_title(), _parse_section(), _parse_toc(), ...
      - 各メソッドが SlideContent を返す
  ↓
3. Presentation オブジェクトを組み立て
  ↓
出力: Presentation
```

#### レイアウト別パーサーメソッド

`SlidesMarkdownParser` に以下のプライベートメソッドを定義する:

| メソッド | 対応レイアウト | 主な処理 |
|----------|-------------|----------|
| `_parse_title(lines)` | title | ##と###の抽出 |
| `_parse_section(lines)` | section | #の抽出 |
| `_parse_toc(lines)` | toc | ##, ###, 箇条書きの抽出 |
| `_parse_bullet_points(lines)` | bullet_points | ##, ###, 箇条書き（ネスト対応）の抽出 |
| `_parse_numbered_list(lines)` | numbered_list | ##, ###, 番号付きリスト（太字タイトル+説明文）の抽出 |
| `_parse_columns(lines, n)` | two/three/four_column | ##, ###, ####×n個の抽出 |
| `_parse_metrics(lines)` | metrics | ##, ###, `**値** ラベル` の抽出 |
| `_parse_quote(lines)` | quote | ##, `>`引用, `— 著者` の抽出 |
| `_parse_faq(lines)` | faq | ##, `**Q:**`/`A:` ペアの抽出 |
| `_parse_comparison_table(lines)` | comparison_table | ##, ###, Markdownテーブルの抽出 |
| `_parse_image_with_text(lines)` | image_with_text | ##, ###, `![]()`画像, 本文の抽出 |
| `_parse_chart(lines)` | chart | ##, ###, テーブル→ChartData変換（自動判定含む） |
| `_parse_cta(lines)` | cta | ##, ###, `[ボタン]`の抽出 |

カラム系レイアウト（two_column, three_column, four_column）は `_parse_columns(lines, n)` で共通化し、引数 `n` でカラム数を制御する。

#### チャートタイプ自動判定の詳細

`_detect_chart_type(table, explicit_type)` メソッド:

```python
def _detect_chart_type(self, table: list[list[str]], explicit_type: str | None) -> str:
    """テーブルデータからチャートタイプを自動判定する

    判定優先順位:
    1. explicit_type（<!-- chart_type: xxx -->）が指定されている場合はそれを使用
    2. データ値の判定: %表記 or 合計≒100 → "pie"
    3. ラベルの判定: 時系列パターン → "bar"
    4. デフォルト → "horizontal_bar"
    """
```

時系列パターンの正規表現:
```python
TIME_SERIES_PATTERNS = [
    r'^\d{4}年?$',           # 2021, 2021年
    r'^\d{1,2}月$',          # 1月, 12月
    r'^Q[1-4]$',             # Q1, Q2, Q3, Q4
    r'^\d{4}[/-]\d{1,2}$',  # 2021/1, 2021-01
    r'^FY\d{2,4}$',          # FY21, FY2021
    r'^\d{4}年度$',          # 2021年度
]
```

#### 数値抽出の処理

テーブルのデータセルから数値を抽出する `_extract_number(text)` メソッド:
- `**太字**` のマークダウン記法を除去
- カンマ区切りを除去（例: `1,000` → `1000`）
- 単位を除去（例: `100万円` → `100`、`50%` → `50`）
- `float()` で変換できない場合は `0.0` を返す

#### python-pptxへの依存

- パーサー（slides_markdown.py）はpython-pptxに依存しない
- 標準ライブラリのみ使用: `dataclasses`, `re`, `json`, `argparse`, `sys`, `unicodedata`, `typing`
- これにより単体テストがpython-pptxなしで実行可能

#### エラーハンドリング方針

- パースエラーは可能な限り寛容に処理する（壊れた部分をスキップして続行）
- 致命的なエラー（ファイル読み込み失敗等）のみ例外をraiseする
- 警告は `logging` モジュールで出力する（`logging.warning()`）
- CLIの `--validate` は警告を集約して一括出力する

## 受け入れ条件

- [ ] 全15レイアウトのslides.md記法を正しくパースできる
- [ ] `SlideContent`, `ChartData`, `Presentation` データクラスが設計通りに定義されている
- [ ] チャートタイプ自動判定が4パターン（bar, horizontal_bar, pie, line）全て動作する
  - [ ] `<!-- chart_type: line -->` 指定時 → `line`
  - [ ] %値データ → `pie`
  - [ ] 時系列ラベル → `bar`
  - [ ] その他 → `horizontal_bar`
- [ ] テーブルの `**太字**` ラベルが `highlight_indices` に正しく反映される
- [ ] `count_display_width()` が全角/半角混在文字列で正しい幅を返す
- [ ] `validate_character_limits()` が文字数超過を正しく検出する
- [ ] `SlidesMarkdownSerializer` によるラウンドトリップ（パース→シリアライズ→パース）で情報が保持される
- [ ] CLIの `--input` でパース結果がJSON出力される
- [ ] CLIの `--validate` で文字数制限チェックが実行される
- [ ] python-pptxに依存せず、標準ライブラリのみで動作する
- [ ] PRが作成され、mainにマージされている

## レビューメモ

<!-- レビュー時のコメント・指摘事項を記録 -->
