# TICKET-006: カラム・テーブルレイアウト実装

| 項目 | 内容 |
|------|------|
| チケット番号 | TICKET-006 |
| 作成日 | 2026-02-22 |
| 更新日 | 2026-02-22 |
| 担当者 | Claude |
| ブランチ | feature/TICKET-006-column-table-layouts |

## 概要

slide_generator_pptx.pyにカラム系・テーブル系・特殊レイアウト計8種のRendererを追加する。TICKET-005で構築したBaseRendererを継承し、より複雑な構造化コンテンツを描画する。

## 背景・目的

TICKET-005で構築した基盤（DesignConfig, BaseRenderer, CLI）の上に、コンテンツの構造化が必要な8レイアウトを追加する。カラム分割、テーブル描画、メトリクス表示、引用、FAQ、画像配置など、ビジネスプレゼンテーションで頻出するレイアウトパターンを網羅する。

## 要件定義

### カラム系レイアウト

- [ ] **TwoColumnRenderer**: 2カラムレイアウト
  - 背景: 白色（`#FFFFFF`）
  - スライドタイトル: 24pt, bold
  - キーメッセージ: 16pt
  - 2カラム: コンテンツ領域を左右均等分割（間にガター0.3in）
  - 各カラム: ヘッダー（14pt, bold, primary色）+ 本文（12pt）
  - カラム間に縦の区切り線（オプション、薄グレー）
  - 文字数制限: タイトル25文字、キーメッセージ40文字、見出し10文字、本文80文字

- [ ] **ThreeColumnRenderer**: 3カラムレイアウト
  - 背景: 白色
  - スライドタイトル: 24pt, bold
  - キーメッセージ: 16pt
  - 3カラム: コンテンツ領域を3等分（ガター0.25in × 2）
  - 各カラム: ヘッダー（14pt, bold, primary色）+ 本文（12pt）
  - 文字数制限: タイトル25文字、キーメッセージ40文字、見出し8文字、本文60文字

- [ ] **FourColumnRenderer**: 4カラムレイアウト
  - 背景: 白色
  - スライドタイトル: 24pt, bold
  - キーメッセージ: 16pt
  - 4カラム: コンテンツ領域を4等分（ガター0.2in × 3）
  - 各カラム: ヘッダー（12pt, bold, primary色）+ 本文（11pt）
  - 文字数制限: タイトル25文字、キーメッセージ40文字、見出し6文字、本文40文字

### メトリクス・数値レイアウト

- [ ] **MetricsRenderer**: 数値・KPIレイアウト
  - 背景: 白色
  - スライドタイトル: 24pt, bold
  - キーメッセージ: 16pt
  - メトリクスカード: 2〜4個を横並びで均等配置
  - 各カード: 数値（36pt, bold, accent色）+ ラベル（12pt, text_secondary色）
  - カード間に縦の区切り線（薄グレー、1pt）
  - 数値は中央揃え
  - 文字数制限: タイトル25文字、キーメッセージ40文字、数値10文字、ラベル15文字

### テーブルレイアウト

- [ ] **ComparisonTableRenderer**: 比較表レイアウト
  - 背景: 白色
  - スライドタイトル: 24pt, bold
  - キーメッセージ: 16pt
  - テーブル: python-pptxの`add_table()`で描画
  - ヘッダー行: 背景色primary（`#1E3A5F`）、テキスト白色、12pt bold
  - データ行: 交互背景色（白 `#FFFFFF` / 薄グレー `#F5F5F5`）、テキスト `#333333`、11pt
  - 太字ラベル（`**text**`）: accent色（`#3AA899`）で表示
  - セル内余白: 上下左右0.05in
  - 推奨サイズ: 3〜5行 × 3〜5列
  - 文字数制限: タイトル25文字、キーメッセージ40文字、ヘッダ10文字、データ15文字

### 特殊レイアウト

- [ ] **FaqRenderer**: Q&Aレイアウト
  - 背景: 白色
  - スライドタイトル: 24pt, bold
  - Q&Aペア: 縦並びで2〜3組
  - Q（質問）: 14pt, bold, primary色（`#1E3A5F`）、先頭に "Q. " プレフィックス
  - A（回答）: 12pt, text_primary色（`#333333`）、先頭に "A. " プレフィックス
  - Q&Aペア間に水平区切り線（薄グレー、1pt）
  - 文字数制限: タイトル25文字、Q30文字、A80文字

- [ ] **QuoteRenderer**: 引用レイアウト
  - 背景: 白色
  - スライドタイトル: 24pt, bold
  - 左側にaccent色（`#3AA899`）の縦線（装飾バー）
  - 引用文: 16pt, italic, text_primary色
  - 出典/著者: 12pt, text_secondary色, 右寄せ、先頭に "-- " プレフィックス
  - 文字数制限: タイトル25文字、引用100文字、著者20文字

- [ ] **ImageWithTextRenderer**: 画像+テキストレイアウト
  - 背景: 白色
  - スライドタイトル: 24pt, bold
  - キーメッセージ: 16pt
  - 左半分: 画像（`add_picture()`）。画像なし時は `assets/placeholder.png` を使用
  - 右半分: テキスト説明（12pt）
  - 画像はアスペクト比を維持してフィット
  - 文字数制限: タイトル25文字、キーメッセージ40文字、説明100文字

### 共通要件

- [ ] 8レイアウト全てを `RENDERER_MAP` に追加登録
- [ ] 既存の6レイアウト（TICKET-005）が壊れないこと
- [ ] BaseRendererの共通メソッド（`_render_slide_title`, `_render_key_message`等）を適切に再利用

## 設計

### 変更対象

| ファイル | 変更内容 |
|----------|----------|
| `skills/slide-generator/scripts/slide_generator_pptx.py` | 8レイアウトRendererクラスの追加、RENDERER_MAPへの追加登録 |

### 設計方針

#### 共通座標（TICKET-005のBaseRendererから継承）

| 領域 | Left | Top | Width | Height |
|------|------|-----|-------|--------|
| スライドタイトル | 0.5in | 0.3in | 12.333in | 0.6in |
| キーメッセージ | 0.5in | 0.95in | 12.333in | 0.45in |
| コンテンツ領域 | 0.5in | 1.55in | 12.333in | 5.45in |

全レイアウトでスライドタイトルとキーメッセージは `_render_slide_title()` と `_render_key_message()` を使用する（titleとsectionを除く）。

#### TwoColumnRenderer 設計

```
背景: #FFFFFF

┌──────────────────────────────────────────────┐
│  [スライドタイトル: 24pt]                      │
│  [キーメッセージ: 16pt]                        │
│  ─────────────────────────────────            │
│                                              │
│  ┌──────────┐  │  ┌──────────┐               │
│  │ ヘッダー1  │  │  │ ヘッダー2 │               │
│  │           │  │  │          │               │
│  │ 本文1     │  │  │ 本文2    │               │
│  │           │  │  │          │               │
│  └──────────┘  │  └──────────┘               │
│                                              │
└──────────────────────────────────────────────┘
```

コンテンツ領域（12.333in幅）をガター0.3inで2分割:
- カラム幅: `(12.333 - 0.3) / 2 = 6.017in`

| 要素 | Left | Top | Width | Height | フォント | 色 |
|------|------|-----|-------|--------|---------|-----|
| カラム1ヘッダー | 0.5in | 1.55in | 6.017in | 0.4in | 14pt, bold | #1E3A5F |
| カラム1本文 | 0.5in | 2.0in | 6.017in | 4.5in | 12pt | #333333 |
| 区切り線（縦） | 6.667in | 1.55in | 1pt | 5.45in | - | #CCCCCC |
| カラム2ヘッダー | 6.817in | 1.55in | 6.017in | 0.4in | 14pt, bold | #1E3A5F |
| カラム2本文 | 6.817in | 2.0in | 6.017in | 4.5in | 12pt | #333333 |

- 縦区切り線: `MSO_SHAPE.RECTANGLE` で幅1ptの細い矩形として描画
- ヘッダーの下にaccent色の装飾線（幅2.0in、太さ2pt）を配置

#### ThreeColumnRenderer 設計

```
背景: #FFFFFF

┌──────────────────────────────────────────────┐
│  [スライドタイトル: 24pt]                      │
│  [キーメッセージ: 16pt]                        │
│  ─────────────────────────────────            │
│                                              │
│  ┌────────┐  ┌────────┐  ┌────────┐          │
│  │ヘッダー1│  │ヘッダー2│  │ヘッダー3│          │
│  │        │  │        │  │        │          │
│  │ 本文1  │  │ 本文2  │  │ 本文3  │          │
│  │        │  │        │  │        │          │
│  └────────┘  └────────┘  └────────┘          │
│                                              │
└──────────────────────────────────────────────┘
```

コンテンツ領域（12.333in幅）をガター0.25in × 2で3分割:
- カラム幅: `(12.333 - 0.25 * 2) / 3 = 3.944in`

| 要素 | Left | Top | Width | Height | フォント | 色 |
|------|------|-----|-------|--------|---------|-----|
| カラム1ヘッダー | 0.5in | 1.55in | 3.944in | 0.4in | 14pt, bold | #1E3A5F |
| カラム1本文 | 0.5in | 2.0in | 3.944in | 4.5in | 12pt | #333333 |
| カラム2ヘッダー | 4.694in | 1.55in | 3.944in | 0.4in | 14pt, bold | #1E3A5F |
| カラム2本文 | 4.694in | 2.0in | 3.944in | 4.5in | 12pt | #333333 |
| カラム3ヘッダー | 8.889in | 1.55in | 3.944in | 0.4in | 14pt, bold | #1E3A5F |
| カラム3本文 | 8.889in | 2.0in | 3.944in | 4.5in | 12pt | #333333 |

- カラム座標の算出: `left = CONTENT_LEFT + col_index * (col_width + gutter)`
  - col0: 0.5in
  - col1: 0.5 + 3.944 + 0.25 = 4.694in
  - col2: 0.5 + 2 * (3.944 + 0.25) = 8.889in
- ヘッダーの下にaccent色の装飾線（カラム幅全体、太さ2pt）を配置

#### FourColumnRenderer 設計

```
背景: #FFFFFF

┌──────────────────────────────────────────────┐
│  [スライドタイトル: 24pt]                      │
│  [キーメッセージ: 16pt]                        │
│  ─────────────────────────────────            │
│                                              │
│  ┌──────┐ ┌──────┐ ┌──────┐ ┌──────┐         │
│  │ hdr1 │ │ hdr2 │ │ hdr3 │ │ hdr4 │         │
│  │      │ │      │ │      │ │      │         │
│  │ body │ │ body │ │ body │ │ body │         │
│  │      │ │      │ │      │ │      │         │
│  └──────┘ └──────┘ └──────┘ └──────┘         │
│                                              │
└──────────────────────────────────────────────┘
```

コンテンツ領域（12.333in幅）をガター0.2in × 3で4分割:
- カラム幅: `(12.333 - 0.2 * 3) / 4 = 2.933in`

| 要素 | Left | Top | Width | Height | フォント | 色 |
|------|------|-----|-------|--------|---------|-----|
| カラムNヘッダー | 計算値 | 1.55in | 2.933in | 0.35in | 12pt, bold | #1E3A5F |
| カラムN本文 | 計算値 | 1.95in | 2.933in | 4.5in | 11pt | #333333 |

- カラム座標: `left = CONTENT_LEFT + col_index * (col_width + gutter)`
  - col0: 0.5in
  - col1: 0.5 + 2.933 + 0.2 = 3.633in
  - col2: 0.5 + 2 * (2.933 + 0.2) = 6.766in
  - col3: 0.5 + 3 * (2.933 + 0.2) = 9.899in
- フォントサイズをやや小さく（ヘッダー12pt、本文11pt）してスペース確保

#### カラム系共通の実装方針

カラム系3レイアウトは共通のパターンに従うため、BaseRendererに以下のヘルパーメソッドを追加する:

```python
def _render_columns(self, slide, columns_data: list, config: DesignConfig,
                    gutter: float, header_font_size=Pt(14),
                    body_font_size=Pt(12)):
    """カラムレイアウトの共通描画処理"""
    num_cols = len(columns_data)
    col_width = (self.CONTENT_WIDTH - gutter * (num_cols - 1)) / num_cols

    for i, col_data in enumerate(columns_data):
        col_left = Inches(self.CONTENT_LEFT + i * (col_width + gutter))
        # ヘッダー描画
        self.add_text_box(
            slide,
            left=col_left,
            top=Inches(self.CONTENT_TOP),
            width=Inches(col_width),
            height=Inches(0.4),
            text=col_data.get("header", ""),
            font_size=header_font_size,
            font_color=config.primary_color,
            bold=True,
            font_family=config.font_family,
        )
        # ヘッダー下の装飾線
        self.add_horizontal_line(
            slide,
            left=col_left,
            top=Inches(self.CONTENT_TOP + 0.45),
            width=Inches(col_width),
            color=config.accent_color,
            thickness=2.0,
        )
        # 本文描画
        self.add_text_box(
            slide,
            left=col_left,
            top=Inches(self.CONTENT_TOP + 0.55),
            width=Inches(col_width),
            height=Inches(self.CONTENT_HEIGHT - 0.55),
            text=col_data.get("body", ""),
            font_size=body_font_size,
            font_color=config.text_primary,
            font_family=config.font_family,
        )
```

#### MetricsRenderer 設計

```
背景: #FFFFFF

┌──────────────────────────────────────────────┐
│  [スライドタイトル: 24pt]                      │
│  [キーメッセージ: 16pt]                        │
│  ─────────────────────────────────            │
│                                              │
│    ┌────────┐ │ ┌────────┐ │ ┌────────┐      │
│    │ 98.5%  │ │ │ ¥12M   │ │ │ 1,234  │      │
│    │  KPI1  │ │ │  KPI2  │ │ │  KPI3  │      │
│    └────────┘ │ └────────┘ │ └────────┘      │
│                                              │
│                                              │
└──────────────────────────────────────────────┘
```

メトリクス数に応じてコンテンツ領域を等分:
- 2個: 幅6.017in × 2（ガター0.3in）
- 3個: 幅3.944in × 3（ガター0.25in）
- 4個: 幅2.933in × 4（ガター0.2in）

| 要素 | Left | Top | Width | Height | フォント | 色 |
|------|------|-----|-------|--------|---------|-----|
| 数値テキスト | 各カード中央 | 2.5in | カード幅 | 1.2in | 36pt, bold, center | #3AA899 |
| ラベルテキスト | 各カード中央 | 3.8in | カード幅 | 0.5in | 12pt, center | #666666 |
| カード間区切り線（縦） | カード間 | 2.2in | 1pt | 2.5in | - | #CCCCCC |

- 数値は中央揃え（`PP_ALIGN.CENTER`）
- 数値のTop: コンテンツ領域の上部から約1.0in下（`1.55 + 0.95 = 2.5in`）
- ラベルのTop: 数値の下（`2.5 + 1.3 = 3.8in`）
- 区切り線: 各カードの間に配置（カード数-1本）

```python
def render(self, slide, slide_data, config):
    self._render_slide_title(slide, slide_data.get("title", ""), config)
    if slide_data.get("key_message"):
        self._render_key_message(slide, slide_data["key_message"], config)

    metrics = slide_data.get("metrics", [])
    num = len(metrics)
    if num < 2:
        num = 2
    if num > 4:
        num = 4

    # ガターとカード幅の計算
    gutter = {2: 0.3, 3: 0.25, 4: 0.2}.get(num, 0.25)
    card_width = (self.CONTENT_WIDTH - gutter * (num - 1)) / num

    for i, metric in enumerate(metrics[:4]):
        card_left = Inches(self.CONTENT_LEFT + i * (card_width + gutter))
        # 数値
        self.add_text_box(
            slide, left=card_left, top=Inches(2.5),
            width=Inches(card_width), height=Inches(1.2),
            text=metric.get("value", ""),
            font_size=Pt(36), font_color=config.accent_color,
            bold=True, alignment=PP_ALIGN.CENTER,
            font_family=config.font_family,
        )
        # ラベル
        self.add_text_box(
            slide, left=card_left, top=Inches(3.8),
            width=Inches(card_width), height=Inches(0.5),
            text=metric.get("label", ""),
            font_size=Pt(12), font_color=config.text_secondary,
            alignment=PP_ALIGN.CENTER,
            font_family=config.font_family,
        )
        # 区切り線（最後のカード以外）
        if i < num - 1:
            line_left = Inches(self.CONTENT_LEFT + (i + 1) * card_width + i * gutter + gutter / 2)
            self.add_horizontal_line(  # 縦線なのでadd_shapeで代替
                slide, left=line_left, top=Inches(2.2),
                width=Pt(1), color="#CCCCCC", thickness=2.5 * 72,  # 高さ2.5in
            )
```

注: 縦の区切り線はadd_horizontal_lineでは描画できないため、BaseRendererに `add_vertical_line()` メソッドを追加する:

```python
def add_vertical_line(self, slide, left, top, height,
                      color: str = "#CCCCCC", thickness: float = 1.0):
    """縦線を描画（細い矩形として実装）"""
    shape = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE,
        left, top, Pt(thickness), height
    )
    shape.fill.solid()
    shape.fill.fore_color.rgb = RGBColor.from_string(color.lstrip("#"))
    shape.line.fill.background()
    return shape
```

#### ComparisonTableRenderer 設計

```
背景: #FFFFFF

┌──────────────────────────────────────────────┐
│  [スライドタイトル: 24pt]                      │
│  [キーメッセージ: 16pt]                        │
│  ─────────────────────────────────            │
│                                              │
│  ┌────────┬────────┬────────┬────────┐        │
│  │ Header │ Header │ Header │ Header │ ← 紺色│
│  ├────────┼────────┼────────┼────────┤        │
│  │ data   │ data   │ data   │ data   │ ← 白  │
│  ├────────┼────────┼────────┼────────┤        │
│  │ data   │ data   │ data   │ data   │ ← 灰  │
│  ├────────┼────────┼────────┼────────┤        │
│  │ data   │ data   │ data   │ data   │ ← 白  │
│  └────────┴────────┴────────┴────────┘        │
│                                              │
└──────────────────────────────────────────────┘
```

| 要素 | Left | Top | Width | Height | フォント | 色 |
|------|------|-----|-------|--------|---------|-----|
| テーブル全体 | 0.5in | 1.55in | 12.333in | 自動計算 | - | - |
| ヘッダー行 | - | - | - | 0.5in | 12pt, bold | 背景: #1E3A5F, 文字: #FFFFFF |
| データ行（奇数） | - | - | - | 0.45in | 11pt | 背景: #FFFFFF, 文字: #333333 |
| データ行（偶数） | - | - | - | 0.45in | 11pt | 背景: #F5F5F5, 文字: #333333 |
| 太字セル | - | - | - | - | 11pt, bold | 文字: #3AA899 |

python-pptx テーブルAPI:
```python
def render(self, slide, slide_data, config):
    self._render_slide_title(slide, slide_data.get("title", ""), config)
    if slide_data.get("key_message"):
        self._render_key_message(slide, slide_data["key_message"], config)

    table_data = slide_data.get("table", {})
    headers = table_data.get("headers", [])
    rows = table_data.get("rows", [])
    num_rows = len(rows) + 1  # ヘッダー含む
    num_cols = len(headers)

    # テーブル作成
    table_height = Inches(0.5 + 0.45 * len(rows))  # ヘッダー + データ行
    table = slide.shapes.add_table(
        num_rows, num_cols,
        left=Inches(self.CONTENT_LEFT),
        top=Inches(self.CONTENT_TOP),
        width=Inches(self.CONTENT_WIDTH),
        height=table_height,
    ).table

    # ヘッダー行のスタイル設定
    for col_idx, header_text in enumerate(headers):
        cell = table.cell(0, col_idx)
        cell.text = header_text
        # 背景色: primary
        cell.fill.solid()
        cell.fill.fore_color.rgb = RGBColor.from_string(
            config.primary_color.lstrip("#"))
        # フォント
        for paragraph in cell.text_frame.paragraphs:
            paragraph.font.size = Pt(12)
            paragraph.font.bold = True
            paragraph.font.color.rgb = RGBColor.from_string("FFFFFF")
            paragraph.font.name = config.font_family
            paragraph.alignment = PP_ALIGN.CENTER
        # セル余白
        cell.margin_left = Inches(0.05)
        cell.margin_right = Inches(0.05)
        cell.margin_top = Inches(0.05)
        cell.margin_bottom = Inches(0.05)

    # データ行のスタイル設定
    for row_idx, row_data in enumerate(rows):
        is_even_row = (row_idx % 2 == 1)  # 0始まりで偶数行が白、奇数行がグレー
        bg_color = config.bg_secondary if is_even_row else config.bg_primary

        for col_idx, cell_text in enumerate(row_data):
            cell = table.cell(row_idx + 1, col_idx)
            # 太字ラベル（**text**）の検出
            is_bold_label = cell_text.startswith("**") and cell_text.endswith("**")
            if is_bold_label:
                cell_text = cell_text[2:-2]
            cell.text = cell_text
            # 背景色
            cell.fill.solid()
            cell.fill.fore_color.rgb = RGBColor.from_string(bg_color.lstrip("#"))
            # フォント
            for paragraph in cell.text_frame.paragraphs:
                paragraph.font.size = Pt(11)
                paragraph.font.name = config.font_family
                if is_bold_label:
                    paragraph.font.bold = True
                    paragraph.font.color.rgb = RGBColor.from_string(
                        config.accent_color.lstrip("#"))
                else:
                    paragraph.font.color.rgb = RGBColor.from_string(
                        config.text_primary.lstrip("#"))
            # セル余白
            cell.margin_left = Inches(0.05)
            cell.margin_right = Inches(0.05)
            cell.margin_top = Inches(0.05)
            cell.margin_bottom = Inches(0.05)

    # テーブル罫線の設定（外枠・内部線）
    # python-pptxではテーブルの罫線はXML直接操作が必要
    # → ShapeHelpersで対応（references/shape-helpers.mdを参照）
```

#### FaqRenderer 設計

```
背景: #FFFFFF

┌──────────────────────────────────────────────┐
│  [スライドタイトル: 24pt]                      │
│  ─────────────────────────────────            │
│                                              │
│  Q. 質問テキスト1                              │
│  A. 回答テキスト1                              │
│                                              │
│  ──────────────────────────────               │
│                                              │
│  Q. 質問テキスト2                              │
│  A. 回答テキスト2                              │
│                                              │
│  ──────────────────────────────               │
│                                              │
│  Q. 質問テキスト3                              │
│  A. 回答テキスト3                              │
│                                              │
└──────────────────────────────────────────────┘
```

Q&Aペア数に応じてコンテンツ領域を等分:
- 2ペア: 各2.725in（5.45 / 2）
- 3ペア: 各1.817in（5.45 / 3）

| 要素 | Left | Top | Width | Height | フォント | 色 |
|------|------|-----|-------|--------|---------|-----|
| スライドタイトル | 0.5in | 0.3in | 12.333in | 0.6in | 24pt, bold | #333333 |
| Q テキスト | 0.5in | 計算値 | 12.333in | 0.4in | 14pt, bold | #1E3A5F |
| A テキスト | 0.5in | Q下 | 12.333in | 残り | 12pt | #333333 |
| 区切り線 | 0.5in | ペア間 | 12.333in | - | - | #CCCCCC, 1pt |

- 各Q&Aペアの開始位置: `CONTENT_TOP + pair_index * pair_height`
- Qテキストは "Q. " プレフィックス付き、先頭Runをprimary色bold
- Aテキストは "A. " プレフィックス付き、Qの直下に配置
- 区切り線: 各ペアの下部（最終ペア以外）

```python
def render(self, slide, slide_data, config):
    self._render_slide_title(slide, slide_data.get("title", ""), config)

    faq_items = slide_data.get("faq", [])
    num_pairs = min(len(faq_items), 3)
    pair_height = self.CONTENT_HEIGHT / max(num_pairs, 1)

    for i, item in enumerate(faq_items[:3]):
        pair_top = self.CONTENT_TOP + i * pair_height

        # 質問
        self.add_text_box(
            slide,
            left=Inches(self.CONTENT_LEFT),
            top=Inches(pair_top),
            width=Inches(self.CONTENT_WIDTH),
            height=Inches(0.5),
            text=f"Q.  {item.get('question', '')}",
            font_size=Pt(14),
            font_color=config.primary_color,
            bold=True,
            font_family=config.font_family,
        )

        # 回答
        self.add_text_box(
            slide,
            left=Inches(self.CONTENT_LEFT),
            top=Inches(pair_top + 0.55),
            width=Inches(self.CONTENT_WIDTH),
            height=Inches(pair_height - 0.75),
            text=f"A.  {item.get('answer', '')}",
            font_size=Pt(12),
            font_color=config.text_primary,
            font_family=config.font_family,
        )

        # 区切り線（最終ペア以外）
        if i < num_pairs - 1:
            self.add_horizontal_line(
                slide,
                left=Inches(self.CONTENT_LEFT),
                top=Inches(pair_top + pair_height - 0.15),
                width=Inches(self.CONTENT_WIDTH),
                color="#CCCCCC",
            )
```

#### QuoteRenderer 設計

```
背景: #FFFFFF

┌──────────────────────────────────────────────┐
│  [スライドタイトル: 24pt]                      │
│  ─────────────────────────────────            │
│                                              │
│  ┃                                           │
│  ┃   "引用テキスト..."                         │
│  ┃    (16pt, italic)                         │
│  ┃                                           │
│                                              │
│                    -- 著者名                   │
│                    (12pt, right)              │
│                                              │
└──────────────────────────────────────────────┘
```

| 要素 | Left | Top | Width | Height | フォント | 色 |
|------|------|-----|-------|--------|---------|-----|
| スライドタイトル | 0.5in | 0.3in | 12.333in | 0.6in | 24pt, bold | #333333 |
| 装飾バー（縦線） | 1.0in | 1.8in | 0.06in | 2.5in | - | 背景: #3AA899 |
| 引用文 | 1.4in | 1.8in | 10.933in | 2.5in | 16pt, italic | #333333 |
| 著者名 | 1.4in | 4.5in | 10.933in | 0.5in | 12pt, right | #666666 |

- 装飾バー: `MSO_SHAPE.RECTANGLE` で幅0.06inの縦長矩形、fill=accent色
- 引用文: italic、行間1.5倍
- 著者名: "-- " プレフィックス付き、右揃え（`PP_ALIGN.RIGHT`）

```python
def render(self, slide, slide_data, config):
    self._render_slide_title(slide, slide_data.get("title", ""), config)

    quote_text = slide_data.get("quote", "")
    author = slide_data.get("author", "")

    # accent色の装飾バー（縦線）
    self.add_shape(
        slide,
        MSO_SHAPE.RECTANGLE,
        left=Inches(1.0),
        top=Inches(1.8),
        width=Inches(0.06),
        height=Inches(2.5),
        fill_color=config.accent_color,
    )

    # 引用文
    self.add_text_box(
        slide,
        left=Inches(1.4),
        top=Inches(1.8),
        width=Inches(10.933),
        height=Inches(2.5),
        text=quote_text,
        font_size=Pt(16),
        font_color=config.text_primary,
        italic=True,
        font_family=config.font_family,
    )

    # 著者名
    if author:
        self.add_text_box(
            slide,
            left=Inches(1.4),
            top=Inches(4.5),
            width=Inches(10.933),
            height=Inches(0.5),
            text=f"-- {author}",
            font_size=Pt(12),
            font_color=config.text_secondary,
            alignment=PP_ALIGN.RIGHT,
            font_family=config.font_family,
        )
```

#### ImageWithTextRenderer 設計

```
背景: #FFFFFF

┌──────────────────────────────────────────────┐
│  [スライドタイトル: 24pt]                      │
│  [キーメッセージ: 16pt]                        │
│  ─────────────────────────────────            │
│                                              │
│  ┌──────────────┐  ┌──────────────┐          │
│  │              │  │              │          │
│  │    画像      │  │  テキスト     │          │
│  │              │  │  説明文      │          │
│  │              │  │              │          │
│  └──────────────┘  └──────────────┘          │
│                                              │
└──────────────────────────────────────────────┘
```

コンテンツ領域を左右2分割（ガター0.4in）:
- 画像エリア: 幅5.967in
- テキストエリア: 幅5.967in

| 要素 | Left | Top | Width | Height | フォント | 色 |
|------|------|-----|-------|--------|---------|-----|
| スライドタイトル | 0.5in | 0.3in | 12.333in | 0.6in | 24pt, bold | #333333 |
| キーメッセージ | 0.5in | 0.95in | 12.333in | 0.45in | 16pt | #666666 |
| 画像エリア | 0.5in | 1.55in | 5.967in | 5.45in | - | - |
| テキスト説明 | 6.867in | 1.55in | 5.967in | 5.45in | 12pt | #333333 |

- 画像エリア・テキストエリアの幅: `(12.333 - 0.4) / 2 = 5.967in`
- テキストエリアのLeft: `0.5 + 5.967 + 0.4 = 6.867in`

画像配置の実装:
```python
def render(self, slide, slide_data, config):
    self._render_slide_title(slide, slide_data.get("title", ""), config)
    if slide_data.get("key_message"):
        self._render_key_message(slide, slide_data["key_message"], config)

    image_path = slide_data.get("image_path", "")
    description = slide_data.get("description", "")

    # 画像配置
    img_left = Inches(0.5)
    img_top = Inches(self.CONTENT_TOP)
    img_width = Inches(5.967)
    img_height = Inches(self.CONTENT_HEIGHT)

    if image_path and os.path.exists(image_path):
        # アスペクト比を維持してフィット
        from PIL import Image
        with Image.open(image_path) as img:
            img_w, img_h = img.size
        aspect = img_w / img_h
        target_aspect = 5.967 / self.CONTENT_HEIGHT

        if aspect > target_aspect:
            # 横長: 幅に合わせ、高さを調整
            actual_width = img_width
            actual_height = Inches(5.967 / aspect)
            actual_top = img_top + (img_height - actual_height) // 2
            actual_left = img_left
        else:
            # 縦長: 高さに合わせ、幅を調整
            actual_height = img_height
            actual_width = Inches(self.CONTENT_HEIGHT * aspect)
            actual_left = img_left + (img_width - actual_width) // 2
            actual_top = img_top

        slide.shapes.add_picture(
            image_path, actual_left, actual_top, actual_width, actual_height
        )
    else:
        # placeholder.png を使用
        placeholder_path = os.path.join(
            os.path.dirname(__file__), "..", "assets", "placeholder.png"
        )
        if os.path.exists(placeholder_path):
            slide.shapes.add_picture(
                placeholder_path, img_left, img_top, img_width, img_height
            )
        else:
            # プレースホルダー画像もない場合はグレー矩形で代替
            self.add_shape(
                slide, MSO_SHAPE.RECTANGLE,
                img_left, img_top, img_width, img_height,
                fill_color="#E0E0E0",
            )

    # テキスト説明
    self.add_text_box(
        slide,
        left=Inches(6.867),
        top=Inches(self.CONTENT_TOP),
        width=Inches(5.967),
        height=Inches(self.CONTENT_HEIGHT),
        text=description,
        font_size=Pt(12),
        font_color=config.text_primary,
        font_family=config.font_family,
    )
```

注: PILはアスペクト比計算のみに使用。python-pptxのadd_picture()だけでも配置は可能だが、正確なフィット計算のためPILを併用する。PILが利用できない場合はフォールバックとして固定サイズで配置する。

#### RENDERER_MAP追加

```python
# TICKET-005の6レイアウトに追加
RENDERER_MAP.update({
    "two_column": TwoColumnRenderer,
    "three_column": ThreeColumnRenderer,
    "four_column": FourColumnRenderer,
    "metrics": MetricsRenderer,
    "comparison_table": ComparisonTableRenderer,
    "faq": FaqRenderer,
    "quote": QuoteRenderer,
    "image_with_text": ImageWithTextRenderer,
})
```

#### BaseRendererへの追加メソッド

本チケットで以下のメソッドをBaseRendererに追加する:

```python
def add_vertical_line(self, slide, left, top, height,
                      color: str = "#CCCCCC", thickness: float = 1.0):
    """縦線を描画（細い矩形として実装）"""
    shape = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE,
        left, top, Pt(thickness), height
    )
    shape.fill.solid()
    shape.fill.fore_color.rgb = RGBColor.from_string(color.lstrip("#"))
    shape.line.fill.background()
    return shape

def _render_columns(self, slide, columns_data: list, config: DesignConfig,
                    gutter: float, header_font_size=Pt(14),
                    body_font_size=Pt(12)):
    """カラムレイアウト共通の描画処理"""
    # 上記設計の通り
```

## 受け入れ条件

- [ ] 8レイアウト（two_column, three_column, four_column, metrics, comparison_table, faq, quote, image_with_text）全てのスライドが正しく生成される
- [ ] カラム系レイアウトで各カラムのヘッダーと本文が正しい位置に配置される
- [ ] MetricsRendererで2〜4個の数値カードが均等配置される
- [ ] ComparisonTableRendererでヘッダー行が紺色背景・白文字、データ行が交互背景色
- [ ] 太字ラベル（`**text**`）がaccent色で表示される
- [ ] FaqRendererでQ&Aペアが縦並びで正しく表示される
- [ ] QuoteRendererで左側に装飾バー、引用文がitalic表示される
- [ ] ImageWithTextRendererで画像がアスペクト比を維持して配置される
- [ ] 画像なし時にplaceholder.pngまたはグレー矩形が表示される
- [ ] RENDERER_MAPに8レイアウト全てが登録されている
- [ ] 既存の6レイアウト（TICKET-005分）が壊れていない
- [ ] BaseRendererに追加したヘルパーメソッド（add_vertical_line, _render_columns）が動作する

## レビューメモ

<!-- レビュー時のコメント・指摘事項を記録 -->
