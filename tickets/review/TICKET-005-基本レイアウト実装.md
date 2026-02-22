# TICKET-005: 基本レイアウト実装

| 項目 | 内容 |
|------|------|
| チケット番号 | TICKET-005 |
| 作成日 | 2026-02-22 |
| 更新日 | 2026-02-22 |
| 担当者 | Claude |
| ブランチ | feature/TICKET-005-basic-layouts |

## 概要

slide_generator_pptx.pyの基盤（DesignConfig, BaseRenderer, CLI引数処理, テンプレート処理）と基本6レイアウト（title, section, toc, bullet_points, numbered_list, cta）のRendererを実装する。

## 背景・目的

PPTX生成スクリプトの土台を構築し、最もシンプルな6レイアウトで動作を確認する。BaseRendererに共通処理を集約することで、以降のチケット（TICKET-006〜007）で追加するレイアウトが継承するだけで一貫したデザインを保てるようにする。

## 要件定義

### 共通基盤

- [ ] **DesignConfig クラス**: config.jsonを読み込み、カラーパレット・フォント・サイズ設定をプロパティとして提供する
  - palette: primary, secondary, accent, gray, text(primary/secondary/light), background(primary/secondary/dark), chart(line_color_3)
  - font: family（デフォルト "Meiryo UI"）
  - font_size: title=44pt, section=36pt, slide_title=24pt, key_message=16pt, body=12pt
  - slide_size: width=13.333in, height=7.5in（16:9）
- [ ] **BaseRenderer クラス**: 全Rendererの基底クラス。以下の共通処理を提供
  - `render(slide, slide_data, config)`: サブクラスで実装する抽象メソッド
  - `add_text_box(slide, left, top, width, height, text, font_size, font_color, bold, italic, alignment, font_family)`: テキストボックスを追加するヘルパー
  - `set_background_color(slide, hex_color)`: スライド背景色を設定するヘルパー
  - `set_background_gradient(slide, color1, color2)`: グラデーション背景のヘルパー（将来用、本チケットでは未使用でもよい）
  - `add_shape(slide, shape_type, left, top, width, height, fill_color, line_color)`: 図形追加ヘルパー
  - `add_horizontal_line(slide, left, top, width, color, thickness)`: 水平線ヘルパー
  - `set_font(text_frame, font_family, font_size, font_color, bold, italic, alignment, line_spacing)`: テキストフレームのフォント設定ヘルパー
- [ ] **RENDERER_MAP**: レイアウト名（文字列）→ Rendererクラスの辞書マッピング
  - 本チケットでは6レイアウト分を登録
  - TICKET-006, 007で追加される際はこのMAPに追記する形式
- [ ] **CLI引数処理**: argparseによる以下の引数
  - `--markdown-file`（必須）: slides.mdファイルのパス
  - `--config`（任意）: config.jsonファイルのパス（デフォルト: `skills/slide-generator/config.json`）
  - `--title`（任意）: 出力ファイル名のプレフィックス
  - `--output-dir`（任意）: 出力ディレクトリ（デフォルト: config.jsonのoutput.dir）
  - `--template`（任意）: テンプレートPPTXファイルのパス
  - `--layout`（任意）: 特定レイアウトのみを生成（デバッグ用）
- [ ] **テンプレート処理**:
  - テンプレート優先順位: CLI `--template` > config.json `template.pptx_path` > なし（白紙）
  - テンプレート使用時: `pptx.Presentation(template_path)` で読み込み、スライドマスタ・レイアウトを保持し、既存スライドを全削除、各スライドのプレースホルダーを除去
  - テンプレート未使用時: `pptx.Presentation()` で白紙プレゼンを新規作成、背景色はコードで設定
- [ ] **スライドサイズ**: 16:9（幅13.333in × 高さ7.5in）を `prs.slide_width` / `prs.slide_height` で設定
- [ ] **行間**: 全テキストボックスのデフォルト行間は1.5倍（`paragraph.line_spacing = 1.5`）
- [ ] **フォント**: 全テキストに Meiryo UI を適用（`run.font.name = "Meiryo UI"`）
- [ ] **slides_markdown.py連携**: `SlidesMarkdownParser` をimportし、`parse()` の返却する `Presentation` オブジェクトからスライドデータを取得

### 6レイアウトRenderer

- [ ] **TitleRenderer**: 表紙レイアウト
  - 背景: 紺色（`#1E3A5F` = palette.background.dark）
  - メインタイトル: 44pt, bold, 白色, 中央揃え
  - サブタイトル: 16pt, 白色, 中央揃え
  - 文字数制限: タイトル30文字、サブタイトル50文字
- [ ] **SectionRenderer**: セクション区切りレイアウト
  - 背景: 薄グレー（`#F5F5F5` = palette.background.secondary）
  - セクションタイトル: 36pt, bold, テキスト色primary（`#333333`）, 中央揃え
  - 文字数制限: タイトル20文字
- [ ] **TocRenderer**: 目次・サマリーレイアウト
  - 背景: 白色（`#FFFFFF`）
  - スライドタイトル: 24pt
  - キーメッセージ: 16pt
  - 目次項目: 12pt、箇条書き形式、番号付き（1. 2. 3. ...）
  - 推奨項目数: 3〜7項目
  - 文字数制限: タイトル25文字、キーメッセージ40文字、1項目20文字
- [ ] **BulletPointsRenderer**: 箇条書きレイアウト（デフォルト）
  - 背景: 白色
  - スライドタイトル: 24pt
  - キーメッセージ: 16pt
  - 箇条書き項目: 12pt、ビュレットマーカー（`●`）付き
  - 推奨項目数: 3〜5項目
  - 文字数制限: タイトル25文字、キーメッセージ40文字、1項目50文字
- [ ] **NumberedListRenderer**: 手順・ステップレイアウト
  - 背景: 白色
  - スライドタイトル: 24pt
  - キーメッセージ: 16pt
  - 手順項目: ステップ番号（accent色の丸数字または番号）+ タイトル（14pt bold）+ 説明（12pt）
  - 推奨項目数: 3〜5項目
  - 文字数制限: タイトル25文字、キーメッセージ40文字、項目タイトル15文字、説明50文字
- [ ] **CtaRenderer**: 行動喚起レイアウト
  - 背景: 薄グレー（`#F5F5F5`）
  - メッセージ: 24pt, bold, 中央揃え
  - 説明: 16pt, テキスト色secondary, 中央揃え
  - ボタン風テキスト: accent色（`#3AA899`）背景の角丸矩形内にテキスト（14pt, 白色, bold）
  - 文字数制限: メッセージ20文字、説明40文字、ボタン15文字

## 設計

### 変更対象

| ファイル | 変更内容 |
|----------|----------|
| `skills/slide-generator/scripts/slide_generator_pptx.py` | 新規実装。DesignConfig, BaseRenderer, 6レイアウトRenderer, RENDERER_MAP, CLI, main() |

### 設計方針

#### 全体アーキテクチャ

```
slide_generator_pptx.py
├── DesignConfig           # config.json読み込み・デザイン設定管理
├── BaseRenderer           # 共通描画処理の基底クラス
├── TitleRenderer          # layout: title
├── SectionRenderer        # layout: section
├── TocRenderer            # layout: toc
├── BulletPointsRenderer   # layout: bullet_points
├── NumberedListRenderer   # layout: numbered_list
├── CtaRenderer            # layout: cta
├── RENDERER_MAP           # {"title": TitleRenderer, ...}
├── generate_presentation  # メインの生成関数
└── main                   # CLI引数処理・エントリーポイント
```

#### スライド座標系（共通）

スライドサイズ: 13.333in × 7.5in（16:9）。余白を統一してレイアウトの一貫性を保つ。

| 領域 | Left | Top | Width | Height | 用途 |
|------|------|-----|-------|--------|------|
| スライドタイトル | 0.5in | 0.3in | 12.333in | 0.6in | 24pt タイトル |
| キーメッセージ | 0.5in | 0.95in | 12.333in | 0.45in | 16pt キーメッセージ |
| コンテンツ領域 | 0.5in | 1.55in | 12.333in | 5.45in | レイアウト固有の本文領域 |
| 左右余白 | 0.5in | - | - | - | 両側 |
| 下余白 | - | - | - | 0.5in | 下部 |

#### DesignConfig 設計

```python
class DesignConfig:
    """config.jsonからデザイン設定を読み込むクラス"""

    def __init__(self, config_path: str = None):
        # config_pathがNoneの場合はデフォルト値を使用
        # JSONを読み込み、self._config に格納

    @property
    def font_family(self) -> str:
        return self._config.get("font", {}).get("family", "Meiryo UI")

    @property
    def palette(self) -> dict:
        return self._config.get("palette", {})

    # 以下、よく使う色をショートカットプロパティで提供
    @property
    def primary_color(self) -> str:      # "#1E3A5F"
    @property
    def secondary_color(self) -> str:    # "#4A6FA5"
    @property
    def accent_color(self) -> str:       # "#3AA899"
    @property
    def text_primary(self) -> str:       # "#333333"
    @property
    def text_secondary(self) -> str:     # "#666666"
    @property
    def text_light(self) -> str:         # "#FFFFFF"
    @property
    def bg_primary(self) -> str:         # "#FFFFFF"
    @property
    def bg_secondary(self) -> str:       # "#F5F5F5"
    @property
    def bg_dark(self) -> str:            # "#1E3A5F"
```

#### BaseRenderer 設計

```python
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.enum.shapes import MSO_SHAPE

class BaseRenderer:
    """全Rendererの基底クラス"""

    # 共通座標定数（Inches）
    MARGIN_LEFT = 0.5
    MARGIN_RIGHT = 0.5
    MARGIN_BOTTOM = 0.5
    CONTENT_LEFT = 0.5
    CONTENT_WIDTH = 12.333  # 13.333 - 0.5 - 0.5

    TITLE_TOP = 0.3
    TITLE_HEIGHT = 0.6
    KEY_MSG_TOP = 0.95
    KEY_MSG_HEIGHT = 0.45
    CONTENT_TOP = 1.55
    CONTENT_HEIGHT = 5.45   # 7.5 - 1.55 - 0.5

    def render(self, slide, slide_data: dict, config: DesignConfig):
        """サブクラスで実装する描画メソッド"""
        raise NotImplementedError

    def _render_slide_title(self, slide, text: str, config: DesignConfig,
                            font_color: str = None):
        """共通: スライドタイトルの描画（24pt）"""
        color = font_color or config.text_primary
        self.add_text_box(
            slide,
            left=Inches(self.CONTENT_LEFT),
            top=Inches(self.TITLE_TOP),
            width=Inches(self.CONTENT_WIDTH),
            height=Inches(self.TITLE_HEIGHT),
            text=text,
            font_size=Pt(24),
            font_color=color,
            bold=True,
            font_family=config.font_family,
        )

    def _render_key_message(self, slide, text: str, config: DesignConfig,
                            font_color: str = None):
        """共通: キーメッセージの描画（16pt）"""
        color = font_color or config.text_secondary
        self.add_text_box(
            slide,
            left=Inches(self.CONTENT_LEFT),
            top=Inches(self.KEY_MSG_TOP),
            width=Inches(self.CONTENT_WIDTH),
            height=Inches(self.KEY_MSG_HEIGHT),
            text=text,
            font_size=Pt(16),
            font_color=color,
            font_family=config.font_family,
        )

    def add_text_box(self, slide, left, top, width, height,
                     text: str, font_size=Pt(12),
                     font_color: str = "#333333",
                     bold: bool = False, italic: bool = False,
                     alignment=PP_ALIGN.LEFT,
                     font_family: str = "Meiryo UI",
                     line_spacing: float = 1.5):
        """テキストボックスを追加し、フォント設定を適用"""
        txBox = slide.shapes.add_textbox(left, top, width, height)
        tf = txBox.text_frame
        tf.word_wrap = True
        p = tf.paragraphs[0]
        p.text = text
        p.font.size = font_size
        p.font.color.rgb = RGBColor.from_string(font_color.lstrip("#"))
        p.font.bold = bold
        p.font.italic = italic
        p.font.name = font_family
        p.alignment = alignment
        p.line_spacing = line_spacing
        return txBox

    def set_background_color(self, slide, hex_color: str):
        """スライドの背景色を単色で設定"""
        background = slide.background
        fill = background.fill
        fill.solid()
        fill.fore_color.rgb = RGBColor.from_string(hex_color.lstrip("#"))

    def add_shape(self, slide, shape_type, left, top, width, height,
                  fill_color: str = None, line_color: str = None):
        """図形を追加"""
        shape = slide.shapes.add_shape(shape_type, left, top, width, height)
        if fill_color:
            shape.fill.solid()
            shape.fill.fore_color.rgb = RGBColor.from_string(fill_color.lstrip("#"))
        if line_color:
            shape.line.color.rgb = RGBColor.from_string(line_color.lstrip("#"))
        else:
            shape.line.fill.background()  # 枠線なし
        return shape

    def add_horizontal_line(self, slide, left, top, width,
                            color: str = "#CCCCCC", thickness: float = 1.0):
        """水平線を描画（細い矩形として実装）"""
        shape = slide.shapes.add_shape(
            MSO_SHAPE.RECTANGLE,
            left, top, width, Pt(thickness)
        )
        shape.fill.solid()
        shape.fill.fore_color.rgb = RGBColor.from_string(color.lstrip("#"))
        shape.line.fill.background()
        return shape
```

#### TitleRenderer 設計

```
背景: #1E3A5F（紺色 = palette.background.dark）

┌──────────────────────────────────────────────┐
│                                              │
│                                              │
│         [メインタイトル: 44pt, 白, bold]       │
│            中央寄せ、上下中央付近              │
│                                              │
│         [サブタイトル: 16pt, 白]              │
│            中央寄せ、タイトル下               │
│                                              │
│                                              │
└──────────────────────────────────────────────┘
```

| 要素 | Left | Top | Width | Height | フォント | 色 |
|------|------|-----|-------|--------|---------|-----|
| メインタイトル | 1.5in | 2.2in | 10.333in | 1.2in | 44pt, bold | #FFFFFF |
| サブタイトル | 1.5in | 3.6in | 10.333in | 0.8in | 16pt | #FFFFFF |

- 共通のスライドタイトル・キーメッセージ描画は使用しない（表紙は独自配置）
- `set_background_color(slide, config.bg_dark)` で背景設定
- タイトル: `PP_ALIGN.CENTER`
- サブタイトル: `PP_ALIGN.CENTER`

#### SectionRenderer 設計

```
背景: #F5F5F5（薄グレー = palette.background.secondary）

┌──────────────────────────────────────────────┐
│                                              │
│                                              │
│                                              │
│       [セクションタイトル: 36pt, bold]         │
│            中央寄せ、上下中央配置              │
│                                              │
│                                              │
│                                              │
└──────────────────────────────────────────────┘
```

| 要素 | Left | Top | Width | Height | フォント | 色 |
|------|------|-----|-------|--------|---------|-----|
| セクションタイトル | 1.0in | 2.5in | 11.333in | 1.5in | 36pt, bold | #333333 |

- `set_background_color(slide, config.bg_secondary)` で背景設定
- `PP_ALIGN.CENTER`
- タイトルの下にアクセント色の水平線を配置（装飾）
  - 水平線: left=5.667in, top=4.2in, width=2.0in, color=`#3AA899`, thickness=3pt

#### TocRenderer 設計

```
背景: #FFFFFF

┌──────────────────────────────────────────────┐
│  [スライドタイトル: 24pt]                      │
│  [キーメッセージ: 16pt]                        │
│  ─────────────────────────────────            │
│                                              │
│  1. 項目テキスト                              │
│  2. 項目テキスト                              │
│  3. 項目テキスト                              │
│  ...                                         │
│                                              │
└──────────────────────────────────────────────┘
```

| 要素 | Left | Top | Width | Height | フォント | 色 |
|------|------|-----|-------|--------|---------|-----|
| スライドタイトル | 0.5in | 0.3in | 12.333in | 0.6in | 24pt, bold | #333333 |
| キーメッセージ | 0.5in | 0.95in | 12.333in | 0.45in | 16pt | #666666 |
| 区切り線 | 0.5in | 1.45in | 12.333in | - | - | #CCCCCC, 1pt |
| 目次項目（各行） | 0.8in | 1.55in〜 | 11.733in | 各0.55in | 14pt | #333333 |

- `_render_slide_title()` と `_render_key_message()` を使用
- 目次項目は番号付き: `"1.  {item_text}"` のように番号+テキスト
- 番号部分は `accent_color`（`#3AA899`）で装飾する。ただし python-pptx では同一段落内の部分的色分けが可能（Run単位）:
  - 1つのテキストボックスに複数段落（paragraph）で各項目を配置
  - 各段落に2つのRun: 番号Run（accent色, bold）+ テキストRun（text_primary色）
- 項目間隔: 行間1.5倍で自動調整

#### BulletPointsRenderer 設計

```
背景: #FFFFFF

┌──────────────────────────────────────────────┐
│  [スライドタイトル: 24pt]                      │
│  [キーメッセージ: 16pt]                        │
│  ─────────────────────────────────            │
│                                              │
│  ● 箇条書き項目1                              │
│  ● 箇条書き項目2                              │
│  ● 箇条書き項目3                              │
│  ...                                         │
│                                              │
└──────────────────────────────────────────────┘
```

| 要素 | Left | Top | Width | Height | フォント | 色 |
|------|------|-----|-------|--------|---------|-----|
| スライドタイトル | 0.5in | 0.3in | 12.333in | 0.6in | 24pt, bold | #333333 |
| キーメッセージ | 0.5in | 0.95in | 12.333in | 0.45in | 16pt | #666666 |
| 区切り線 | 0.5in | 1.45in | 12.333in | - | - | #CCCCCC, 1pt |
| 箇条書き領域 | 0.5in | 1.55in | 12.333in | 5.45in | 12pt | #333333 |

- `_render_slide_title()` と `_render_key_message()` を使用
- 箇条書きは1つのテキストボックス内に複数段落で配置
- 各段落のテキスト: `"●  {item_text}"`
  - ビュレットマーカーは `accent_color`（`#3AA899`）のRun、テキストは `text_primary` のRun
- `PP_ALIGN.LEFT`
- 行間1.5倍

#### NumberedListRenderer 設計

```
背景: #FFFFFF

┌──────────────────────────────────────────────┐
│  [スライドタイトル: 24pt]                      │
│  [キーメッセージ: 16pt]                        │
│  ─────────────────────────────────            │
│                                              │
│  ① ステップタイトル1                          │
│     説明テキスト...                            │
│                                              │
│  ② ステップタイトル2                          │
│     説明テキスト...                            │
│                                              │
│  ③ ステップタイトル3                          │
│     説明テキスト...                            │
│                                              │
└──────────────────────────────────────────────┘
```

| 要素 | Left | Top | Width | Height | フォント | 色 |
|------|------|-----|-------|--------|---------|-----|
| スライドタイトル | 0.5in | 0.3in | 12.333in | 0.6in | 24pt, bold | #333333 |
| キーメッセージ | 0.5in | 0.95in | 12.333in | 0.45in | 16pt | #666666 |
| 区切り線 | 0.5in | 1.45in | 12.333in | - | - | #CCCCCC, 1pt |
| ステップ番号円 | 0.5in | 各ステップのTop | 0.45in | 0.45in | - | 背景: #3AA899 |
| 番号テキスト | 0.5in | 各ステップのTop | 0.45in | 0.45in | 14pt, bold, center | #FFFFFF |
| ステップタイトル | 1.1in | 各ステップのTop | 11.733in | 0.35in | 14pt, bold | #333333 |
| ステップ説明 | 1.1in | タイトル下 | 11.733in | 0.4in | 12pt | #666666 |

- 各ステップは番号（accent色の円）+ タイトル + 説明の3要素で構成
- ステップ番号の円: `MSO_SHAPE.OVAL` で描画、fill=accent色、内部に白文字で番号
- ステップの間隔: 項目数に応じてコンテンツ領域（5.45in）を等分
  - 3項目: 約1.82in間隔
  - 4項目: 約1.36in間隔
  - 5項目: 約1.09in間隔
- ステップの開始Top: `CONTENT_TOP + (index * step_interval)`

#### CtaRenderer 設計

```
背景: #F5F5F5（薄グレー）

┌──────────────────────────────────────────────┐
│                                              │
│                                              │
│        [メッセージ: 24pt, bold, 中央]          │
│                                              │
│        [説明: 16pt, secondary, 中央]          │
│                                              │
│        ┌─────────────────────┐               │
│        │  ボタンテキスト       │               │
│        └─────────────────────┘               │
│                                              │
└──────────────────────────────────────────────┘
```

| 要素 | Left | Top | Width | Height | フォント | 色 |
|------|------|-----|-------|--------|---------|-----|
| メッセージ | 1.5in | 1.8in | 10.333in | 0.8in | 24pt, bold, center | #333333 |
| 説明 | 1.5in | 2.8in | 10.333in | 0.8in | 16pt, center | #666666 |
| ボタン背景 | 4.667in | 4.2in | 4.0in | 0.7in | - | 背景: #3AA899, 角丸 |
| ボタンテキスト | 4.667in | 4.2in | 4.0in | 0.7in | 14pt, bold, center | #FFFFFF |

- `set_background_color(slide, config.bg_secondary)` で背景設定
- ボタンの角丸矩形: `MSO_SHAPE.ROUNDED_RECTANGLE` で描画
  - fill = `accent_color`（`#3AA899`）
  - テキストは図形内テキストフレームに設定（`shape.text_frame`）
  - 角丸半径の調整: `shape.adjustments[0] = 0.15`（python-pptx で角丸率を設定）
- ボタンのX位置: `(13.333 - 4.0) / 2 = 4.667in`（中央配置）

#### テンプレート処理の実装方針

```python
def create_presentation(config: DesignConfig, template_path: str = None) -> Presentation:
    """テンプレートの有無に応じてPresentationオブジェクトを生成"""
    if template_path:
        prs = Presentation(template_path)
        # 既存スライドを全削除（逆順で削除）
        for i in range(len(prs.slides) - 1, -1, -1):
            rId = prs.slides._sldIdLst[i].rId
            prs.part.drop_rel(rId)
            del prs.slides._sldIdLst[i]
        # 各スライド追加時にプレースホルダーを除去
    else:
        prs = Presentation()

    # スライドサイズを16:9に設定
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    return prs

def add_blank_slide(prs: Presentation) -> Slide:
    """空白スライドを追加し、プレースホルダーを除去"""
    slide_layout = prs.slide_layouts[6]  # 空白レイアウト（インデックス6）
    slide = prs.slides.add_slide(slide_layout)
    # プレースホルダーを除去
    for ph in list(slide.placeholders):
        sp = ph._element
        sp.getparent().remove(sp)
    return slide
```

#### メイン処理フロー

```python
def generate_presentation(markdown_path: str, config: DesignConfig,
                          template_path: str = None, output_dir: str = "output",
                          title: str = None, layout_filter: str = None):
    """slides.mdからPPTXを生成するメイン関数"""
    # 1. slides.mdをパース
    parser = SlidesMarkdownParser()
    presentation_data = parser.parse(markdown_path)

    # 2. Presentationオブジェクト生成
    prs = create_presentation(config, template_path)

    # 3. 各スライドを描画
    for slide_data in presentation_data.slides:
        layout = slide_data.get("layout", "bullet_points")
        if layout_filter and layout != layout_filter:
            continue
        renderer = RENDERER_MAP.get(layout)
        if renderer is None:
            print(f"警告: 未対応レイアウト '{layout}' をスキップします")
            continue
        slide = add_blank_slide(prs)
        renderer().render(slide, slide_data, config)

    # 4. 保存
    output_path = os.path.join(output_dir, f"{title or 'presentation'}.pptx")
    os.makedirs(output_dir, exist_ok=True)
    prs.save(output_path)
    return output_path
```

#### python-pptx API使用方針

| 操作 | API |
|------|-----|
| テキストボックス追加 | `slide.shapes.add_textbox(left, top, width, height)` |
| 図形追加 | `slide.shapes.add_shape(MSO_SHAPE.xxx, left, top, width, height)` |
| 背景色設定 | `slide.background.fill.solid()` → `fill.fore_color.rgb = RGBColor(...)` |
| フォント設定 | `run.font.name`, `run.font.size`, `run.font.color.rgb`, `run.font.bold` |
| テキスト揃え | `paragraph.alignment = PP_ALIGN.CENTER` |
| 行間設定 | `paragraph.line_spacing = 1.5` |
| 色指定 | `RGBColor.from_string("1E3A5F")` または `RGBColor(0x1E, 0x3A, 0x5F)` |
| 単位変換 | `Inches(1.0)`, `Pt(12)`, `Emu(914400)` |
| 複数Run | `paragraph.add_run()` で段落内に複数のスタイル付きテキスト |
| ワードラップ | `text_frame.word_wrap = True` |
| 角丸調整 | `shape.adjustments[0] = 0.15` |

## 受け入れ条件

- [ ] DesignConfigがconfig.jsonを正しく読み込み、全プロパティが取得できる
- [ ] BaseRendererの共通メソッド（add_text_box, set_background_color, add_shape等）が動作する
- [ ] 6レイアウト（title, section, toc, bullet_points, numbered_list, cta）のスライドが正しく生成される
- [ ] 各レイアウトのフォントサイズ・色・配置が設計仕様通りである
- [ ] config.jsonの色・フォント設定が反映される
- [ ] CLI引数（--markdown-file, --config, --title, --output-dir, --template, --layout）が全て動作する
- [ ] テンプレートあり/なし両方でスライドが正しく生成される
- [ ] スライドサイズが16:9（13.333in × 7.5in）である
- [ ] 行間が1.5倍に設定されている
- [ ] RENDERER_MAPに6レイアウトが登録されている

## レビューメモ

<!-- レビュー時のコメント・指摘事項を記録 -->
