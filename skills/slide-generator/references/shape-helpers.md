# 図形描画ヘルパー仕様

本ドキュメントはpython-pptxを使用したスライド描画の技術仕様を定義する。
slide_generator_pptx.py が各レイアウトを描画する際のルールを記載する。

## スライド基本設定

- **スライドサイズ**: 16:9（幅: 13333400 EMU = 33.867cm、高さ: 7500000 EMU = 19.05cm）
- **単位**: python-pptxではEMU（English Metric Unit）を使用。`pptx.util.Inches`, `pptx.util.Pt`, `pptx.util.Emu` で変換

## フォント設定

| 用途 | サイズ | 色 |
|------|-------|-----|
| タイトル（title/sectionレイアウト） | 44pt | `#FFFFFF`（背景が暗い場合）/ `#333333` |
| セクション見出し | 36pt | `#333333` |
| スライドタイトル | 24pt | `#333333` |
| キーメッセージ | 16pt | `#666666` |
| 本文 | 12pt | `#333333` |

- フォントファミリーは config.json の `font.family`（デフォルト: Meiryo UI）を使用する
- 行間は1.5倍を基本とする

## カラーパレット

config.json の `palette` から色を取得する。

```python
from pptx.util import Pt
from pptx.dml.color import RGBColor

def hex_to_rgb(hex_color: str) -> RGBColor:
    """16進数カラーコードをRGBColorに変換する"""
    hex_color = hex_color.lstrip('#')
    r = int(hex_color[0:2], 16)
    g = int(hex_color[2:4], 16)
    b = int(hex_color[4:6], 16)
    return RGBColor(r, g, b)
```

## 角丸矩形

カード、ボタン、メトリクスボックス等に使用する。

```python
from pptx.enum.shapes import MSO_SHAPE

shape = slide.shapes.add_shape(
    MSO_SHAPE.ROUNDED_RECTANGLE,
    left, top, width, height
)
# 角丸の半径調整
shape.adjustments[0] = 0.05  # 角丸の度合い（0-1）
```

## テキストフレーム設定

```python
tf = shape.text_frame
tf.word_wrap = True
tf.auto_size = None  # 自動サイズ調整を無効化

p = tf.paragraphs[0]
p.font.name = config["font"]["family"]
p.font.size = Pt(12)
p.font.color.rgb = hex_to_rgb(config["palette"]["text"]["primary"])
p.line_spacing = Pt(18)  # 行間1.5倍（12pt x 1.5）
p.alignment = PP_ALIGN.LEFT  # 左揃え
```

配置オプション:
- `PP_ALIGN.LEFT`: 左揃え（本文デフォルト）
- `PP_ALIGN.CENTER`: 中央揃え（タイトル、メトリクス）
- `PP_ALIGN.RIGHT`: 右揃え（著者名等）

## グリッド線描画

ガイド線や区切り線に使用する。

```python
from pptx.util import Inches, Emu

line = slide.shapes.add_connector(
    MSO_CONNECTOR.STRAIGHT, x1, y1, x2, y2
)
line.line.color.rgb = RGBColor(0xE5, 0xE5, 0xE5)  # #E5E5E5
line.line.dash_style = MSO_LINE_DASH_STYLE.DASH    # 破線
line.line.width = Pt(0.5)
```

## グラフ描画仕様

### 共通ルール

- 通常色: `#999999`（gray）
- 強調色: `#3AA899`（accent）
- テーブルのラベルを `**太字**` にすると強調色で表示
- 軸ラベルのフォントサイズ: 10pt

### 棒グラフ共通

- 棒は角丸矩形（ROUNDED_RECTANGLE）で描画する
- 通常の棒: グレー `#999999`
- 強調の棒: アクセント色 `#3AA899`

### 縦棒グラフ (bar)

- 時系列ラベル（年/月/Q1等）のデータに使用する
- X軸: ラベル（時系列）
- Y軸: 値
- 棒は下から上に伸びる

### 横棒グラフ (horizontal_bar)

- デフォルトのグラフタイプ
- X軸: 値
- Y軸: ラベル
- 棒は左から右に伸びる
- カテゴリの比較に適する

### 折れ線グラフ (line)

- 線幅: 2.5pt
- マーカー: 円形、サイズ8pt、白枠付き
- 複数系列に対応:
  - 系列1: primary色 `#1E3A5F`
  - 系列2: accent色 `#3AA899`
  - 系列3: chart.line_color_3 `#EDB120`
- `<!-- chart_type: line -->` コメントで明示的に指定する

### 100%積み上げ棒グラフ (pie)

- 値が%表記または合計が約100の場合に使用する
- 横棒の積み上げで構成比を表現する
- 各セグメントに割合ラベルを表示する

## テンプレート対応ルール

### 優先順位

テンプレートPPTXの指定は以下の順序で決定する:

1. CLIの `--template` 引数
2. config.json の `template.pptx_path`
3. 指定なし（白背景で自動生成）

### テンプレート使用時

- スライドマスタ・レイアウトを保持する
- 既存スライドは全て削除する
- プレースホルダは除去する
- 背景色はテンプレートのものを尊重する

### テンプレート未使用時

- 白紙プレゼンを新規作成する
- 背景色はコードで設定する:
  - title: 紺色背景 `#1E3A5F`
  - section: 薄グレー `#F5F5F5`
  - cta: 薄グレー `#F5F5F5`
  - その他: 白 `#FFFFFF`
