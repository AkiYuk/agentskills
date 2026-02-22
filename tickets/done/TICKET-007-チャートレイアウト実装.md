# TICKET-007: チャートレイアウト実装

| 項目 | 内容 |
|------|------|
| チケット番号 | TICKET-007 |
| 作成日 | 2026-02-22 |
| 更新日 | 2026-02-22 |
| 担当者 | Claude |
| ブランチ | feature/TICKET-007-chart-layout |

## 概要

slide_generator_pptx.pyにチャートレイアウト（横棒・縦棒・折れ線・円グラフの4種）を描画するChartRendererと、OOXML直接操作を行うShapeHelpersクラスを実装する。

## 背景・目的

チャートレイアウトはpython-pptxのChart APIとOOXML（lxml）直接操作を組み合わせた最も複雑なレイアウト。python-pptxのAPIだけでは角丸バー、破線グリッド線、マーカースタイルの細かい制御ができないため、ShapeHelpersクラスでXML要素を直接操作する。TICKET-005で構築したBaseRenderer基盤とTICKET-004のパーサーが出力するSlideContent.chart_data（ChartDataオブジェクト）を利用する。

## 要件定義

### チャートタイプ自動判定（TICKET-004のパーサー側で実施済み）

パーサーが `SlideContent.chart_data.chart_type` に以下の値を設定する（本チケットではRendererがこの値を参照するのみ）:

- `pie` : 値が`%`表記、または合計が概ね100
- `bar` : ラベルが時系列（年/月/Q1等のパターン）
- `line` : `<!-- chart_type: line -->` コメントで明示指定
- `horizontal_bar` : 上記のいずれにも該当しない場合（デフォルト）

### ChartRenderer

- [ ] BaseRendererを継承し、`render(slide, slide_content)` メソッドを実装する
- [ ] タイトル（`##`）とキーメッセージ（`###`）はBaseRendererの共通処理で配置する
- [ ] `chart_type` に応じて `_render_horizontal_bar()`, `_render_bar()`, `_render_line()`, `_render_pie()` にディスパッチする
- [ ] チャートの配置領域: 左=Inches(0.8), 上=Inches(2.0), 幅=Inches(11.5), 高さ=Inches(5.0)（BaseRendererのタイトル・キーメッセージ領域の下）
- [ ] テーブルの `**太字**` ラベルに強調色（accent: #3AA899）を適用する。太字ラベルはパーサーが `chart_data.bold_labels: list[int]` としてインデックスのリストで提供する
- [ ] ChartRendererをRENDERER_MAPに `"chart": ChartRenderer` として登録する

### 横棒グラフ（horizontal_bar）

- [ ] `XL_CHART_TYPE.BAR_CLUSTERED` を使用する
- [ ] `CategoryChartData` にカテゴリとシリーズデータを設定する
- [ ] 通常バーの色: グレー（#999999）
- [ ] 太字ラベルのバーの色: アクセント色（#3AA899）
- [ ] バーの形状: 角丸（ShapeHelpers.set_bar_shape_rounded()でXML操作）
- [ ] 値軸（横軸）: 非表示
- [ ] カテゴリ軸（縦軸）: フォントサイズ10pt、色#333333
- [ ] データラベル: バーの右端に値を表示、フォントサイズ9pt
- [ ] グリッド線: 破線、色#E5E5E5（ShapeHelpers.set_gridline_dashed()でXML操作）
- [ ] 凡例: 非表示（単一系列のため）
- [ ] プロットエリアの枠線: なし

### 縦棒グラフ（bar）

- [ ] `XL_CHART_TYPE.COLUMN_CLUSTERED` を使用する
- [ ] `CategoryChartData` にカテゴリとシリーズデータを設定する
- [ ] 通常バーの色: グレー（#999999）
- [ ] 太字ラベルのバーの色: アクセント色（#3AA899）
- [ ] バーの形状: 角丸（ShapeHelpers.set_bar_shape_rounded()で`ROUNDED_RECTANGLE`設定）
- [ ] カテゴリ軸（横軸）: フォントサイズ10pt、色#333333
- [ ] 値軸（縦軸）: フォントサイズ9pt、色#666666
- [ ] データラベル: バーの上部に値を表示、フォントサイズ9pt
- [ ] グリッド線: 値軸の主グリッド線を破線、色#E5E5E5
- [ ] 凡例: 非表示（単一系列のため）
- [ ] プロットエリアの枠線: なし

### 折れ線グラフ（line）

- [ ] `XL_CHART_TYPE.LINE_MARKERS` を使用する
- [ ] `CategoryChartData` に複数系列（2列目以降がそれぞれ1系列）を設定する
- [ ] 線の太さ: 2.5pt（Pt(2.5)）
- [ ] 系列1の色: primary（#1E3A5F）
- [ ] 系列2の色: accent（#3AA899）
- [ ] 系列3の色: chart.line_color_3（#EDB120、config.jsonから読み込み）
- [ ] 系列4以降の色: secondary（#4A6FA5）をベースに自動生成
- [ ] マーカー: 円形（XL_MARKER_STYLE.CIRCLE）、サイズ8pt、白塗り（#FFFFFF）、枠線は系列色と同色・1.5pt
- [ ] カテゴリ軸: フォントサイズ10pt、色#333333
- [ ] 値軸: フォントサイズ9pt、色#666666
- [ ] グリッド線: 値軸の主グリッド線を破線、色#E5E5E5
- [ ] 凡例: 系列が2つ以上の場合にチャート下部に表示
- [ ] マーカーの設定にはShapeHelpers.set_line_marker()を使用する

### 円グラフ（pie）-- 100%積み上げ棒として実装

- [ ] 要件定義書の仕様に従い、`pie` タイプは100%積み上げ棒グラフとして実装する
- [ ] `XL_CHART_TYPE.BAR_STACKED_100` を使用する
- [ ] 各カテゴリのデータを系列として設定する（カテゴリ=1つの固定値、系列=各データ項目）
- [ ] 系列の色: primary, accent, secondary, gray の順で割り当て
- [ ] 太字ラベルの系列はアクセント色（#3AA899）を優先割り当て
- [ ] バーの高さ: プロットエリア全体（ギャップ幅を小さくする）
- [ ] データラベル: 各セグメント内に値（%表示）、フォントサイズ10pt、白色
- [ ] カテゴリ軸: 非表示
- [ ] 値軸: 非表示
- [ ] 凡例: チャート下部に表示
- [ ] プロットエリアの枠線: なし

### ShapeHelpersクラス

ShapeHelpersはpython-pptxのAPIでは表現しきれないデザインをOOXML（lxml）直接操作で実現するヘルパークラス。全メソッドは `@staticmethod` で実装する。

- [ ] `ShapeHelpers.set_bar_shape_rounded(chart: Chart) -> None`
  - チャートXMLの `<c:barChart>` または `<c:bar3DChart>` 要素内の各 `<c:ser>` に `<c:shape val="cylinder"/>` を設定する（python-pptxには角丸バーのAPIがないため）
  - 実装: `chart.element` から `c:barChart/c:ser` を取得し、`<c:shape>` 要素を追加
  - 名前空間: `c = "http://schemas.openxmlformats.org/drawingml/2006/chart"`
  - 注: OOXMLのbar shapeで「角丸」に最も近いのは `cylinder` または直接 `<c:spPr>` に角丸の `<a:prstGeom prst="roundRect"/>` を設定する方法。後者はバー個別の `<c:spPr>` に設定する

- [ ] `ShapeHelpers.set_gridline_dashed(axis: Axis) -> None`
  - 軸のメジャーグリッド線を破線スタイルに設定する
  - 実装: `axis.major_gridlines` が存在しない場合は `axis.has_major_gridlines = True` で有効化
  - `axis.major_gridlines.format.line` から `<a:ln>` 要素にアクセスし、`<a:prstDash val="dash"/>` を設定
  - 色は `<a:solidFill><a:srgbClr val="E5E5E5"/></a:solidFill>` を `<a:ln>` 内に設定
  - 名前空間: `a = "http://schemas.openxmlformats.org/drawingml/2006/main"`

- [ ] `ShapeHelpers.set_line_marker(series: ChartSeries, color_hex: str, marker_size: int = 8) -> None`
  - 折れ線グラフの系列にマーカースタイルを設定する
  - python-pptxの `series.marker.style = XL_MARKER_STYLE.CIRCLE` と `series.marker.size = marker_size` はAPIで設定可能
  - 白塗り + 枠線色の設定は一部XML操作が必要:
    - `series.marker.format.fill.solid()` で白塗り設定
    - `series.marker.format.line.color.rgb = RGBColor.from_string(color_hex)` で枠線色
    - `series.marker.format.line.width = Pt(1.5)` で枠線幅
  - マーカーのXML要素: `<c:marker>` 内の `<c:spPr>` に `<a:solidFill>` と `<a:ln>` を設定

- [ ] `ShapeHelpers.set_bar_color(chart: Chart, series_index: int, point_index: int, color_hex: str) -> None`
  - 個別のバー（データポイント）の色を設定する
  - 実装: `chart.series[series_index].points[point_index].format.fill.solid()` でAPIアクセス可能
  - `point.format.fill.fore_color.rgb = RGBColor.from_string(color_hex)` で色を設定
  - API対応しているためXML直接操作は基本不要だが、point_indexのアクセスでXML操作が必要になる場合に備える

- [ ] `ShapeHelpers.remove_chart_border(chart: Chart) -> None`
  - チャートのプロットエリアの枠線を非表示にする
  - 実装: `chart.element` から `<c:plotArea>` の `<c:spPr>` を取得し、`<a:ln><a:noFill/></a:ln>` を設定

### 共通デザイン設定

- [ ] 全てのチャートでフォントファミリーはconfig.jsonの `font.family`（デフォルト: "Meiryo UI"）を使用する
- [ ] 色はconfig.jsonのカラーパレットから取得する:
  - primary: #1E3A5F
  - secondary: #4A6FA5
  - accent: #3AA899
  - gray: #999999
  - text.primary: #333333
  - text.secondary: #666666
  - text.light: #FFFFFF
  - chart.line_color_3: #EDB120

## 設計

### 変更対象

| ファイル | 変更内容 |
|----------|----------|
| `skills/slide-generator/scripts/slide_generator_pptx.py` | ChartRendererクラスの追加、ShapeHelpersクラスの追加、RENDERER_MAPへのchart登録 |

### 設計方針

#### クラス構成

```python
class ShapeHelpers:
    """python-pptxのAPIでは表現できないデザインをOOXML直接操作で実現するヘルパー"""

    @staticmethod
    def set_bar_shape_rounded(chart: Chart) -> None:
        """バーの形状を角丸に設定する（OOXML操作）"""
        ...

    @staticmethod
    def set_gridline_dashed(axis) -> None:
        """軸のグリッド線を破線スタイル・#E5E5E5に設定する（OOXML操作）"""
        ...

    @staticmethod
    def set_line_marker(series, color_hex: str, marker_size: int = 8) -> None:
        """折れ線グラフのマーカーを設定する（APIとOOXML併用）"""
        ...

    @staticmethod
    def set_bar_color(chart: Chart, series_index: int, point_index: int, color_hex: str) -> None:
        """個別のバーポイントの色を設定する"""
        ...

    @staticmethod
    def remove_chart_border(chart: Chart) -> None:
        """チャートのプロットエリア枠線を非表示にする（OOXML操作）"""
        ...


class ChartRenderer(BaseRenderer):
    """チャートレイアウト（4種のグラフ）のRenderer"""

    def render(self, slide, slide_content: SlideContent) -> None:
        """chart_typeに応じてサブメソッドにディスパッチする"""
        self._add_title(slide, slide_content)       # BaseRendererの共通処理
        self._add_key_message(slide, slide_content)  # BaseRendererの共通処理

        chart_type = slide_content.chart_data.chart_type
        if chart_type == "horizontal_bar":
            self._render_horizontal_bar(slide, slide_content)
        elif chart_type == "bar":
            self._render_bar(slide, slide_content)
        elif chart_type == "line":
            self._render_line(slide, slide_content)
        elif chart_type == "pie":
            self._render_pie(slide, slide_content)

    def _render_horizontal_bar(self, slide, slide_content) -> None: ...
    def _render_bar(self, slide, slide_content) -> None: ...
    def _render_line(self, slide, slide_content) -> None: ...
    def _render_pie(self, slide, slide_content) -> None: ...
```

#### 各チャート描画の処理フロー

**横棒グラフ（_render_horizontal_bar）:**
1. `CategoryChartData` を作成し、カテゴリ（ラベル列）と系列（値列）を設定
2. `slide.shapes.add_chart(XL_CHART_TYPE.BAR_CLUSTERED, ...)` でチャート追加
3. `ShapeHelpers.set_bar_shape_rounded(chart)` で角丸設定
4. 太字ラベルのデータポイントに `ShapeHelpers.set_bar_color()` でアクセント色を適用、それ以外はグレー
5. `chart.value_axis.visible = False` で値軸を非表示
6. カテゴリ軸のフォント設定（10pt、#333333）
7. `ShapeHelpers.set_gridline_dashed(chart.value_axis)` で破線グリッド
8. データラベルの有効化と設定
9. `chart.has_legend = False` で凡例非表示
10. `ShapeHelpers.remove_chart_border(chart)` で枠線非表示

**縦棒グラフ（_render_bar）:**
1. `CategoryChartData` を作成し、カテゴリと系列を設定
2. `slide.shapes.add_chart(XL_CHART_TYPE.COLUMN_CLUSTERED, ...)` でチャート追加
3. `ShapeHelpers.set_bar_shape_rounded(chart)` で角丸設定
4. 太字ラベルのデータポイントにアクセント色を適用、それ以外はグレー
5. カテゴリ軸のフォント設定（10pt、#333333）
6. 値軸のフォント設定（9pt、#666666）
7. `ShapeHelpers.set_gridline_dashed(chart.value_axis)` で破線グリッド
8. データラベルの有効化と設定（バー上部）
9. `chart.has_legend = False` で凡例非表示
10. `ShapeHelpers.remove_chart_border(chart)` で枠線非表示

**折れ線グラフ（_render_line）:**
1. `CategoryChartData` を作成し、カテゴリと複数系列を設定
2. `slide.shapes.add_chart(XL_CHART_TYPE.LINE_MARKERS, ...)` でチャート追加
3. 各系列にループ処理:
   - 系列色の決定（1→primary、2→accent、3→line_color_3、4+→secondary系）
   - `series.format.line.width = Pt(2.5)` で線幅設定
   - `series.format.line.color.rgb = RGBColor(...)` で線色設定
   - `ShapeHelpers.set_line_marker(series, color_hex, marker_size=8)` でマーカー設定
4. カテゴリ軸のフォント設定（10pt、#333333）
5. 値軸のフォント設定（9pt、#666666）
6. `ShapeHelpers.set_gridline_dashed(chart.value_axis)` で破線グリッド
7. 系列が2以上の場合 `chart.has_legend = True` + 凡例位置設定
8. `ShapeHelpers.remove_chart_border(chart)` で枠線非表示

**円グラフ -- 100%積み上げ棒（_render_pie）:**
1. `CategoryChartData` を作成、カテゴリ=固定1項目、系列=各データ項目
2. `slide.shapes.add_chart(XL_CHART_TYPE.BAR_STACKED_100, ...)` でチャート追加
3. 各系列に色を割り当て（primary, accent, secondary, gray ...）
4. 太字ラベルの系列にはアクセント色を優先適用
5. ギャップ幅の縮小（`chart.plots[0].gap_width = 50` 等）
6. カテゴリ軸・値軸を非表示
7. データラベル: 各系列のポイントに%値を表示（白色、10pt）
8. `chart.has_legend = True` + 凡例位置設定（下部）
9. `ShapeHelpers.remove_chart_border(chart)` で枠線非表示

#### OOXML名前空間

ShapeHelpersで使用するXML名前空間:

```python
NSMAP = {
    'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
    'c': 'http://schemas.openxmlformats.org/drawingml/2006/chart',
    'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
}
```

#### OOXML操作の具体的な要素

- **角丸バー**: `<c:barChart>/<c:ser>/<c:spPr>/<a:prstGeom prst="roundRect"/>` を各 `<c:ser>` に挿入
- **破線グリッド**: `<c:valAx>/<c:majorGridlines>/<c:spPr>/<a:ln>/<a:prstDash val="dash"/>` と `<a:ln>/<a:solidFill>/<a:srgbClr val="E5E5E5"/>`
- **マーカー**: `<c:ser>/<c:marker>/<c:spPr>/<a:solidFill>/<a:srgbClr val="FFFFFF"/>` と `<a:ln w="19050">/<a:solidFill>/<a:srgbClr val="{color}"/>`（w=19050は1.5pt相当）
- **枠線非表示**: `<c:plotArea>/<c:spPr>/<a:ln>/<a:noFill/>`
- **100%積み上げ棒のギャップ**: `<c:barChart>/<c:gapWidth val="50"/>`

#### python-pptxの主要API使用箇所

```python
from pptx.chart.data import CategoryChartData
from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION, XL_LABEL_POSITION, XL_MARKER_STYLE
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor

# チャート追加
chart_data = CategoryChartData()
chart_data.categories = ['2021', '2022', '2023']
chart_data.add_series('売上', (100, 200, 300))
chart_frame = slide.shapes.add_chart(
    XL_CHART_TYPE.COLUMN_CLUSTERED,
    Inches(0.8), Inches(2.0), Inches(11.5), Inches(5.0),
    chart_data
)
chart = chart_frame.chart

# データポイントの色設定
point = chart.series[0].points[idx]
point.format.fill.solid()
point.format.fill.fore_color.rgb = RGBColor(0x3A, 0xA8, 0x99)

# 折れ線の設定
series.format.line.width = Pt(2.5)
series.format.line.color.rgb = RGBColor(0x1E, 0x3A, 0x5F)

# データラベル
plot = chart.plots[0]
plot.has_data_labels = True
data_labels = plot.data_labels
data_labels.font.size = Pt(9)
data_labels.number_format = '0'
```

## 受け入れ条件

- [ ] 横棒グラフが正しく生成される（角丸バー、通常色=グレー、太字ラベル=アクセント色、破線グリッド、値軸非表示）
- [ ] 縦棒グラフが正しく生成される（角丸バー、通常色=グレー、太字ラベル=アクセント色、破線グリッド、データラベル上部表示）
- [ ] 折れ線グラフが正しく生成される（2.5pt線、円形マーカー8pt白枠付き、複数系列の色分け、3系列目=#EDB120）
- [ ] 100%積み上げ棒グラフ（pieタイプ）が正しく生成される（系列色割り当て、%データラベル、凡例表示）
- [ ] ShapeHelpersの5メソッドが正しく動作し、XML操作がPPTXの破損を引き起こさない
- [ ] テーブルの太字ラベルに強調色（#3AA899）が適用される
- [ ] config.jsonのカラーパレット・フォント設定が正しく反映される
- [ ] 既存の14レイアウト（TICKET-005, TICKET-006）が壊れていない
- [ ] PRが作成され、mainにマージされている

## レビューメモ

<!-- レビュー時のコメント・指摘事項を記録 -->
