"""slides.md パーサー・シリアライザー

slides.md形式のMarkdownファイルを解析し、スライドデータ構造に変換する。
また、スライドデータ構造からslides.md形式に出力する機能を提供する。

使い方:
    # パース（slides.md → JSON）
    python slides_markdown.py --input output/slides.md

    # バリデーション（文字数チェック）
    python slides_markdown.py --validate output/slides.md
"""

from __future__ import annotations

import argparse
import json
import logging
import re
import sys
import unicodedata
from dataclasses import asdict, dataclass, field
from typing import Optional

# ロガー設定
logger = logging.getLogger(__name__)

# ====================================================================
# 定数定義
# ====================================================================

# 有効なレイアウト名一覧
VALID_LAYOUTS = frozenset([
    "title", "section", "toc", "bullet_points", "numbered_list",
    "two_column", "three_column", "four_column", "metrics",
    "quote", "faq", "comparison_table", "image_with_text",
    "chart", "cta",
])

# デフォルトレイアウト
DEFAULT_LAYOUT = "bullet_points"

# 正規表現パターン
LAYOUT_PATTERN = re.compile(r'<!--\s*layout:\s*(\w+)\s*-->', re.IGNORECASE)
CHART_TYPE_PATTERN = re.compile(r'<!--\s*chart_type:\s*(\w+)\s*-->', re.IGNORECASE)
H1_PATTERN = re.compile(r'^#\s+(.+)$')
H2_PATTERN = re.compile(r'^##\s+(.+)$')
H3_PATTERN = re.compile(r'^###\s+(.+)$')
H4_PATTERN = re.compile(r'^####\s+(.+)$')
IMAGE_PATTERN = re.compile(r'^!\[([^\]]*)\]\(([^)]+)\)$')
BOLD_PATTERN = re.compile(r'\*\*(.+?)\*\*')
# CTAボタン: [テキスト] だが [テキスト](url) はリンクなので除外
CTA_BUTTON_PATTERN = re.compile(r'^\[([^\]]+)\](?!\()$')

# 時系列パターン（チャートタイプ自動判定用）
TIME_SERIES_PATTERNS = [
    re.compile(r'^\d{4}年?$'),           # 2021, 2021年
    re.compile(r'^\d{1,2}月$'),          # 1月, 12月
    re.compile(r'^Q[1-4]$'),             # Q1, Q2, Q3, Q4
    re.compile(r'^\d{4}[/-]\d{1,2}$'),   # 2021/1, 2021-01
    re.compile(r'^FY\d{2,4}$'),          # FY21, FY2021
    re.compile(r'^\d{4}年度$'),          # 2021年度
]

# 文字数制限テーブル（レイアウト別）
# 各エントリは (フィールド名, 制限値) のリスト
CHARACTER_LIMITS: dict[str, list[tuple[str, int]]] = {
    "title": [("タイトル", 30), ("サブタイトル", 50)],
    "section": [("タイトル", 20)],
    "toc": [("タイトル", 25), ("キーメッセージ", 40), ("項目", 20)],
    "bullet_points": [("タイトル", 25), ("キーメッセージ", 40), ("項目", 50)],
    "numbered_list": [("タイトル", 25), ("キーメッセージ", 40), ("項目タイトル", 15), ("項目説明", 50)],
    "two_column": [("タイトル", 25), ("キーメッセージ", 40), ("カラム見出し", 10), ("カラム本文", 80)],
    "three_column": [("タイトル", 25), ("キーメッセージ", 40), ("カラム見出し", 8), ("カラム本文", 60)],
    "four_column": [("タイトル", 25), ("キーメッセージ", 40), ("カラム見出し", 6), ("カラム本文", 40)],
    "metrics": [("タイトル", 25), ("キーメッセージ", 40), ("数値", 10), ("ラベル", 15)],
    "quote": [("タイトル", 25), ("引用", 100), ("著者", 20)],
    "faq": [("タイトル", 25), ("質問", 30), ("回答", 80)],
    "comparison_table": [("タイトル", 25), ("キーメッセージ", 40), ("ヘッダ", 10), ("データ", 15)],
    "image_with_text": [("タイトル", 25), ("キーメッセージ", 40), ("説明", 100)],
    "chart": [("タイトル", 25), ("キーメッセージ", 40), ("ラベル", 10)],
    "cta": [("メッセージ", 20), ("説明", 40), ("ボタン", 15)],
}

# Markdown記法を除去するための正規表現
MARKDOWN_STRIP_PATTERNS = [
    re.compile(r'\*\*(.+?)\*\*'),   # **太字**
    re.compile(r'\*(.+?)\*'),        # *斜体*
    re.compile(r'`(.+?)`'),          # `コード`
    re.compile(r'^#{1,6}\s+'),       # 見出し
    re.compile(r'^\s*[-*+]\s+'),     # 箇条書き
    re.compile(r'^\s*\d+\.\s+'),     # 番号リスト
]


# ====================================================================
# データクラス
# ====================================================================

@dataclass
class ChartData:
    """グラフデータを保持するデータクラス"""

    chart_type: str = "horizontal_bar"       # "bar" | "horizontal_bar" | "pie" | "line"
    labels: list[str] = field(default_factory=list)        # 行ラベル
    series: list[dict] = field(default_factory=list)       # 系列データ [{"name": str, "values": [float]}]
    highlight_indices: list[int] = field(default_factory=list)  # 強調表示するインデックス


@dataclass
class SlideContent:
    """1スライド分のコンテンツを保持するデータクラス"""

    layout: str = "bullet_points"
    title: str = ""
    subtitle: str = ""
    body: str = ""
    columns: list[dict] = field(default_factory=list)
    items: list[dict] = field(default_factory=list)
    table: list[list[str]] = field(default_factory=list)
    chart: Optional[ChartData] = None
    image_path: str = ""
    quote: str = ""
    quote_author: str = ""
    faq_items: list[dict] = field(default_factory=list)
    cta_button: str = ""


@dataclass
class Presentation:
    """プレゼンテーション全体を保持するデータクラス"""

    slides: list[SlideContent] = field(default_factory=list)
    title: str = ""


# ====================================================================
# ユーティリティ関数
# ====================================================================

def _strip_markdown(text: str) -> str:
    """Markdown記法を除去してプレーンテキストにする"""
    result = text
    # **太字** → 中身のみ
    result = re.sub(r'\*\*(.+?)\*\*', r'\1', result)
    # *斜体* → 中身のみ
    result = re.sub(r'\*(.+?)\*', r'\1', result)
    # `コード` → 中身のみ
    result = re.sub(r'`(.+?)`', r'\1', result)
    # 見出し記号
    result = re.sub(r'^#{1,6}\s+', '', result)
    return result


def count_display_width(text: str) -> int:
    """テキストの表示幅を計算する

    全角文字（East Asian Width が W, F）= 2、半角文字 = 1として計算する。
    Markdown記法は除去してからカウントする。

    Args:
        text: 計算対象のテキスト

    Returns:
        表示幅の合計値
    """
    # Markdown記法を除去
    plain = _strip_markdown(text)
    width = 0
    for ch in plain:
        eaw = unicodedata.east_asian_width(ch)
        if eaw in ('W', 'F'):
            width += 2
        else:
            width += 1
    return width


def validate_character_limits(presentation: Presentation) -> list[str]:
    """プレゼンテーション内の全スライドについて文字数制限を検証する

    各レイアウトの制限に基づいて超過箇所を検出し、
    警告メッセージのリストを返す。

    Args:
        presentation: 検証対象のPresentationオブジェクト

    Returns:
        警告メッセージのリスト。制限内の場合は空リスト
    """
    warnings = []
    for i, slide in enumerate(presentation.slides, 1):
        layout = slide.layout
        limits = CHARACTER_LIMITS.get(layout)
        if not limits:
            continue

        for field_name, limit in limits:
            # フィールドに対応するテキスト値を取得してチェック
            texts_to_check = _get_texts_for_field(slide, field_name)
            for text in texts_to_check:
                actual = count_display_width(text)
                if actual > limit:
                    warnings.append(
                        f"スライド{i} ({layout}): {field_name}が制限{limit}に対して{actual}です"
                    )
    return warnings


def _get_texts_for_field(slide: SlideContent, field_name: str) -> list[str]:
    """バリデーション用: フィールド名に対応するテキスト値をリストで返す"""
    layout = slide.layout

    if field_name == "タイトル" or field_name == "メッセージ":
        return [slide.title] if slide.title else []
    elif field_name == "キーメッセージ" or field_name == "サブタイトル" or field_name == "説明":
        if field_name == "説明" and layout == "image_with_text":
            return [slide.body] if slide.body else []
        return [slide.subtitle] if slide.subtitle else []
    elif field_name == "項目":
        return [item.get("title", "") for item in slide.items if item.get("title")]
    elif field_name == "項目タイトル":
        return [item.get("title", "") for item in slide.items if item.get("title")]
    elif field_name == "項目説明":
        return [item.get("description", "") for item in slide.items if item.get("description")]
    elif field_name == "カラム見出し":
        return [col.get("heading", "") for col in slide.columns if col.get("heading")]
    elif field_name == "カラム本文":
        return [col.get("body", "") for col in slide.columns if col.get("body")]
    elif field_name == "数値":
        return [item.get("value", "") for item in slide.items if item.get("value")]
    elif field_name == "ラベル":
        return [item.get("label", "") for item in slide.items if item.get("label")]
    elif field_name == "引用":
        return [slide.quote] if slide.quote else []
    elif field_name == "著者":
        return [slide.quote_author] if slide.quote_author else []
    elif field_name == "質問":
        return [item.get("question", "") for item in slide.faq_items if item.get("question")]
    elif field_name == "回答":
        return [item.get("answer", "") for item in slide.faq_items if item.get("answer")]
    elif field_name == "ヘッダ":
        if slide.table and len(slide.table) > 0:
            return list(slide.table[0])
        return []
    elif field_name == "データ":
        if slide.table and len(slide.table) > 1:
            texts = []
            for row in slide.table[1:]:
                texts.extend(row)
            return texts
        return []
    elif field_name == "ボタン":
        return [slide.cta_button] if slide.cta_button else []
    elif field_name == "ラベル" and layout == "chart":
        if slide.chart:
            return list(slide.chart.labels)
        return []
    return []


# ====================================================================
# SlidesMarkdownParser
# ====================================================================

class SlidesMarkdownParser:
    """slides.md形式のMarkdownをパースしてPresentationに変換する"""

    def parse(self, markdown_text: str) -> Presentation:
        """Markdownテキストを解析してPresentationオブジェクトを返す

        Args:
            markdown_text: slides.md形式のMarkdownテキスト

        Returns:
            解析結果のPresentationオブジェクト
        """
        # ---でスライドブロックに分割
        blocks = re.split(r'^-{3,}\s*$', markdown_text, flags=re.MULTILINE)

        slides = []
        pres_title = ""

        for block in blocks:
            # 空ブロックをスキップ
            stripped = block.strip()
            if not stripped:
                continue

            slide = self._parse_block(stripped)
            if slide:
                slides.append(slide)
                # 最初のtitleレイアウトからプレゼンタイトルを取得
                if not pres_title and slide.layout == "title":
                    pres_title = slide.title

        return Presentation(slides=slides, title=pres_title)

    def _parse_block(self, block_text: str) -> Optional[SlideContent]:
        """1ブロック分のテキストをパースしてSlideContentを返す"""
        lines = block_text.split('\n')

        # レイアウト抽出
        layout = self._extract_layout(lines)

        # h1で始まる場合はsectionレイアウトに上書き
        content_lines = [l for l in lines if not LAYOUT_PATTERN.match(l.strip())
                         and not CHART_TYPE_PATTERN.match(l.strip())]
        for line in content_lines:
            if H1_PATTERN.match(line.strip()):
                layout = "section"
                break

        # chart_typeコメントの抽出
        chart_type_explicit = None
        for line in lines:
            m = CHART_TYPE_PATTERN.match(line.strip())
            if m:
                chart_type_explicit = m.group(1).lower()

        # 不明なレイアウトへのフォールバック
        if layout not in VALID_LAYOUTS:
            logger.warning(f"不明なレイアウト '{layout}' が指定されました。bullet_pointsにフォールバックします。")
            layout = DEFAULT_LAYOUT

        # レイアウト別のパース処理を呼び出し
        try:
            if layout == "title":
                return self._parse_title(content_lines)
            elif layout == "section":
                return self._parse_section(content_lines)
            elif layout == "toc":
                return self._parse_toc(content_lines)
            elif layout == "bullet_points":
                return self._parse_bullet_points(content_lines)
            elif layout == "numbered_list":
                return self._parse_numbered_list(content_lines)
            elif layout == "two_column":
                return self._parse_columns(content_lines, 2)
            elif layout == "three_column":
                return self._parse_columns(content_lines, 3)
            elif layout == "four_column":
                return self._parse_columns(content_lines, 4)
            elif layout == "metrics":
                return self._parse_metrics(content_lines)
            elif layout == "quote":
                return self._parse_quote(content_lines)
            elif layout == "faq":
                return self._parse_faq(content_lines)
            elif layout == "comparison_table":
                return self._parse_comparison_table(content_lines)
            elif layout == "image_with_text":
                return self._parse_image_with_text(content_lines)
            elif layout == "chart":
                return self._parse_chart(content_lines, chart_type_explicit)
            elif layout == "cta":
                return self._parse_cta(content_lines)
        except Exception as e:
            logger.warning(f"スライドのパース中にエラーが発生しました: {e}")
            return SlideContent(layout=layout)

        return SlideContent(layout=layout)

    def _extract_layout(self, lines: list[str]) -> str:
        """行リストからレイアウト名を抽出する"""
        for line in lines:
            m = LAYOUT_PATTERN.match(line.strip())
            if m:
                return m.group(1).lower()
        return DEFAULT_LAYOUT

    def _extract_h1(self, lines: list[str]) -> str:
        """行リストからh1見出しを抽出する"""
        for line in lines:
            m = H1_PATTERN.match(line.strip())
            if m:
                return m.group(1).strip()
        return ""

    def _extract_h2(self, lines: list[str]) -> str:
        """行リストからh2見出しを抽出する"""
        for line in lines:
            m = H2_PATTERN.match(line.strip())
            if m:
                return m.group(1).strip()
        return ""

    def _extract_h3(self, lines: list[str]) -> str:
        """行リストからh3見出しを抽出する"""
        for line in lines:
            m = H3_PATTERN.match(line.strip())
            if m:
                return m.group(1).strip()
        return ""

    def _extract_h4s(self, lines: list[str]) -> list[str]:
        """行リストから全てのh4見出しを抽出する"""
        headings = []
        for line in lines:
            m = H4_PATTERN.match(line.strip())
            if m:
                headings.append(m.group(1).strip())
        return headings

    def _get_body_lines(self, lines: list[str]) -> list[str]:
        """見出し行やHTMLコメント以外の本文行を取得する"""
        body = []
        for line in lines:
            stripped = line.strip()
            if (not stripped or
                H1_PATTERN.match(stripped) or
                H2_PATTERN.match(stripped) or
                H3_PATTERN.match(stripped) or
                H4_PATTERN.match(stripped) or
                LAYOUT_PATTERN.match(stripped) or
                CHART_TYPE_PATTERN.match(stripped)):
                continue
            body.append(line)
        return body

    def _parse_bullet_items(self, lines: list[str]) -> list[str]:
        """箇条書き行（-で始まる行）のテキストを抽出する"""
        items = []
        for line in lines:
            stripped = line.strip()
            if re.match(r'^[-*+]\s+', stripped):
                text = re.sub(r'^[-*+]\s+', '', stripped)
                items.append(text)
        return items

    def _parse_markdown_table(self, lines: list[str]) -> list[list[str]]:
        """Markdownテーブルを2次元リストにパースする

        区切り行（|---|---|）は除外する。
        """
        table = []
        for line in lines:
            stripped = line.strip()
            if not stripped.startswith('|'):
                continue
            # 区切り行を除外（セル内容がすべて - やスペースのみ）
            cells = [c.strip() for c in stripped.split('|')]
            # 先頭と末尾の空要素を除去（|で始まり|で終わるため）
            cells = [c for c in cells if c != '']
            if not cells:
                continue
            # 区切り行判定: 全セルが -+ やスペースのみで構成
            is_separator = all(re.match(r'^[-:]+$', c) for c in cells)
            if is_separator:
                continue
            table.append(cells)
        return table

    # ----------------------------------------------------------------
    # レイアウト別パーサーメソッド
    # ----------------------------------------------------------------

    def _parse_title(self, lines: list[str]) -> SlideContent:
        """titleレイアウトのパース"""
        title = self._extract_h2(lines)
        subtitle = self._extract_h3(lines)
        # 見出し以外の本文を取得
        body_lines = []
        for line in lines:
            stripped = line.strip()
            if (not stripped or
                H2_PATTERN.match(stripped) or
                H3_PATTERN.match(stripped)):
                continue
            body_lines.append(stripped)
        body = '\n'.join(body_lines).strip()
        return SlideContent(layout="title", title=title, subtitle=subtitle, body=body)

    def _parse_section(self, lines: list[str]) -> SlideContent:
        """sectionレイアウトのパース"""
        # h1があればそれをタイトルに、なければh2を使用
        title = self._extract_h1(lines)
        if not title:
            title = self._extract_h2(lines)
        return SlideContent(layout="section", title=title)

    def _parse_toc(self, lines: list[str]) -> SlideContent:
        """tocレイアウトのパース"""
        title = self._extract_h2(lines)
        subtitle = self._extract_h3(lines)
        bullet_texts = self._parse_bullet_items(lines)
        items = [{"title": t, "description": ""} for t in bullet_texts]
        return SlideContent(layout="toc", title=title, subtitle=subtitle, items=items)

    def _parse_bullet_points(self, lines: list[str]) -> SlideContent:
        """bullet_pointsレイアウトのパース（ネスト対応）"""
        title = self._extract_h2(lines)
        subtitle = self._extract_h3(lines)
        items = []
        current_item = None

        for line in lines:
            stripped = line.strip()
            # トップレベルの箇条書き
            m = re.match(r'^[-*+]\s+(.+)$', stripped)
            if m and not line.startswith('  '):
                # 前のアイテムがあれば保存
                if current_item is not None:
                    items.append(current_item)
                current_item = {"title": m.group(1), "description": ""}
                continue
            # ネストした箇条書き（2スペース以上のインデント）
            nested_m = re.match(r'^\s{2,}[-*+]\s+(.+)$', line)
            if nested_m and current_item is not None:
                if current_item["description"]:
                    current_item["description"] += '\n' + nested_m.group(1)
                else:
                    current_item["description"] = nested_m.group(1)

        # 最後のアイテムを追加
        if current_item is not None:
            items.append(current_item)

        return SlideContent(layout="bullet_points", title=title, subtitle=subtitle, items=items)

    def _parse_numbered_list(self, lines: list[str]) -> SlideContent:
        """numbered_listレイアウトのパース

        形式: `1. **タイトル**\n   説明文` または `1. **タイトル**: 説明文`
        """
        title = self._extract_h2(lines)
        subtitle = self._extract_h3(lines)
        items = []
        current_item = None

        for line in lines:
            stripped = line.strip()
            # 番号付きリストの開始行
            m = re.match(r'^\d+\.\s+(.+)$', stripped)
            if m:
                # 前のアイテムがあれば保存
                if current_item is not None:
                    items.append(current_item)

                content = m.group(1)
                # **太字** からタイトルを抽出
                bold_m = BOLD_PATTERN.search(content)
                if bold_m:
                    item_title = bold_m.group(1)
                    # タイトルの後の説明（: または改行後）
                    # "**タイトル**: 説明" の形式
                    after_bold = content[bold_m.end():].strip()
                    # 先頭のコロンやハイフンを除去
                    after_bold = re.sub(r'^[:\s]+', '', after_bold).strip()
                    current_item = {"title": item_title, "description": after_bold}
                else:
                    # 太字がない場合はコロンで分割を試みる
                    if ':' in content or '：' in content:
                        parts = re.split(r'[:：]\s*', content, maxsplit=1)
                        current_item = {"title": parts[0].strip(), "description": parts[1].strip() if len(parts) > 1 else ""}
                    else:
                        current_item = {"title": content, "description": ""}
                continue

            # 説明の継続行（インデント付き）
            if current_item is not None and stripped and not H2_PATTERN.match(stripped) and not H3_PATTERN.match(stripped):
                cont_m = re.match(r'^\s{2,}(.+)$', line)
                if cont_m:
                    if current_item["description"]:
                        current_item["description"] += '\n' + cont_m.group(1).strip()
                    else:
                        current_item["description"] = cont_m.group(1).strip()

        # 最後のアイテムを追加
        if current_item is not None:
            items.append(current_item)

        return SlideContent(layout="numbered_list", title=title, subtitle=subtitle, items=items)

    def _parse_columns(self, lines: list[str], n: int) -> SlideContent:
        """two/three/four_columnレイアウトのパース

        ####見出しとその後の本文からカラムデータを構築する。
        """
        layout_names = {2: "two_column", 3: "three_column", 4: "four_column"}
        layout = layout_names.get(n, "two_column")
        title = self._extract_h2(lines)
        subtitle = self._extract_h3(lines)

        columns = []
        current_heading = None
        current_body_lines: list[str] = []

        for line in lines:
            stripped = line.strip()
            m = H4_PATTERN.match(stripped)
            if m:
                # 前のカラムがあれば保存
                if current_heading is not None:
                    columns.append({
                        "heading": current_heading,
                        "body": '\n'.join(current_body_lines).strip()
                    })
                current_heading = m.group(1).strip()
                current_body_lines = []
                continue

            # h4の後の本文行を収集（他の見出しやコメント以外）
            if current_heading is not None:
                if (not H1_PATTERN.match(stripped) and
                    not H2_PATTERN.match(stripped) and
                    not H3_PATTERN.match(stripped)):
                    current_body_lines.append(stripped)

        # 最後のカラムを追加
        if current_heading is not None:
            columns.append({
                "heading": current_heading,
                "body": '\n'.join(current_body_lines).strip()
            })

        return SlideContent(layout=layout, title=title, subtitle=subtitle, columns=columns)

    def _parse_metrics(self, lines: list[str]) -> SlideContent:
        """metricsレイアウトのパース

        形式: `- **数値** ラベル`
        """
        title = self._extract_h2(lines)
        subtitle = self._extract_h3(lines)
        items = []

        for line in lines:
            stripped = line.strip()
            # `- **数値** ラベル` 形式
            m = re.match(r'^[-*+]\s+\*\*(.+?)\*\*\s+(.+)$', stripped)
            if m:
                items.append({"value": m.group(1), "label": m.group(2)})
                continue
            # テーブル形式のフォールバック（layout-rules.mdの形式）
            # テーブルパースは別途行うため、ここではスキップ

        # テーブル形式の場合のフォールバック
        if not items:
            table = self._parse_markdown_table(lines)
            if table and len(table) > 1:
                # ヘッダー行を確認して数値/ラベルカラムを特定
                for row in table[1:]:
                    if len(row) >= 2:
                        items.append({"value": row[0], "label": row[1]})

        return SlideContent(layout="metrics", title=title, subtitle=subtitle, items=items)

    def _parse_quote(self, lines: list[str]) -> SlideContent:
        """quoteレイアウトのパース"""
        title = self._extract_h2(lines)
        quote_lines = []
        quote_author = ""

        for line in lines:
            stripped = line.strip()
            # 引用行
            if stripped.startswith('>'):
                content = stripped[1:].strip()
                # 発言者判定（— または -- で始まる行）
                if re.match(r'^[—–-]{1,2}\s*', content):
                    quote_author = re.sub(r'^[—–-]{1,2}\s*', '', content).strip()
                else:
                    quote_lines.append(content)
                continue
            # 引用ブロック外の発言者行（`-- 著者` 形式）
            if re.match(r'^[—–-]{1,2}\s+.+', stripped):
                quote_author = re.sub(r'^[—–-]{1,2}\s*', '', stripped).strip()

        quote = '\n'.join(quote_lines).strip()
        return SlideContent(layout="quote", title=title, quote=quote, quote_author=quote_author)

    def _parse_faq(self, lines: list[str]) -> SlideContent:
        """faqレイアウトのパース

        形式: `**Q: 質問文**` / `A: 回答文`
        """
        title = self._extract_h2(lines)
        faq_items = []
        current_question = None
        current_answer_lines: list[str] = []

        for line in lines:
            stripped = line.strip()
            # 質問行: **Q: ... ** または **Q. ... **
            q_match = re.match(r'^\*\*Q[.:：]\s*(.+?)\*\*$', stripped)
            if q_match:
                # 前のQ&Aペアを保存
                if current_question is not None:
                    faq_items.append({
                        "question": current_question,
                        "answer": '\n'.join(current_answer_lines).strip()
                    })
                current_question = q_match.group(1).strip()
                current_answer_lines = []
                continue

            # 回答行: A: ... または A. ...
            a_match = re.match(r'^A[.:：]\s*(.+)$', stripped)
            if a_match and current_question is not None:
                current_answer_lines.append(a_match.group(1).strip())
                continue

            # 回答の継続行
            if current_question is not None and stripped and not H2_PATTERN.match(stripped):
                # 見出しや質問行でなければ回答の継続
                if not re.match(r'^\*\*Q[.:：]', stripped):
                    current_answer_lines.append(stripped)

        # 最後のQ&Aペアを追加
        if current_question is not None:
            faq_items.append({
                "question": current_question,
                "answer": '\n'.join(current_answer_lines).strip()
            })

        return SlideContent(layout="faq", title=title, faq_items=faq_items)

    def _parse_comparison_table(self, lines: list[str]) -> SlideContent:
        """comparison_tableレイアウトのパース"""
        title = self._extract_h2(lines)
        subtitle = self._extract_h3(lines)
        table = self._parse_markdown_table(lines)
        return SlideContent(layout="comparison_table", title=title, subtitle=subtitle, table=table)

    def _parse_image_with_text(self, lines: list[str]) -> SlideContent:
        """image_with_textレイアウトのパース"""
        title = self._extract_h2(lines)
        subtitle = self._extract_h3(lines)
        image_path = ""
        body_lines = []
        image_found = False

        for line in lines:
            stripped = line.strip()
            # 画像行
            m = IMAGE_PATTERN.match(stripped)
            if m:
                image_path = m.group(2)
                image_found = True
                continue
            # 見出し以外の行を本文として収集（画像の後のみ）
            if image_found and stripped:
                if (not H2_PATTERN.match(stripped) and
                    not H3_PATTERN.match(stripped)):
                    body_lines.append(stripped)

        body = '\n'.join(body_lines).strip()
        return SlideContent(
            layout="image_with_text", title=title, subtitle=subtitle,
            image_path=image_path, body=body
        )

    def _parse_chart(self, lines: list[str], explicit_type: Optional[str] = None) -> SlideContent:
        """chartレイアウトのパース

        Markdownテーブルをパースし、ChartDataに変換する。
        """
        title = self._extract_h2(lines)
        subtitle = self._extract_h3(lines)
        table = self._parse_markdown_table(lines)

        chart = None
        if table and len(table) > 1:
            chart = self._build_chart_data(table, explicit_type)

        return SlideContent(
            layout="chart", title=title, subtitle=subtitle,
            table=table, chart=chart
        )

    def _parse_cta(self, lines: list[str]) -> SlideContent:
        """ctaレイアウトのパース"""
        title = self._extract_h2(lines)
        subtitle = self._extract_h3(lines)
        cta_button = ""

        for line in lines:
            stripped = line.strip()
            # [ボタンテキスト] 形式（リンクではない）
            m = CTA_BUTTON_PATTERN.match(stripped)
            if m:
                cta_button = m.group(1)
                break
            # 箇条書き形式のフォールバック（layout-rules.mdの形式）
            bullet_m = re.match(r'^[-*+]\s+(.+)$', stripped)
            if bullet_m and not cta_button:
                cta_button = bullet_m.group(1)

        return SlideContent(layout="cta", title=title, subtitle=subtitle, cta_button=cta_button)

    # ----------------------------------------------------------------
    # チャート関連
    # ----------------------------------------------------------------

    def _build_chart_data(self, table: list[list[str]], explicit_type: Optional[str] = None) -> ChartData:
        """テーブルデータからChartDataを構築する"""
        headers = table[0]
        data_rows = table[1:]

        # ラベル（1列目）
        labels = []
        highlight_indices = []
        for i, row in enumerate(data_rows):
            if not row:
                continue
            label = row[0]
            # 太字ラベルの検出
            bold_m = BOLD_PATTERN.search(label)
            if bold_m:
                labels.append(bold_m.group(1))
                highlight_indices.append(i)
            else:
                labels.append(label)

        # 系列データ（2列目以降）
        series = []
        for col_idx in range(1, len(headers)):
            values = []
            for row in data_rows:
                if col_idx < len(row):
                    values.append(self._extract_number(row[col_idx]))
                else:
                    values.append(0.0)
            series.append({
                "name": headers[col_idx],
                "values": values
            })

        # チャートタイプ判定
        chart_type = self._detect_chart_type(table, explicit_type)

        return ChartData(
            chart_type=chart_type,
            labels=labels,
            series=series,
            highlight_indices=highlight_indices
        )

    def _detect_chart_type(self, table: list[list[str]], explicit_type: Optional[str] = None) -> str:
        """テーブルデータからチャートタイプを自動判定する

        判定優先順位:
        1. explicit_type が指定されている場合はそれを使用
        2. データ値の判定: %表記 or 合計≒100 → "pie"
        3. ラベルの判定: 時系列パターン → "bar"
        4. デフォルト → "horizontal_bar"
        """
        # 1. 明示指定
        if explicit_type:
            return explicit_type

        data_rows = table[1:]

        # 2. パーセント/合計≒100判定
        has_percent = False
        all_values = []
        for row in data_rows:
            for cell in row[1:]:
                if '%' in cell:
                    has_percent = True
                val = self._extract_number(cell)
                all_values.append(val)

        if has_percent:
            return "pie"
        if all_values and 95 <= sum(all_values) <= 105:
            return "pie"

        # 3. 時系列パターン判定
        labels = []
        for row in data_rows:
            if row:
                # 太字を除去してからチェック
                label = BOLD_PATTERN.sub(r'\1', row[0]).strip()
                labels.append(label)

        if labels and self._is_time_series(labels):
            return "bar"

        # 4. デフォルト
        return "horizontal_bar"

    def _is_time_series(self, labels: list[str]) -> bool:
        """ラベルリストが時系列パターンに一致するか判定する"""
        match_count = 0
        for label in labels:
            for pattern in TIME_SERIES_PATTERNS:
                if pattern.match(label):
                    match_count += 1
                    break
        # 半数以上が時系列パターンに一致すれば時系列と判定
        return match_count > len(labels) / 2

    def _extract_number(self, text: str) -> float:
        """テキストから数値を抽出する

        太字マークダウン記法、カンマ区切り、単位を除去してfloatに変換する。
        変換できない場合は0.0を返す。
        """
        # **太字**を除去
        cleaned = BOLD_PATTERN.sub(r'\1', text).strip()
        # カンマ区切りを除去
        cleaned = cleaned.replace(',', '')
        # 数値部分を抽出（先頭の数値+小数点）
        m = re.search(r'[-+]?\d+\.?\d*', cleaned)
        if m:
            try:
                return float(m.group(0))
            except ValueError:
                return 0.0
        return 0.0


# ====================================================================
# SlidesMarkdownSerializer
# ====================================================================

class SlidesMarkdownSerializer:
    """PresentationオブジェクトをMarkdown形式にシリアライズする"""

    def serialize(self, presentation: Presentation) -> str:
        """PresentationオブジェクトをMarkdownテキストに変換する

        Args:
            presentation: シリアライズ対象のPresentationオブジェクト

        Returns:
            slides.md形式のMarkdownテキスト
        """
        slide_blocks = []
        for slide in presentation.slides:
            block = self._serialize_slide(slide)
            slide_blocks.append(block)
        return '\n\n---\n\n'.join(slide_blocks) + '\n'

    def _serialize_slide(self, slide: SlideContent) -> str:
        """1スライド分のSlideContentをMarkdownテキストに変換する"""
        layout = slide.layout
        if layout == "title":
            return self._serialize_title(slide)
        elif layout == "section":
            return self._serialize_section(slide)
        elif layout == "toc":
            return self._serialize_toc(slide)
        elif layout == "bullet_points":
            return self._serialize_bullet_points(slide)
        elif layout == "numbered_list":
            return self._serialize_numbered_list(slide)
        elif layout in ("two_column", "three_column", "four_column"):
            return self._serialize_columns(slide)
        elif layout == "metrics":
            return self._serialize_metrics(slide)
        elif layout == "quote":
            return self._serialize_quote(slide)
        elif layout == "faq":
            return self._serialize_faq(slide)
        elif layout == "comparison_table":
            return self._serialize_comparison_table(slide)
        elif layout == "image_with_text":
            return self._serialize_image_with_text(slide)
        elif layout == "chart":
            return self._serialize_chart(slide)
        elif layout == "cta":
            return self._serialize_cta(slide)
        else:
            # フォールバック: bullet_points形式で出力
            return self._serialize_bullet_points(slide)

    def _serialize_title(self, slide: SlideContent) -> str:
        """titleレイアウトのシリアライズ"""
        parts = ["<!-- layout: title -->"]
        if slide.title:
            parts.append(f"## {slide.title}")
        if slide.subtitle:
            parts.append(f"### {slide.subtitle}")
        if slide.body:
            parts.append("")
            parts.append(slide.body)
        return '\n'.join(parts)

    def _serialize_section(self, slide: SlideContent) -> str:
        """sectionレイアウトのシリアライズ"""
        if slide.title:
            return f"# {slide.title}"
        return "# "

    def _serialize_toc(self, slide: SlideContent) -> str:
        """tocレイアウトのシリアライズ"""
        parts = ["<!-- layout: toc -->"]
        if slide.title:
            parts.append(f"## {slide.title}")
        if slide.subtitle:
            parts.append(f"### {slide.subtitle}")
        if slide.items:
            parts.append("")
            for item in slide.items:
                parts.append(f"- {item.get('title', '')}")
        return '\n'.join(parts)

    def _serialize_bullet_points(self, slide: SlideContent) -> str:
        """bullet_pointsレイアウトのシリアライズ"""
        parts = ["<!-- layout: bullet_points -->"]
        if slide.title:
            parts.append(f"## {slide.title}")
        if slide.subtitle:
            parts.append(f"### {slide.subtitle}")
        if slide.items:
            parts.append("")
            for item in slide.items:
                parts.append(f"- {item.get('title', '')}")
                desc = item.get('description', '')
                if desc:
                    for sub_line in desc.split('\n'):
                        parts.append(f"  - {sub_line}")
        return '\n'.join(parts)

    def _serialize_numbered_list(self, slide: SlideContent) -> str:
        """numbered_listレイアウトのシリアライズ"""
        parts = ["<!-- layout: numbered_list -->"]
        if slide.title:
            parts.append(f"## {slide.title}")
        if slide.subtitle:
            parts.append(f"### {slide.subtitle}")
        if slide.items:
            parts.append("")
            for i, item in enumerate(slide.items, 1):
                title = item.get('title', '')
                desc = item.get('description', '')
                if desc:
                    parts.append(f"{i}. **{title}**: {desc}")
                else:
                    parts.append(f"{i}. **{title}**")
        return '\n'.join(parts)

    def _serialize_columns(self, slide: SlideContent) -> str:
        """two/three/four_columnレイアウトのシリアライズ"""
        parts = [f"<!-- layout: {slide.layout} -->"]
        if slide.title:
            parts.append(f"## {slide.title}")
        if slide.subtitle:
            parts.append(f"### {slide.subtitle}")
        if slide.columns:
            parts.append("")
            for col in slide.columns:
                heading = col.get('heading', '')
                body = col.get('body', '')
                parts.append(f"#### {heading}")
                if body:
                    parts.append(body)
                parts.append("")
        return '\n'.join(parts).rstrip()

    def _serialize_metrics(self, slide: SlideContent) -> str:
        """metricsレイアウトのシリアライズ"""
        parts = ["<!-- layout: metrics -->"]
        if slide.title:
            parts.append(f"## {slide.title}")
        if slide.subtitle:
            parts.append(f"### {slide.subtitle}")
        if slide.items:
            parts.append("")
            for item in slide.items:
                value = item.get('value', '')
                label = item.get('label', '')
                parts.append(f"- **{value}** {label}")
        return '\n'.join(parts)

    def _serialize_quote(self, slide: SlideContent) -> str:
        """quoteレイアウトのシリアライズ"""
        parts = ["<!-- layout: quote -->"]
        if slide.title:
            parts.append(f"## {slide.title}")
        if slide.quote:
            parts.append("")
            # 引用テキストの各行に > を付ける
            for line in slide.quote.split('\n'):
                parts.append(f"> {line}")
        if slide.quote_author:
            parts.append("")
            parts.append(f"-- {slide.quote_author}")
        return '\n'.join(parts)

    def _serialize_faq(self, slide: SlideContent) -> str:
        """faqレイアウトのシリアライズ"""
        parts = ["<!-- layout: faq -->"]
        if slide.title:
            parts.append(f"## {slide.title}")
        if slide.faq_items:
            parts.append("")
            for item in slide.faq_items:
                q = item.get('question', '')
                a = item.get('answer', '')
                parts.append(f"**Q. {q}**")
                parts.append(f"A. {a}")
                parts.append("")
        return '\n'.join(parts).rstrip()

    def _serialize_comparison_table(self, slide: SlideContent) -> str:
        """comparison_tableレイアウトのシリアライズ"""
        parts = ["<!-- layout: comparison_table -->"]
        if slide.title:
            parts.append(f"## {slide.title}")
        if slide.subtitle:
            parts.append(f"### {slide.subtitle}")
        if slide.table:
            parts.append("")
            parts.append(self._table_to_markdown(slide.table))
        return '\n'.join(parts)

    def _serialize_image_with_text(self, slide: SlideContent) -> str:
        """image_with_textレイアウトのシリアライズ"""
        parts = ["<!-- layout: image_with_text -->"]
        if slide.title:
            parts.append(f"## {slide.title}")
        if slide.subtitle:
            parts.append(f"### {slide.subtitle}")
        if slide.image_path:
            parts.append("")
            parts.append(f"![image]({slide.image_path})")
        if slide.body:
            parts.append("")
            parts.append(slide.body)
        return '\n'.join(parts)

    def _serialize_chart(self, slide: SlideContent) -> str:
        """chartレイアウトのシリアライズ"""
        parts = ["<!-- layout: chart -->"]
        # lineタイプの場合はchart_typeコメントを追加
        if slide.chart and slide.chart.chart_type == "line":
            parts.append("<!-- chart_type: line -->")
        if slide.title:
            parts.append(f"## {slide.title}")
        if slide.subtitle:
            parts.append(f"### {slide.subtitle}")
        if slide.table:
            parts.append("")
            # チャートの場合はtableデータからテーブルを再構成
            # highlight_indicesを考慮して太字ラベルを復元
            table_with_bold = self._restore_bold_labels(slide.table, slide.chart)
            parts.append(self._table_to_markdown(table_with_bold))
        return '\n'.join(parts)

    def _serialize_cta(self, slide: SlideContent) -> str:
        """ctaレイアウトのシリアライズ"""
        parts = ["<!-- layout: cta -->"]
        if slide.title:
            parts.append(f"## {slide.title}")
        if slide.subtitle:
            parts.append(f"### {slide.subtitle}")
        if slide.cta_button:
            parts.append("")
            parts.append(f"[{slide.cta_button}]")
        return '\n'.join(parts)

    # ----------------------------------------------------------------
    # ヘルパー
    # ----------------------------------------------------------------

    def _table_to_markdown(self, table: list[list[str]]) -> str:
        """2次元リストをMarkdownテーブルに変換する"""
        if not table:
            return ""
        # 各列の最大幅を計算
        col_count = max(len(row) for row in table)
        col_widths = [0] * col_count
        for row in table:
            for i, cell in enumerate(row):
                col_widths[i] = max(col_widths[i], len(cell))

        lines = []
        for row_idx, row in enumerate(table):
            cells = []
            for i in range(col_count):
                cell = row[i] if i < len(row) else ""
                cells.append(f" {cell.ljust(col_widths[i])} ")
            lines.append("|" + "|".join(cells) + "|")
            # ヘッダー行の後に区切り行を追加
            if row_idx == 0:
                sep_cells = ["-" * (w + 2) for w in col_widths]
                lines.append("|" + "|".join(sep_cells) + "|")
        return '\n'.join(lines)

    def _restore_bold_labels(self, table: list[list[str]], chart: Optional[ChartData]) -> list[list[str]]:
        """チャートのhighlight_indicesに基づいてテーブルのラベルに太字マークを復元する"""
        if not chart or not chart.highlight_indices or not table:
            return table

        result = [row[:] for row in table]  # ディープコピー
        for idx in chart.highlight_indices:
            data_row_idx = idx + 1  # ヘッダー行を考慮
            if data_row_idx < len(result):
                label = result[data_row_idx][0]
                if not label.startswith('**'):
                    result[data_row_idx][0] = f"**{label}**"
        return result


# ====================================================================
# CLI
# ====================================================================

def main():
    """CLIエントリポイント"""
    parser = argparse.ArgumentParser(
        description="slides.md パーサー・シリアライザー"
    )
    parser.add_argument(
        "--input",
        type=str,
        help="slides.mdファイルのパス（パース実行、結果をJSON出力）",
    )
    parser.add_argument(
        "--validate",
        type=str,
        help="slides.mdファイルのパス（文字数制限チェック）",
    )

    args = parser.parse_args()

    if args.input:
        # パースモード: slides.md → JSON出力
        try:
            with open(args.input, encoding="utf-8") as f:
                markdown_text = f.read()
        except FileNotFoundError:
            print(f"エラー: ファイルが見つかりません: {args.input}", file=sys.stderr)
            sys.exit(1)
        except OSError as e:
            print(f"エラー: ファイルの読み込みに失敗しました: {e}", file=sys.stderr)
            sys.exit(1)

        md_parser = SlidesMarkdownParser()
        presentation = md_parser.parse(markdown_text)
        # 結果をJSON形式で標準出力に出力
        print(json.dumps(asdict(presentation), ensure_ascii=False, indent=2))

    elif args.validate:
        # バリデーションモード: 文字数制限チェック
        try:
            with open(args.validate, encoding="utf-8") as f:
                markdown_text = f.read()
        except FileNotFoundError:
            print(f"エラー: ファイルが見つかりません: {args.validate}", file=sys.stderr)
            sys.exit(1)
        except OSError as e:
            print(f"エラー: ファイルの読み込みに失敗しました: {e}", file=sys.stderr)
            sys.exit(1)

        md_parser = SlidesMarkdownParser()
        presentation = md_parser.parse(markdown_text)
        warnings = validate_character_limits(presentation)

        if warnings:
            for w in warnings:
                print(w, file=sys.stderr)
            sys.exit(1)
        else:
            print("文字数制限チェック: 問題なし")
            sys.exit(0)

    else:
        parser.print_help()


if __name__ == "__main__":
    main()
