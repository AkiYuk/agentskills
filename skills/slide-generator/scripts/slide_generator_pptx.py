"""PowerPointスライド生成スクリプト

slides.md形式のMarkdownファイルからPowerPoint（.pptx）ファイルを生成する。
15種類のレイアウトに対応し、config.jsonで指定されたデザイン設定を適用する。

使い方:
    python slide_generator_pptx.py --markdown-file output/slides.md --config config.json
    python slide_generator_pptx.py --markdown-file output/slides.md --config config.json --template assets/template.pptx
    python slide_generator_pptx.py --markdown-file output/slides.md --config config.json --output-dir my_output
"""

from __future__ import annotations

import argparse
import json
import os
import sys
from pathlib import Path

from pptx import Presentation as PptxPresentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN
from pptx.util import Inches, Pt

from slides_markdown import SlidesMarkdownParser, SlideContent


# ====================================================================
# DesignConfig
# ====================================================================

class DesignConfig:
    """config.jsonからデザイン設定を読み込むクラス"""

    def __init__(self, config_path: str = None):
        """設定を初期化する

        Args:
            config_path: config.jsonファイルのパス。Noneの場合はデフォルト値を使用
        """
        if config_path and os.path.exists(config_path):
            with open(config_path, encoding="utf-8") as f:
                self._config = json.load(f)
        else:
            self._config = {}

    @property
    def font_family(self) -> str:
        return self._config.get("font", {}).get("family", "Meiryo UI")

    @property
    def palette(self) -> dict:
        return self._config.get("palette", {})

    @property
    def primary_color(self) -> str:
        """プライマリカラー (#1E3A5F)"""
        return self.palette.get("primary", "#1E3A5F")

    @property
    def secondary_color(self) -> str:
        """セカンダリカラー (#4A6FA5)"""
        return self.palette.get("secondary", "#4A6FA5")

    @property
    def accent_color(self) -> str:
        """アクセントカラー (#3AA899)"""
        return self.palette.get("accent", "#3AA899")

    @property
    def text_primary(self) -> str:
        """テキスト色 プライマリ (#333333)"""
        return self.palette.get("text", {}).get("primary", "#333333")

    @property
    def text_secondary(self) -> str:
        """テキスト色 セカンダリ (#666666)"""
        return self.palette.get("text", {}).get("secondary", "#666666")

    @property
    def text_light(self) -> str:
        """テキスト色 ライト (#FFFFFF)"""
        return self.palette.get("text", {}).get("light", "#FFFFFF")

    @property
    def bg_primary(self) -> str:
        """背景色 プライマリ (#FFFFFF)"""
        return self.palette.get("background", {}).get("primary", "#FFFFFF")

    @property
    def bg_secondary(self) -> str:
        """背景色 セカンダリ (#F5F5F5)"""
        return self.palette.get("background", {}).get("secondary", "#F5F5F5")

    @property
    def bg_dark(self) -> str:
        """背景色 ダーク (#1E3A5F)"""
        return self.palette.get("background", {}).get("dark", "#1E3A5F")

    @property
    def output_dir(self) -> str:
        """出力ディレクトリ"""
        return self._config.get("output", {}).get("dir", "output")

    @property
    def template_path(self) -> str:
        """テンプレートPPTXのパス"""
        return self._config.get("template", {}).get("pptx_path", "")


# ====================================================================
# BaseRenderer
# ====================================================================

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
    CONTENT_HEIGHT = 5.45  # 7.5 - 1.55 - 0.5

    def render(self, slide, slide_data: SlideContent, config: DesignConfig):
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

    def _render_separator_line(self, slide, config: DesignConfig):
        """共通: タイトル・キーメッセージ下の区切り線"""
        self.add_horizontal_line(
            slide,
            left=Inches(self.CONTENT_LEFT),
            top=Inches(1.45),
            width=Inches(self.CONTENT_WIDTH),
            color="#CCCCCC",
            thickness=1.0,
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

    def set_background_gradient(self, slide, color1: str, color2: str):
        """グラデーション背景を設定（将来用）"""
        background = slide.background
        fill = background.fill
        fill.gradient()
        fill.gradient_stops[0].color.rgb = RGBColor.from_string(color1.lstrip("#"))
        fill.gradient_stops[1].color.rgb = RGBColor.from_string(color2.lstrip("#"))

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

    def set_font(self, text_frame, font_family: str = "Meiryo UI",
                 font_size=None, font_color: str = None,
                 bold: bool = None, italic: bool = None,
                 alignment=None, line_spacing: float = None):
        """テキストフレームのフォント設定ヘルパー"""
        for p in text_frame.paragraphs:
            if font_family:
                p.font.name = font_family
            if font_size is not None:
                p.font.size = font_size
            if font_color is not None:
                p.font.color.rgb = RGBColor.from_string(font_color.lstrip("#"))
            if bold is not None:
                p.font.bold = bold
            if italic is not None:
                p.font.italic = italic
            if alignment is not None:
                p.alignment = alignment
            if line_spacing is not None:
                p.line_spacing = line_spacing


# ====================================================================
# TitleRenderer
# ====================================================================

class TitleRenderer(BaseRenderer):
    """表紙レイアウトRenderer

    紺色背景に白文字のメインタイトルとサブタイトルを中央配置する。
    """

    def render(self, slide, slide_data: SlideContent, config: DesignConfig):
        """表紙スライドを描画"""
        # 背景: 紺色
        self.set_background_color(slide, config.bg_dark)

        # メインタイトル: 44pt, bold, 白色, 中央揃え
        self.add_text_box(
            slide,
            left=Inches(1.5),
            top=Inches(2.2),
            width=Inches(10.333),
            height=Inches(1.2),
            text=slide_data.title or "",
            font_size=Pt(44),
            font_color=config.text_light,
            bold=True,
            alignment=PP_ALIGN.CENTER,
            font_family=config.font_family,
        )

        # サブタイトル: 16pt, 白色, 中央揃え
        if slide_data.subtitle:
            self.add_text_box(
                slide,
                left=Inches(1.5),
                top=Inches(3.6),
                width=Inches(10.333),
                height=Inches(0.8),
                text=slide_data.subtitle,
                font_size=Pt(16),
                font_color=config.text_light,
                alignment=PP_ALIGN.CENTER,
                font_family=config.font_family,
            )


# ====================================================================
# SectionRenderer
# ====================================================================

class SectionRenderer(BaseRenderer):
    """セクション区切りレイアウトRenderer

    薄グレー背景にセクションタイトルを中央配置し、アクセント色の水平線で装飾する。
    """

    def render(self, slide, slide_data: SlideContent, config: DesignConfig):
        """セクション区切りスライドを描画"""
        # 背景: 薄グレー
        self.set_background_color(slide, config.bg_secondary)

        # セクションタイトル: 36pt, bold, テキスト色primary, 中央揃え
        self.add_text_box(
            slide,
            left=Inches(1.0),
            top=Inches(2.5),
            width=Inches(11.333),
            height=Inches(1.5),
            text=slide_data.title or "",
            font_size=Pt(36),
            font_color=config.text_primary,
            bold=True,
            alignment=PP_ALIGN.CENTER,
            font_family=config.font_family,
        )

        # アクセント色の水平線装飾
        self.add_horizontal_line(
            slide,
            left=Inches(5.667),
            top=Inches(4.2),
            width=Inches(2.0),
            color=config.accent_color,
            thickness=3.0,
        )


# ====================================================================
# TocRenderer
# ====================================================================

class TocRenderer(BaseRenderer):
    """目次・サマリーレイアウトRenderer

    番号付きの目次項目を表示する。番号部分はアクセント色で装飾する。
    """

    def render(self, slide, slide_data: SlideContent, config: DesignConfig):
        """目次スライドを描画"""
        # スライドタイトルとキーメッセージ
        if slide_data.title:
            self._render_slide_title(slide, slide_data.title, config)
        if slide_data.subtitle:
            self._render_key_message(slide, slide_data.subtitle, config)

        # 区切り線
        self._render_separator_line(slide, config)

        # 目次項目: 番号付きリスト
        items = slide_data.items or []
        if not items:
            return

        # 1つのテキストボックスに複数段落で配置
        txBox = slide.shapes.add_textbox(
            Inches(0.8), Inches(self.CONTENT_TOP),
            Inches(11.733), Inches(self.CONTENT_HEIGHT)
        )
        tf = txBox.text_frame
        tf.word_wrap = True

        for i, item in enumerate(items):
            # 最初の項目は既存のparagraphsを使用、2番目以降はadd_paragraph
            if i == 0:
                p = tf.paragraphs[0]
            else:
                p = tf.add_paragraph()

            # 番号Run（accent色, bold）
            num_run = p.add_run()
            num_run.text = f"{i + 1}.  "
            num_run.font.size = Pt(14)
            num_run.font.color.rgb = RGBColor.from_string(config.accent_color.lstrip("#"))
            num_run.font.bold = True
            num_run.font.name = config.font_family

            # テキストRun（text_primary色）
            text_run = p.add_run()
            text_run.text = item.get("title", "")
            text_run.font.size = Pt(14)
            text_run.font.color.rgb = RGBColor.from_string(config.text_primary.lstrip("#"))
            text_run.font.name = config.font_family

            p.line_spacing = 1.5
            p.alignment = PP_ALIGN.LEFT


# ====================================================================
# BulletPointsRenderer
# ====================================================================

class BulletPointsRenderer(BaseRenderer):
    """箇条書きレイアウトRenderer

    ビュレットマーカー付きの箇条書き項目を表示する。マーカーはアクセント色。
    """

    def render(self, slide, slide_data: SlideContent, config: DesignConfig):
        """箇条書きスライドを描画"""
        # スライドタイトルとキーメッセージ
        if slide_data.title:
            self._render_slide_title(slide, slide_data.title, config)
        if slide_data.subtitle:
            self._render_key_message(slide, slide_data.subtitle, config)

        # 区切り線
        self._render_separator_line(slide, config)

        # 箇条書き項目
        items = slide_data.items or []
        if not items:
            return

        # 1つのテキストボックスに複数段落で配置
        txBox = slide.shapes.add_textbox(
            Inches(self.CONTENT_LEFT), Inches(self.CONTENT_TOP),
            Inches(self.CONTENT_WIDTH), Inches(self.CONTENT_HEIGHT)
        )
        tf = txBox.text_frame
        tf.word_wrap = True

        for i, item in enumerate(items):
            if i == 0:
                p = tf.paragraphs[0]
            else:
                p = tf.add_paragraph()

            # ビュレットマーカーRun（accent色）
            marker_run = p.add_run()
            marker_run.text = "\u25cf  "  # ●
            marker_run.font.size = Pt(12)
            marker_run.font.color.rgb = RGBColor.from_string(config.accent_color.lstrip("#"))
            marker_run.font.name = config.font_family

            # テキストRun（text_primary色）
            text_run = p.add_run()
            text_run.text = item.get("title", "")
            text_run.font.size = Pt(12)
            text_run.font.color.rgb = RGBColor.from_string(config.text_primary.lstrip("#"))
            text_run.font.name = config.font_family

            p.line_spacing = 1.5
            p.alignment = PP_ALIGN.LEFT


# ====================================================================
# NumberedListRenderer
# ====================================================================

class NumberedListRenderer(BaseRenderer):
    """手順・ステップレイアウトRenderer

    円形ステップ番号（accent色）+ タイトル + 説明で各ステップを表示する。
    """

    def render(self, slide, slide_data: SlideContent, config: DesignConfig):
        """手順・ステップスライドを描画"""
        # スライドタイトルとキーメッセージ
        if slide_data.title:
            self._render_slide_title(slide, slide_data.title, config)
        if slide_data.subtitle:
            self._render_key_message(slide, slide_data.subtitle, config)

        # 区切り線
        self._render_separator_line(slide, config)

        # ステップ項目
        items = slide_data.items or []
        if not items:
            return

        # 項目数に応じてコンテンツ領域を等分
        num_items = len(items)
        step_interval = self.CONTENT_HEIGHT / max(num_items, 1)

        for i, item in enumerate(items):
            step_top = self.CONTENT_TOP + (i * step_interval)

            # ステップ番号の円（accent色）
            circle = self.add_shape(
                slide,
                MSO_SHAPE.OVAL,
                Inches(self.CONTENT_LEFT),
                Inches(step_top),
                Inches(0.45),
                Inches(0.45),
                fill_color=config.accent_color,
            )

            # 円内の番号テキスト（白色, 中央揃え）
            tf = circle.text_frame
            tf.word_wrap = False
            p = tf.paragraphs[0]
            p.text = str(i + 1)
            p.font.size = Pt(14)
            p.font.color.rgb = RGBColor.from_string(config.text_light.lstrip("#"))
            p.font.bold = True
            p.font.name = config.font_family
            p.alignment = PP_ALIGN.CENTER

            # ステップタイトル（14pt, bold）
            self.add_text_box(
                slide,
                left=Inches(1.1),
                top=Inches(step_top),
                width=Inches(11.733),
                height=Inches(0.35),
                text=item.get("title", ""),
                font_size=Pt(14),
                font_color=config.text_primary,
                bold=True,
                font_family=config.font_family,
                line_spacing=1.5,
            )

            # ステップ説明（12pt）
            description = item.get("description", "")
            if description:
                self.add_text_box(
                    slide,
                    left=Inches(1.1),
                    top=Inches(step_top + 0.35),
                    width=Inches(11.733),
                    height=Inches(0.4),
                    text=description,
                    font_size=Pt(12),
                    font_color=config.text_secondary,
                    font_family=config.font_family,
                    line_spacing=1.5,
                )


# ====================================================================
# CtaRenderer
# ====================================================================

class CtaRenderer(BaseRenderer):
    """行動喚起レイアウトRenderer

    薄グレー背景にメッセージ・説明・ボタン風テキストを中央配置する。
    """

    def render(self, slide, slide_data: SlideContent, config: DesignConfig):
        """CTAスライドを描画"""
        # 背景: 薄グレー
        self.set_background_color(slide, config.bg_secondary)

        # メッセージ: 24pt, bold, 中央揃え
        self.add_text_box(
            slide,
            left=Inches(1.5),
            top=Inches(1.8),
            width=Inches(10.333),
            height=Inches(0.8),
            text=slide_data.title or "",
            font_size=Pt(24),
            font_color=config.text_primary,
            bold=True,
            alignment=PP_ALIGN.CENTER,
            font_family=config.font_family,
        )

        # 説明: 16pt, テキスト色secondary, 中央揃え
        if slide_data.subtitle:
            self.add_text_box(
                slide,
                left=Inches(1.5),
                top=Inches(2.8),
                width=Inches(10.333),
                height=Inches(0.8),
                text=slide_data.subtitle,
                font_size=Pt(16),
                font_color=config.text_secondary,
                alignment=PP_ALIGN.CENTER,
                font_family=config.font_family,
            )

        # ボタン風テキスト: accent色背景の角丸矩形
        button_text = slide_data.cta_button or ""
        if button_text:
            # ボタン背景（角丸矩形）
            button_shape = self.add_shape(
                slide,
                MSO_SHAPE.ROUNDED_RECTANGLE,
                Inches(4.667),
                Inches(4.2),
                Inches(4.0),
                Inches(0.7),
                fill_color=config.accent_color,
            )
            # 角丸率を設定
            button_shape.adjustments[0] = 0.15

            # ボタンテキスト（図形内テキストフレーム）
            tf = button_shape.text_frame
            tf.word_wrap = True
            p = tf.paragraphs[0]
            p.text = button_text
            p.font.size = Pt(14)
            p.font.color.rgb = RGBColor.from_string(config.text_light.lstrip("#"))
            p.font.bold = True
            p.font.name = config.font_family
            p.alignment = PP_ALIGN.CENTER


# ====================================================================
# RENDERER_MAP
# ====================================================================

RENDERER_MAP = {
    "title": TitleRenderer,
    "section": SectionRenderer,
    "toc": TocRenderer,
    "bullet_points": BulletPointsRenderer,
    "numbered_list": NumberedListRenderer,
    "cta": CtaRenderer,
}


# ====================================================================
# テンプレート処理・プレゼンテーション生成
# ====================================================================

def create_presentation(config: DesignConfig, template_path: str = None) -> PptxPresentation:
    """テンプレートの有無に応じてPresentationオブジェクトを生成する

    Args:
        config: デザイン設定
        template_path: テンプレートPPTXファイルのパス

    Returns:
        python-pptxのPresentationオブジェクト
    """
    if template_path and os.path.exists(template_path):
        prs = PptxPresentation(template_path)
        # 既存スライドを全削除（逆順で削除）
        for i in range(len(prs.slides) - 1, -1, -1):
            rId = prs.slides._sldIdLst[i].rId
            prs.part.drop_rel(rId)
            del prs.slides._sldIdLst[i]
    else:
        prs = PptxPresentation()

    # スライドサイズを16:9に設定
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    return prs


def add_blank_slide(prs: PptxPresentation):
    """空白スライドを追加し、プレースホルダーを除去する

    Args:
        prs: python-pptxのPresentationオブジェクト

    Returns:
        追加されたスライドオブジェクト
    """
    slide_layout = prs.slide_layouts[6]  # 空白レイアウト（インデックス6）
    slide = prs.slides.add_slide(slide_layout)
    # プレースホルダーを除去
    for ph in list(slide.placeholders):
        sp = ph._element
        sp.getparent().remove(sp)
    return slide


def resolve_template_path(args_template: str, config: DesignConfig) -> str:
    """テンプレートPPTXのパスを優先順位に従って解決する

    優先順位: CLI --template > config.json template.pptx_path > なし

    Args:
        args_template: CLIで指定されたテンプレートパス
        config: デザイン設定

    Returns:
        テンプレートファイルのパス（なしの場合はNone）
    """
    if args_template:
        return args_template
    config_template = config.template_path
    if config_template and os.path.exists(config_template):
        return config_template
    return None


def generate_presentation(markdown_path: str, config: DesignConfig,
                          template_path: str = None, output_dir: str = "output",
                          title: str = None, layout_filter: str = None) -> str:
    """slides.mdからPPTXを生成するメイン関数

    Args:
        markdown_path: slides.mdファイルのパス
        config: デザイン設定
        template_path: テンプレートPPTXのパス
        output_dir: 出力ディレクトリ
        title: 出力ファイル名のプレフィックス
        layout_filter: 特定レイアウトのみ生成（デバッグ用）

    Returns:
        生成されたPPTXファイルのパス
    """
    # 1. slides.mdをパース
    with open(markdown_path, encoding="utf-8") as f:
        markdown_text = f.read()

    parser = SlidesMarkdownParser()
    presentation_data = parser.parse(markdown_text)

    # 2. Presentationオブジェクト生成
    prs = create_presentation(config, template_path)

    # 3. タイトルの決定
    output_title = title or presentation_data.title or "presentation"

    # 4. 各スライドを描画
    for slide_data in presentation_data.slides:
        layout = slide_data.layout or "bullet_points"
        if layout_filter and layout != layout_filter:
            continue
        renderer_cls = RENDERER_MAP.get(layout)
        if renderer_cls is None:
            print(f"警告: 未対応レイアウト '{layout}' をスキップします")
            continue
        slide = add_blank_slide(prs)
        renderer_cls().render(slide, slide_data, config)

    # 5. 保存
    os.makedirs(output_dir, exist_ok=True)
    output_path = os.path.join(output_dir, f"{output_title}.pptx")
    prs.save(output_path)
    return output_path


# ====================================================================
# CLI
# ====================================================================

def main():
    """CLIエントリポイント"""
    parser = argparse.ArgumentParser(
        description="slides.mdからPowerPointファイルを生成する"
    )
    parser.add_argument(
        "--markdown-file",
        type=str,
        required=True,
        help="slides.mdファイルのパス（必須）",
    )
    parser.add_argument(
        "--config",
        type=str,
        default=None,
        help="config.jsonファイルのパス（デフォルト: skills/slide-generator/config.json）",
    )
    parser.add_argument(
        "--title",
        type=str,
        default=None,
        help="出力ファイル名のプレフィックス（オプション）",
    )
    parser.add_argument(
        "--output-dir",
        type=str,
        default=None,
        help="出力ディレクトリ（デフォルト: config.jsonのoutput.dir）",
    )
    parser.add_argument(
        "--template",
        type=str,
        default=None,
        help="テンプレートPPTXのパス（オプション）",
    )
    parser.add_argument(
        "--layout",
        type=str,
        default=None,
        help="特定レイアウトのみを生成（デバッグ用）",
    )

    args = parser.parse_args()

    # config.jsonのパスを解決
    config_path = args.config
    if config_path is None:
        # デフォルト: スクリプトの親ディレクトリのconfig.json
        script_dir = Path(__file__).resolve().parent
        default_config = script_dir.parent / "config.json"
        if default_config.exists():
            config_path = str(default_config)

    # DesignConfigの読み込み
    config = DesignConfig(config_path)

    # 出力ディレクトリの決定
    output_dir = args.output_dir or config.output_dir

    # テンプレートパスの解決
    template_path = resolve_template_path(args.template, config)

    # PPTX生成
    output_path = generate_presentation(
        markdown_path=args.markdown_file,
        config=config,
        template_path=template_path,
        output_dir=output_dir,
        title=args.title,
        layout_filter=args.layout,
    )

    print(f"生成完了: {output_path}")


if __name__ == "__main__":
    main()
