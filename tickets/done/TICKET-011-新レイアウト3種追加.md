# TICKET-011: 新レイアウト3種追加（timeline, thank_you, matrix）

| 項目 | 内容 |
|------|------|
| チケット番号 | TICKET-011 |
| 作成日 | 2026-02-22 |
| 更新日 | 2026-02-22 |
| 担当者 | Claude |
| ブランチ | - |

## 概要

slide-generatorスキルのレイアウトを15種から18種に拡張する。timeline（ロードマップ・年表）、thank_you（終了スライド）、matrix（2x2マトリクス）の3レイアウトを追加する。

## 背景・目的

現在の15レイアウトではカバーできないプレゼンでよく使われるパターンがある。特にロードマップ表示、プレゼン末尾の締め、SWOT分析等の2軸分類は需要が高く、新レイアウトとして追加することでスキルの表現力を向上させる。

## 要件定義

- [ ] timelineレイアウトを追加する（横軸タイムライン線+マーカー+ラベル、テーブル記法、3-5項目推奨）
- [ ] thank_youレイアウトを追加する（紺背景+白文字+装飾線、title類似の構造）
- [ ] matrixレイアウトを追加する（2x2グリッド+十字区切り線、####で4象限定義）
- [ ] 各レイアウトのパーサー・シリアライザーをslides_markdown.pyに実装する
- [ ] 各レイアウトのRendererをslide_generator_pptx.pyに実装する
- [ ] 文字数制限を定義し、バリデーションに対応する
- [ ] ドキュメント（SKILL.md, layout-rules.md, character-limits.md, layout-selection-guide.md）を更新する
- [ ] E2Eサンプルに3レイアウトのテストスライドを追加する

## 設計

### 変更対象

| ファイル | 変更内容 |
|----------|----------|
| `scripts/slides_markdown.py` | VALID_LAYOUTS・CHARACTER_LIMITS追加、パースメソッド3つ、シリアライズメソッド3つ追加 |
| `scripts/slide_generator_pptx.py` | TimelineRenderer・ThankYouRenderer・MatrixRenderer追加、RENDERER_MAP登録 |
| `SKILL.md` | レイアウト一覧テーブル・slides.md記法例を15→18に更新 |
| `references/layout-rules.md` | 3レイアウトの記法例・仕様を追加 |
| `references/character-limits.md` | 3レイアウトの文字数制限を追加 |
| `references/layout-selection-guide.md` | フローチャート・選択基準に3レイアウト追加 |
| `tests/e2e_sample.md` | 3レイアウトのテストスライドを追加 |

### 設計方針

#### timeline

- **Markdown記法**: テーブル形式（`| 時期 | 内容 |`）
- **パース**: `_parse_timeline` メソッド追加。テーブルから `items: [{label, body}]` を生成
- **描画**: 横方向タイムライン線（グレー）+ アクセント色円形マーカー + 上に時期ラベル + 下に内容テキスト
- **背景**: 白
- **文字数制限**: タイトル25, キーメッセージ40, 時期ラベル10, 内容30

#### thank_you

- **Markdown記法**: `## メッセージ` + `### サブテキスト`（titleと同構造）
- **パース**: `_parse_thank_you` メソッド追加。h2/h3抽出のみ
- **描画**: 紺背景 + 36pt白文字メッセージ中央配置 + アクセント色装飾線 + 16ptサブテキスト
- **背景**: 紺色 `#1E3A5F`
- **文字数制限**: メッセージ25, サブテキスト50

#### matrix

- **Markdown記法**: `####` で4象限を定義（four_columnと同構造）
- **パース**: `_parse_matrix` メソッド追加。カラムパースと同一ロジック、レイアウト名のみ変更
- **描画**: 十字の区切り線で4象限に分割。各象限にヘッダ（プライマリ色）+ アクセント装飾線 + 本文
- **配置順**: 左上→右上→左下→右下
- **背景**: 白
- **文字数制限**: タイトル25, キーメッセージ40, 見出し8, 本文60

### 既存パターンの再利用

- BaseRendererの共通メソッド（`_render_slide_title`, `_render_key_message`, `_render_separator_line`, `add_text_box`, `add_shape`, `add_horizontal_line`, `add_vertical_line`）をそのまま利用
- RENDERER_MAPへの登録のみでレンダラーを有効化
- Pythonスクリプト以外の変更はドキュメント更新のみ

## 受け入れ条件

- [ ] slides_markdown.pyで3レイアウトのパース・シリアライズが正常に動作する
- [ ] slide_generator_pptx.pyで3レイアウトのPPTXが正常に生成される
- [ ] 文字数バリデーションが3レイアウトに対応している
- [ ] SKILL.md・layout-rules.md・character-limits.md・layout-selection-guide.mdが更新されている
- [ ] e2e_sample.mdに3レイアウトのテストスライドがあり、PPTX生成が成功する

## レビューメモ

<!-- レビュー時のコメント・指摘事項を記録 -->
