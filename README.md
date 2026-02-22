# agentskills

Claude Code等で使えるカスタムスキルを開発・管理するリポジトリ。

## ディレクトリ構成

```
skills/                       # スキル格納ディレクトリ
  _template/                  # 新規スキル作成用の雛形
  <スキル名>/                 # 各スキル（フォルダ単位で独立管理）
    SKILL.md                  # 必須: スキル定義（500行以下目安）
    REFERENCE.md              # 任意: 詳細仕様・技術リファレンス
    EXAMPLES.md               # 任意: 具体的な入出力例・ユースケース
    scripts/                  # 任意: ヘルパースクリプト（Python, Bash等）
    assets/                   # 任意: テンプレート・画像・データファイル

tickets/                      # チケット管理
  _template.md
  todo/ → in-progress/ → review/ → done/
```

## スキルの追加方法

1. `skills/_template/` フォルダをコピーして `skills/<スキル名>/` に配置する
2. `SKILL.md` の frontmatter（name, description）と本文を記入する
3. 必要に応じて補足ファイルを追加する（不要なテンプレートファイルは削除してよい）

## スキルファイルの形式

[Agent Skills標準](https://github.com/anthropics/skills) に準拠したYAMLフロントマター付きMarkdown。

```yaml
---
name: skill-name
description: What this skill does and when to use it.
---

# スキル名

手順・ガイドラインをMarkdownで記述
```

### ベストプラクティス

- `description` には「何をするか」と「いつ使うか」を3人称で書く
- `SKILL.md` は500行以下に保ち、詳細は参照ファイルに分離する
- 参照ファイルは `SKILL.md` から1階層までに留める
- 100行を超えるファイルにはトップに目次を入れる

## ライセンス

MIT
