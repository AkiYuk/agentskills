# TICKET-009: .skillビルドスクリプト追加

| 項目 | 内容 |
|------|------|
| チケット番号 | TICKET-009 |
| 作成日 | 2026-02-22 |
| 更新日 | 2026-02-22 |
| 担当者 | - |
| ブランチ | - |

## 概要

`skills/slide-generator/` をZIPにして `.skill` 拡張子でパッケージングするシェルスクリプトを追加する。

## 背景・目的

Claude Code Skillsとしてスキルを取り込むには、スキルディレクトリをZIPにして `.skill` 拡張子にする必要がある。手動でZIP化するのは手間なので、ビルドスクリプトとして用意する。

## 要件定義

- [ ] `build-skill.sh` を作成する
- [ ] `skills/slide-generator/` 配下を `__pycache__` と `.pyc` を除外してZIP化する
- [ ] 出力先は `output/slide-generator.skill`
- [ ] スクリプトは実行権限を付与する

## 設計

### 変更対象

| ファイル | 変更内容 |
|----------|----------|
| `build-skill.sh` | 新規作成。ZIPパッケージングスクリプト |

### 設計方針

- シンプルなシェルスクリプト（`zip` コマンド使用）
- `set -euo pipefail` でエラー時に即停止
- `__pycache__/` と `*.pyc` を除外
- 出力先ディレクトリ（`output/`）は自動作成

## 受け入れ条件

- [ ] `./build-skill.sh` を実行すると `output/slide-generator.skill` が生成される
- [ ] 生成された `.skill` ファイルがClaude Codeで取り込める

## レビューメモ

- 実装済みの `build-skill.sh` がステージングに残っている（mainへの直接コミットを取り消し済み）
