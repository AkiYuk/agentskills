# TICKET-002: スキル開発リポジトリ土台の構築

| 項目 | 内容 |
|------|------|
| チケット番号 | TICKET-002 |
| 作成日 | 2026-02-22 |
| 更新日 | 2026-02-22 |
| 担当者 | Claude |
| ブランチ | feature/TICKET-002-skills-foundation |

## 概要

Claude Code等で利用するスキルを複数開発・管理するためのディレクトリ構成とテンプレートを整備する。

## 背景・目的

agentskillsリポジトリはスキル開発を目的としているが、現状はCLAUDE.mdとチケットシステムのみでスキル関連の構成が未整備。Anthropic公式リポジトリ（anthropics/skills）の構成を参考に、複数スキルを効率的に管理できる土台を構築する。

## 要件定義

- [ ] `skills/` ディレクトリを作成し、各スキルを独立したフォルダで管理できる構成にする
- [ ] `skills/_template/SKILL.md` にAgent Skills標準準拠のテンプレートを作成する
- [ ] README.mdにプロジェクト概要・ディレクトリ構成・スキル追加手順を記載する

## 設計

### 変更対象

| ファイル | 変更内容 |
|----------|----------|
| `skills/_template/SKILL.md` | 新規作成。YAMLフロントマター付きスキルテンプレート |
| `README.md` | 更新。プロジェクト概要とスキル追加手順を追記 |

### 設計方針

- Anthropic公式リポジトリ（anthropics/skills）の構成に倣い、`skills/` 配下に各スキルをフォルダ単位で配置する
- テンプレートはチケットの `_template.md` と同じパターンで `skills/_template/` に配置する
- SKILL.mdはAgent Skills標準に準拠し、YAMLフロントマター（name, description）+ Markdown本文の構成とする
- README.mdにはディレクトリ構成図とスキルの追加手順を記載する

## 受け入れ条件

- [ ] `skills/_template/SKILL.md` が正しいYAMLフロントマター形式で存在する
- [ ] README.mdにプロジェクト概要・ディレクトリ構成・スキル追加手順が記載されている
- [ ] PRが作成され、mainにマージされている

## レビューメモ

<!-- レビュー時のコメント・指摘事項を記録 -->
