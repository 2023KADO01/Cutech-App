# CLAUDE.md

このファイルは、本リポジトリで作業する Claude Code 向けのガイドです。

## プロジェクト概要

ルービックキューブサークル「Cutech」の記録アプリです。2つのWebアプリで構成されます。

| 略称 | アプリ名 | 役割 |
| --- | --- | --- |
| L | Cutech Log | 記録の入力用 |
| R | Cutech Records | 記録の閲覧用 |

- どちらも Google Apps Script (GAS) で作る。
- どちらも Google スプレッドシートに紐づけ、データの保存・読込を行う。
- スプレッドシートID、シート名、列構成などの具体値は、ユーザーに確認するか既存コードから読み取ること。推測で書かない。

## リポジトリ構成

```
log/       # Cutech Log (L): 記録入力アプリ（L_Code.gs, L_Index.html, .clasp.json）
records/   # Cutech Records (R): 記録閲覧アプリ（R_Code.gs, R_Index.html, .clasp.json）
.github/workflows/deploy.yml   # main へのマージで Apps Script に自動で push・デプロイする
.github/scripts/deploy-gas.sh  # 上のワークフローが使うスクリプト
.github/workflows/update-cubing.yml  # cubing.js（スクランブル生成）の配布元を毎週確認し、変わっていれば更新PRを作る
.github/scripts/fetch-cubing.mjs     # 上のワークフローが使うスクリプト
vendor/cubing/                 # cubing.js の固定コピー（README.md 参照）
assets/cutechicon.png          # ファビコン（L・R の setFaviconUrl が参照）
```

- appsscript.json はリポジトリで管理しない（デプロイ時に Apps Script 上のものを取ってきて使う）。
- R のスプレッドシートIDは、R のスクリプト プロパティ SHEET_ID に設定する（コードには書かない）。

## 作業分担ルール（重要）

**目的はトークンの節約。以下を必ず守ること。**

- **決まった作業はサブエージェント（`model: sonnet`）に任せる。**
  - 対象: 実装、レビュー、調査、定型的な編集など。
- **Opus が行うのは、方針決定・設計判断などの重要な判断だけ。**
  - 上記の作業を Opus が自分で抱え込まないこと。
- サブエージェントを起動するときは、`model: sonnet` を明示的に指定する。
- サブエージェントへの指示には、目的・対象ファイル・完了条件を簡潔に書く。

## UI変更のルール

- UIを変更するときは、PRを作る前に、変更した部分をすべてスクリーンショットに撮ってユーザーに見せる。
  - Playwright で google.script.run をスタブしてページを開き、PC幅とスマホ幅（390px）の両方で撮る。
  - ユーザーの了承を得てからPRを作る。
  - スクショの漢字が中国語フォントにならないよう、日本語フォント（IPAゴシックや Noto Sans CJK JP など）で描画されることを確かめる。コンテナでは fontconfig で日本語フォントを優先し、WenQuanYi などの中国語フォントを外す。
- 格言：**良いUIとは、文字の説明ばかりに頼らずUI自身で説明させるものである。** 説明文や注記を足す前に、配置・強調・アイコン・状態の見せ方で伝えられないかを考える。

## コミュニケーション

- ユーザーとの会話は**日本語**で行う。
- 成果物（ドキュメント、コメント、コミットメッセージ等の文章）も日本語で書く。
  - コード上の識別子は、通常どおり英語でよい。
