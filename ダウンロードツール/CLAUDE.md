# CLAUDE.md - 携帯領収書管理プロジェクト

## プロジェクト概要

携帯キャリア（SoftBank / Y!mobile / au / UQ mobile）の料金明細PDFを自動ダウンロードしてGoogle Driveに保存するツール集。

## アーキテクチャ

- **2スプレッドシート構成**: 認証情報管理シート（外部）+ 携帯領収書管理スプレッドシート（メイン）
- **認証情報シートは6列**: 電話番号 | キャリア | PDFの種類 | 運用端末 | 状態 | ログインID（パスワード列なし、設定シートに一元管理）
- **HTMLサイドバー方式**: GASメニューからサイドバーを開いて電話番号を選択
- **共通モジュール方式**: `shared_utils.py` にすべてのロジックを集約。各キャリアスクリプトは `CarrierConfig`（定数）+ `create_billing_context()` + `run_main()` を呼ぶだけの薄いラッパー（約80行）
- **carrier_family分岐**: `CarrierConfig.carrier_family` で `"softbank"`（SoftBank/Ymobile: WCOシステム）と `"au"`（au/UQmobile: WEB de 請求書）のフローを切り替え

## 用語ルール

- 「アカウント」ではなく「回線」を使用すること（コード・ドキュメント共通）
- Google関連は「Googleアカウント」「サービスアカウント」のままでOK

## コーディングルール

- **Pythonの共通ロジックは必ず `shared_utils.py` に書く**。キャリアスクリプトにロジックを書かない
- キャリア固有の値は `CarrierConfig` の定数として定義する
- 実行時の状態は `BillingContext` で管理する（モジュールレベルのグローバル変数は使わない）
- GASメニュー名は「携帯領収書管理 ツール」で統一
- 解約済回線はグレーアウト表示（削除しない）
- リンクシートの既存データは保存時に削除しない
- ダウンロード履歴にはファイル名を含める

## ドキュメント更新ルール

- コード変更時は必ず関連ドキュメント（README.md、仕様書.md）も更新すること
- README.md、ダウンロードツール/仕様書.md が対象
- 仕様変更（シート構成・メニュー名・機能追加削除等）があった場合、コードだけでなく仕様書.mdとREADME.mdの該当箇所も同時に修正すること
- 「コード修正 → ドキュメント更新 → コミット」の順で1つのコミットにまとめること

## Git運用ルール

- 作業完了時は必ずコミット漏れがないか `git status` で確認すること
- コード変更とドキュメント更新は同じコミットに含めること
- コミットメッセージは `feat:` / `fix:` / `docs:` / `refactor:` / `chore:` のプレフィックスを使用
- コミットメッセージは日本語で記述
- **コミット後は必ず `git push` してGitHubを最新状態に保つこと**

## ファイル構成

```
携帯領収書管理/
├── 2026/                          # 領収書PDF（年別フォルダ）
├── README.md                      # プロジェクト全体のREADME
├── ダウンロードツール/              # スクリプト・設定一式
│   ├── download.py                # 統合エントリポイント（--update-amounts で金額更新）
│   ├── update_amounts.py          # 金額更新の個別スクリプト（download.pyのラッパー）
│   ├── shared_utils.py            # 全ロジック集約
│   ├── setup_spreadsheet.gs       # GASコード（リンクシート更新のみ）
│   ├── CLAUDE.md                  # このファイル
│   ├── 仕様書.md                  # 詳細仕様
│   ├── requirements.txt
│   ├── env.example
│   ├── service_account.json       # 機密・gitignore対象
│   └── bin/                       # CLAUDE.md自動更新フック
└── 認証情報管理シート.gsheet       # スプレッドシートショートカット
```
