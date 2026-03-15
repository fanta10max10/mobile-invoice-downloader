# My SoftBank 料金明細PDF自動ダウンロード

My SoftBankにログインしてその月の請求書PDFを自動でダウンロードするスクリプト。

---

## セットアップ（初回のみ）

### 1. 依存ライブラリのインストール

```bash
pip3 install -r requirements.txt
playwright install chromium
```

### 2. Googleスプレッドシートの準備

`setup_spreadsheet.gs` をGoogleスプレッドシートのApps Scriptに貼り付けて `setupSheet()` を実行すると、アカウント管理用のシートが自動作成される。

シートに入力する内容：

| phone_number | password | PDF保存先 |
|---|---|---|
| 090-XXXX-XXXX | パスワード | https://drive.google.com/drive/folders/XXXXXXXX |

- `phone_number`: SoftBank IDの携帯電話番号（ハイフンあり/なしどちらでもOK）
- `password`: My SoftBankのパスワード
- `PDF保存先`: Google DriveのフォルダURL（`https://drive.google.com/drive/folders/...` 形式）
  - ローカルの絶対パス（`/Users/...`）でも可

その後、スプレッドシートの **ファイル → 共有 → ウェブに公開 → カンマ区切り形式（.csv）** で公開し、URLを `MySoftBank_アカウント管理スプシURL.rtf` に貼り付けて保存しておく。

### 3. 保存先フォルダIDのマッピング

`PDF保存先` にGoogle DriveのフォルダURLを書いている場合（現在の運用）、`drive_path_map.txt` にフォルダID（URLの `folders/` 以降の文字列）とMac上のローカルパスの対応を書く：

```
フォルダID=ローカルパス
1WEVnVT1KSgYAuAGMNAOy0VEL0VwwXplF=/Users/yamamoto/.../確定申告系/携帯領収書管理
```

---

## 実行方法

```bash
# src/ フォルダ内で実行する
cd src/

# 前月分をダウンロード（通常はこれだけ）
python3 mysoftbank_billing.py

# 特定の月を指定
TARGET_MONTH=202602 python3 mysoftbank_billing.py

# ブラウザを表示してデバッグ
HEADLESS=false python3 mysoftbank_billing.py
```

---

## セキュリティ番号（SMS）の入力

ログイン時にSMSで3桁のセキュリティ番号が届く。スクリプトが待機状態になったら、**別のターミナルで** 以下を実行する：

```bash
echo '123' > /tmp/softbank_security_code.txt
```

（`123` の部分をSMSで届いた番号に変える）

入力を待機する時間は `SECURITY_CODE_TIMEOUT`（デフォルト300秒）で変更可能。

### セッションの再利用

認証成功後、セッション情報が `/tmp/softbank_session.json` に保存される。次回実行時にセッションが有効であれば、SMS認証がスキップされる。セッションは `/tmp/` に保存されるためMacを再起動すると消える。

---

## ファイル構成

```
softbank-invoice-downloader/
├── mysoftbank_billing.py              # メインスクリプト
├── MySoftBank_アカウント管理スプシURL.rtf  # スプレッドシートのCSV URL（機密・gitignore対象）
├── drive_path_map.txt                 # Google DriveフォルダID → ローカルパス対応
├── requirements.txt                   # Python依存ライブラリ
├── setup_spreadsheet.gs               # スプレッドシート初期設定用Google Apps Script
├── env.example                        # 環境変数サンプル
├── .env                               # 環境変数（機密・gitignore対象）
├── .gitignore                         # GitHub非公開ファイルの除外設定
├── README.md                          # このファイル
└── 仕様書.md                           # 技術仕様書
```

### フォルダ構成（携帯領収書管理/）

```
携帯領収書管理/
├── 2026/
│   └── 02/
│       └── 202602_SoftBank_09048469405_利用料金明細.pdf
└── debug/
    └── debug_20260315_203000/
        ├── 09048469405_202602_エラー種別.png
        └── 09048469405_202602_エラー種別.txt
```

---

## 環境変数一覧

| 変数名 | デフォルト | 説明 |
|---|---|---|
| `SPREADSHEET_CSV_URL` | （なし） | スプレッドシートのCSV URL。未設定時はRTFファイルから自動取得 |
| `BASE_SAVE_PATH` | （なし） | PDF保存先。未設定時はスプレッドシートの`PDF保存先`列を使用 |
| `TARGET_MONTH` | 前月 | ダウンロード対象月（YYYYMM形式） |
| `HEADLESS` | `true` | ブラウザ非表示モード。`false`でブラウザを表示 |
| `SECURITY_CODE_TIMEOUT` | `300` | SMSセキュリティ番号の入力待機時間（秒） |

---

## トラブルシューティング

**スプレッドシートの更新が反映されない**
Google Sheetsの「ウェブに公開」CSVはキャッシュがある。変更後数分待つか、公開を一度停止して再公開する。

**デバッグスクリーンショット**
エラー時に `携帯領収書管理/debug/debug_{timestamp}/` に自動保存される。`HEADLESS=false` と組み合わせて原因を特定する。

**セキュリティ番号がタイムアウトした**
コードの有効期限切れの可能性がある。スクリプトを再起動してSMSを再送してもらう。`SECURITY_CODE_TIMEOUT` を延ばしても根本解決にはならない（SMSコード自体の有効期限の問題）。
