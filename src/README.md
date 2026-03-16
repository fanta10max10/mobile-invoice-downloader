# My SoftBank 料金明細PDF自動ダウンロード

My SoftBankにログインしてその月の請求書PDFを自動でダウンロードし、Google Driveに直接アップロードするスクリプト。

---

## セットアップ（初回のみ）

### 1. 依存ライブラリのインストール

```bash
# 仮想環境を作成（推奨）
python3 -m venv ../.venv
source ../.venv/bin/activate  # Windows: ..\\.venv\\Scripts\\activate

pip install -r requirements.txt
playwright install chromium
```

### 2. GCPサービスアカウントの準備

スプレッドシートとDrive APIへのアクセスにサービスアカウント認証を使用する。

1. [Google Cloud Console](https://console.cloud.google.com/) でプロジェクトを作成
2. Google Sheets API と Google Drive API を有効化
3. サービスアカウントを作成してJSONキーをダウンロード
4. ダウンロードしたJSONを `src/service_account.json` として配置
5. スプレッドシートをそのサービスアカウントのメールアドレスに「閲覧者」として共有

### 3. Googleスプレッドシートの準備

`setup_spreadsheet.gs` をGoogleスプレッドシートのApps Scriptに貼り付けて `setupSheet()` を実行すると、アカウント管理用のシートが自動作成される。

**アカウントシートに入力する内容：**

| 電話番号 | パスワード | PDFの種類 |
|---|---|---|
| 090-XXXX-XXXX | パスワード | 電話番号別 |

**設定シートに入力する内容：**

| 設定名 | 値 |
|---|---|
| PDF保存先フォルダ | `https://drive.google.com/drive/folders/XXXXXXXX` |
| 対象月 | ドロップダウンで選択（「自動（前月）」= 前月を自動選択） |

### 4. PDF保存先フォルダをサービスアカウントと共有

設定シートに入力したGoogle Driveフォルダをサービスアカウントのメールアドレスと共有する：

1. `service_account.json` 内の `client_email` を確認
2. Google Driveで対象フォルダを開き「共有」→ そのメールアドレスを追加（権限: **編集者**）

### 5. 環境変数の設定

```bash
cp env.example .env
```

`.env` を開いて `SPREADSHEET_URL` にスプレッドシートのURLを貼り付ける（ブラウザのURLをそのままコピーでOK）。

---

## 実行方法

```bash
# src/ フォルダ内で実行する
cd src/

# 前月分をダウンロード（通常はこれだけ）
python3 mysoftbank_billing.py

# 特定の月を指定（環境変数は設定シートより優先）
TARGET_MONTH=202602 python3 mysoftbank_billing.py

# ブラウザを表示してデバッグ
HEADLESS=false python3 mysoftbank_billing.py
```

対象月はスプレッドシートの設定シートからも選択できる（ドロップダウン）。

---

## セキュリティ番号（SMS）の入力

ログイン時にSMSで3桁のセキュリティ番号が届く。スクリプトが待機状態になったらそのままターミナルに入力してEnterを押す。

入力を待機する時間は `SECURITY_CODE_TIMEOUT`（デフォルト300秒）で変更可能。

### セッションの再利用

認証成功後、セッション情報がOSの一時フォルダ（macOS: `/tmp/`、Windows: `%TEMP%`）にアカウントごとに保存される。次回実行時にセッションが有効であれば、SMS認証がスキップされる。OSを再起動すると一時フォルダが消えてセッションもリセットされる。

---

## ファイル構成

```
携帯領収書管理/
├── .venv/                        # Python仮想環境（gitignore対象）
├── src/
│   ├── mysoftbank_billing.py     # メインスクリプト
│   ├── setup_spreadsheet.gs      # スプレッドシート初期設定用GAS
│   ├── service_account.json      # GCPサービスアカウントキー（機密・gitignore対象）
│   ├── requirements.txt          # Python依存ライブラリ
│   ├── env.example               # 環境変数サンプル
│   ├── .env                      # 環境変数（機密・gitignore対象）
│   ├── README.md                 # このファイル
│   └── 仕様書.md                  # 技術仕様書
└── debug/
    └── debug_20260315_203000/
        ├── 090XXXXXXXX_202602_エラー種別.png
        └── 090XXXXXXXX_202602_エラー種別.txt
```

---

## 環境変数一覧

| 変数名 | デフォルト | 説明 |
|---|---|---|
| `SPREADSHEET_URL` | （なし） | スプレッドシートのURL（ブラウザのURLをそのまま貼り付け） |
| `SPREADSHEET_ID` | （なし） | IDのみ指定する場合（SPREADSHEET_URLが優先） |
| `BASE_SAVE_PATH` | （なし） | PDF保存先（設定シートで管理するため通常不要） |
| `TARGET_MONTH` | 前月自動 | ダウンロード対象月（YYYYMM形式。設定シートより優先） |
| `HEADLESS` | `true` | ブラウザ非表示モード。`false` でブラウザを表示 |
| `SECURITY_CODE_TIMEOUT` | `300` | SMSセキュリティ番号の入力待機時間（秒） |

---

## トラブルシューティング

**デバッグスクリーンショット**
エラー時に `src/debug/debug_{timestamp}/` に自動保存される。`HEADLESS=false` と組み合わせて原因を特定する。

**セキュリティ番号がタイムアウトした**
コードの有効期限切れの可能性がある。スクリプトを再起動してSMSを再送する。`SECURITY_CODE_TIMEOUT` を延ばしても根本解決にはならない（SMSコード自体の有効期限の問題）。

**ファイル名が `_利用料金明細.pdf` のままで金額が入らない**
GAS の「PDFから金額を取得・ファイル名更新」を実行する。Drive APIへのアップロード直後は反映に数秒かかる場合があるのでしばらく待ってから再実行する。

**「PDFから金額を取得・ファイル名更新」で `Drive is not defined` エラー**
GAS エディタで「サービス」→「Drive API」を追加していない。Apps Script の「サービスを追加」から Drive API (v2) を有効にする。

**OCRで金額が取得できない（「金額取得失敗」件数が増える）**
GAS のログ（表示 → ログ）に出力されたOCRテキストを確認し、SoftBankの請求書レイアウト変更で正規表現パターンが合わなくなっていないか確認する。
