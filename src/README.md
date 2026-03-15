# My SoftBank 料金明細PDF自動ダウンロード

My SoftBankにログインしてその月の請求書PDFを自動でダウンロードするスクリプト。

---

## セットアップ（初回のみ）

### 1. 依存ライブラリのインストール

```bash
pip3 install -r requirements.txt
playwright install chromium
```

### 2. GCPサービスアカウントの準備

スプレッドシートへのアクセスにサービスアカウント認証を使用する。

1. [Google Cloud Console](https://console.cloud.google.com/) でプロジェクトを作成
2. Google Sheets API を有効化
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

### 4. 保存先フォルダIDのマッピング

設定シートの `PDF保存先フォルダ` にGoogle DriveのURLを入力している場合、`drive_path_map.txt` にフォルダIDとMac上のローカルパスの対応を書く：

```
フォルダID=ローカルパス
1WEVnVT1KSgYAuAGMNAOy0VEL0VwwXplF=/Users/yamamoto/.../確定申告系/携帯領収書管理
```

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

ログイン時にSMSで3桁のセキュリティ番号が届く。スクリプトが待機状態になったら入力する：

```bash
echo '123' > /tmp/softbank_security_code.txt
```

（`123` の部分をSMSで届いた番号に変える）

入力を待機する時間は `SECURITY_CODE_TIMEOUT`（デフォルト300秒）で変更可能。

### セッションの再利用

認証成功後、セッション情報が `/tmp/softbank_session_{電話番号}.json` に保存される（アカウントごとに独立）。次回実行時にセッションが有効であれば、SMS認証がスキップされる。セッションは `/tmp/` に保存されるためMacを再起動すると消える。

---

## ファイル構成

```
携帯領収書管理/
├── src/
│   ├── mysoftbank_billing.py     # メインスクリプト
│   ├── setup_spreadsheet.gs      # スプレッドシート初期設定用GAS
│   ├── service_account.json      # GCPサービスアカウントキー（機密・gitignore対象）
│   ├── drive_path_map.txt        # Google DriveフォルダID → ローカルパス対応
│   ├── requirements.txt          # Python依存ライブラリ
│   ├── env.example               # 環境変数サンプル
│   ├── .env                      # 環境変数（機密・gitignore対象）
│   ├── README.md                 # このファイル
│   └── 仕様書.md                  # 技術仕様書
├── 2026/
│   └── 02/
│       └── SoftBank/
│           └── 202602_SoftBank_090XXXXXXXX_8500円.pdf
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
エラー時に `携帯領収書管理/debug/debug_{timestamp}/` に自動保存される。`HEADLESS=false` と組み合わせて原因を特定する。

**セキュリティ番号がタイムアウトした**
コードの有効期限切れの可能性がある。スクリプトを再起動してSMSを再送する。`SECURITY_CODE_TIMEOUT` を延ばしても根本解決にはならない（SMSコード自体の有効期限の問題）。
