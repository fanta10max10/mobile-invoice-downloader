# 携帯領収書管理

携帯キャリアの料金明細PDFを自動ダウンロードしてGoogle Driveに保存するツール集。

---

## 概要

My SoftBankおよびMy Y!mobileにPlaywrightで自動ログインし、指定月の請求書PDFをダウンロードしてGoogle Drive APIで直接アップロードするPythonスクリプト群。

- 複数アカウントの一括処理
- SMS認証（セッション再利用でスキップ可能）
- Google Drive APIによるPDF直接アップロード（OAuth2優先、サービスアカウントフォールバック）
- 認証情報管理シートによる解約済アカウントの自動スキップ
- 電子帳簿保存法準拠のファイル命名

---

## 対応キャリア

| フォルダ | キャリア | スクリプト | PDFの種類 |
|---|---|---|---|
| [SoftBank/](SoftBank/) | SoftBank | `mysoftbank_billing.py` | 電話番号別 / 一括 / 機種別 |
| [Ymobile/](Ymobile/) | Y!mobile | `myyumobile_billing.py` | 電話番号別のみ |

---

## セットアップ

### 1. Python仮想環境

```bash
# このフォルダ（携帯領収書管理/）で実行
python3 -m venv .venv
source .venv/bin/activate  # Windows: .venv\Scripts\activate
```

各キャリアフォルダで以下を実行する：

```bash
pip install -r requirements.txt
playwright install chromium
```

### 2. GCPプロジェクトの準備

スプレッドシートとDrive APIへのアクセスにGoogle Cloud Platformの認証を使用する。

1. [Google Cloud Console](https://console.cloud.google.com/) でプロジェクトを作成
2. Google Sheets API と Google Drive API を有効化
3. 以下のいずれかの認証方法を選択する（OAuth2推奨）

### 3. 認証ファイルの配置

#### OAuth2認証（推奨）

ユーザー自身のGoogleアカウントでDriveにアップロードする。サービスアカウントへのフォルダ共有設定が不要で、`storageQuotaExceeded` エラーも発生しない。

1. Google Cloud Console で「OAuth 2.0 クライアント ID」を作成（種類: デスクトップアプリ）
2. JSONをダウンロードして `SoftBank/client_secrets.json` として配置
3. 初回実行時にブラウザが開くのでGoogleアカウントで許可する
4. `SoftBank/drive_oauth_token.json` にトークンが自動保存される（次回以降不要）

> `client_secrets.json` をPDF保存先Driveフォルダにアップロードしておくと、スクリプトが自動でダウンロードして配置する。

#### サービスアカウント認証（フォールバック）

`client_secrets.json` が存在しない場合に自動的に使用される。

1. サービスアカウントを作成してJSONキーを `SoftBank/service_account.json` として配置
2. スプレッドシートをサービスアカウントの `client_email` に「閲覧者」として共有する
3. PDF保存先のDriveフォルダをサービスアカウントの `client_email` に「編集者」として共有する

#### 機密ファイルの一元管理

`service_account.json` と `client_secrets.json` は **`SoftBank/` フォルダに1つ置くだけでよい**。YmobileスクリプトはSoftBankフォルダを自動検索して参照する。

### 4. スプレッドシートの準備

2つのスプレッドシートを使用する:
- **認証情報管理シート** — 設定・認証情報・PDFリンク管理（GASを配置する側）
- **回線管理スプレッドシート** — 月別の電話番号・解約済・運用端末等のデータ（別スプレッドシート）

ルートフォルダの `setup_spreadsheet.gs` を認証情報管理シートのApps Scriptに貼り付けて以下の手順を実行する：

1. `setupSheet()` を実行 → 設定・認証情報・SoftBankリンク・Ymobileリンク シートを自動作成
2. 設定シートの「回線管理スプレッドシート」に回線管理表のURLを入力
3. 設定シートの「パスワード」にログインパスワードを入力（全番号共通）
4. メニュー「ダウンロード対象の電話番号を管理」でサイドバーを開き対象を選択

#### シート構成（認証情報管理シート）

| シート名 | 役割 |
|---|---|
| 設定 | PDF保存先・パスワード・対象月・回線管理スプシURL |
| 認証情報 | ダウンロード対象の電話番号・キャリア・PDFの種類・運用端末 |
| SoftBankリンク | SoftBank PDF月別リンク（GAS自動更新） |
| Ymobileリンク | Y!mobile PDF月別リンク（GAS自動更新） |

#### 認証情報シート（ダウンロード対象の電話番号）

メニューから「ダウンロード対象の電話番号を管理」でHTMLサイドバーを開くと、回線管理スプレッドシートの月別シートから電話番号を動的に読み込み、ダウンロード対象をチェックボックスで選択できる。

| 電話番号 | キャリア | PDFの種類 | 運用端末 | 状態 |
|---|---|---|---|---|
| 09012345678 | SoftBank | 電話番号別 | iPhoneAir | |
| 08012345678 | Ymobile | 電話番号別 | iPhone16 | |
| 090XXXXXXXX | SoftBank | | | 解約済 |

- パスワードは設定シートで一元管理（認証情報シートにはパスワード列なし）
- `キャリア` 列で SoftBank / Ymobile スクリプトが自動フィルタリング
- SoftBankは `電話番号別` / `一括` / `機種別` をカンマ区切りで複数指定可
- 運用端末・状態はサイドバー保存時に回線管理スプレッドシートから自動設定

#### サイドバーの「保存」で行われること

1. **認証情報シート**: 全データをクリアして再書き込み（選択中の番号 + 解約済をグレー表示）
2. **リンクシート**: 新しい電話番号を追加（既存行やPDFリンクは削除しない）

#### 設定シート

| 設定名 | 値 |
|---|---|
| このスプレッドシートURL | （自動設定） |
| 回線管理スプレッドシート | 回線管理表のURL |
| PDF保存先フォルダ | `https://drive.google.com/drive/folders/XXXXX` |
| パスワード | SoftBank / Y!mobile 共通のログインパスワード |
| 対象月 | ドロップダウン（「自動（前月）」= 前月自動） |

### 5. 環境変数（.env）

`.env` ファイルはスプレッドシートの `.gsheet` ファイルから自動生成される。自動生成されない場合は手動で作成する：

```bash
cp env.example .env
```

`.env` を開いて `SPREADSHEET_URL` にスプレッドシートのURLを貼り付ける（ブラウザのURLをそのままコピーでOK）。

---

## 実行方法

### SoftBank

```bash
cd SoftBank/

# 前月分をダウンロード（通常はこれだけ）
python3 mysoftbank_billing.py

# 特定の月を指定（環境変数は設定シートより優先）
TARGET_MONTH=202602 python3 mysoftbank_billing.py

# ブラウザを表示してデバッグ
HEADLESS=false python3 mysoftbank_billing.py
```

### Y!mobile

```bash
cd Ymobile/

# 前月分をダウンロード（通常はこれだけ）
python3 myyumobile_billing.py

# 特定の月を指定
TARGET_MONTH=202602 python3 myyumobile_billing.py

# ブラウザを表示してデバッグ
HEADLESS=false python3 myyumobile_billing.py
```

対象月はスプレッドシートの設定シートからも選択できる（ドロップダウン）。環境変数 `TARGET_MONTH` は設定シートより優先される。

---

## SMS認証について

ログイン時にSMSで3桁のセキュリティ番号が届く場合、スクリプトが待機状態になり以下のように端末情報とともに表示される：

```
============================================================
  📱 SMS認証が必要です
  電話番号: 09047695015
  端末    : iPhoneAir
  SMSに届いた3桁のセキュリティ番号を入力してください
  ターミナル入力 または 以下のコマンドで渡してください:
    echo '123' > /tmp/softbank_security_code.txt
============================================================
```

- **ターミナル直接実行時**: そのまま3桁を入力してEnterを押す
- **バックグラウンド実行時**: ファイル経由でコードを渡す
  - Mac/Linux: `echo '854' > /tmp/softbank_security_code.txt`
  - Windows: `echo 854 > %TEMP%\softbank_security_code.txt`
- 入力待機時間は `SECURITY_CODE_TIMEOUT`（デフォルト300秒）で変更可能

### セッションの再利用

認証成功後、セッション情報がOSの一時フォルダにアカウントごとに保存される。次回実行時にセッションが有効であれば、SMS認証がスキップされる。

| OS | セッションファイルの場所 |
|---|---|
| Mac/Linux | `/tmp/softbank_session_{電話番号}.json` |
| Windows | `%TEMP%\softbank_session_{電話番号}.json` |

認証エラーが解消しない場合はセッションファイルを削除して再実行する：

```bash
# Mac/Linux
rm /tmp/softbank_session_*.json

# Windows
del %TEMP%\softbank_session_*.json
```

---

## 環境変数一覧

| 変数名 | デフォルト | 説明 |
|---|---|---|
| `SPREADSHEET_URL` | （自動生成） | 認証情報管理シートのURL（`.gsheet` ファイルから自動生成） |
| `SPREADSHEET_ID` | （なし） | IDのみ指定する場合（SPREADSHEET_URLが優先） |
| `BASE_SAVE_PATH` | （なし） | PDF保存先（設定シートで管理するため通常不要） |
| `TARGET_MONTH` | 前月自動 | ダウンロード対象月（YYYYMM形式。設定シートより優先） |
| `HEADLESS` | `true` | ブラウザ非表示モード。`false` でブラウザを表示 |
| `SECURITY_CODE_TIMEOUT` | `300` | SMSセキュリティ番号の入力待機時間（秒） |
| `DRY_RUN` | `false` | `true` で設定確認のみ（ログインやダウンロードは行わない） |
| `RETRY_PHONES` | （なし） | 特定番号のみ再実行（カンマ区切り。例: `09012345678,08012345678`） |

---

## フォルダ構成

```
携帯領収書管理/
├── .venv/                          # Python仮想環境（gitignore対象）
├── .gitignore
├── README.md                       # このファイル
├── setup_spreadsheet.gs            # 携帯領収書管理スプシ用GAS（統合版）
├── SoftBank/
│   ├── mysoftbank_billing.py       # SoftBankダウンロードスクリプト
│   ├── service_account.json        # GCPサービスアカウントキー（機密・gitignore対象）
│   ├── client_secrets.json         # OAuth2クライアントシークレット（機密・gitignore対象）
│   ├── drive_oauth_token.json      # OAuth2トークン（自動生成・機密・gitignore対象）
│   ├── requirements.txt
│   ├── env.example
│   ├── .env                        # 自動生成・gitignore対象
│   ├── README.md
│   └── 仕様書.md
├── Ymobile/
│   ├── myyumobile_billing.py       # Y!mobileダウンロードスクリプト
│   ├── requirements.txt
│   ├── env.example
│   ├── .env                        # 自動生成・gitignore対象
│   └── README.md
└── {YYYY}/
    └── {MM}/
        └── {キャリア}/
            └── *.pdf
```

---

## 機密ファイル一覧

| ファイル | 配置場所 | 生成方法 | 備考 |
|---|---|---|---|
| `service_account.json` | `SoftBank/` | 手動配置 | Ymobileからも自動検索で参照される |
| `client_secrets.json` | `SoftBank/` | 手動配置 | Ymobileからも自動検索で参照される。PDF保存先フォルダに置けば自動ダウンロード |
| `drive_oauth_token.json` | `SoftBank/` | 自動生成 | 初回OAuth2認証後に生成される |
| `.env` | 各キャリアフォルダ | 自動生成 | `.gsheet` ファイルから `SPREADSHEET_URL` を自動生成 |

---

## トラブルシューティング

### SoftBank / Y!mobile 共通

**デバッグスクリーンショット**
エラー時にキャリアフォルダ内の `debug/debug_{timestamp}/` に自動保存される。`HEADLESS=false` と組み合わせて原因を特定する。

**セキュリティ番号がタイムアウトした**
SMSコードの有効期限切れの可能性がある。スクリプトを再起動してSMSを再送する。`SECURITY_CODE_TIMEOUT` を延ばしても根本解決にはならない（SMSコード自体の有効期限の問題）。

**スプレッドシートへのアクセスで403エラーが出る**
スクリプトがブラウザを自動起動して共有用メールアドレスをクリップボードにコピーする。そのメールアドレスをスプレッドシートの共有設定に追加する。

**`client_secrets.json` が見つからない**
PDF保存先フォルダに `client_secrets.json` を置いておくと自動でダウンロードして配置する。または手動で `SoftBank/client_secrets.json` に配置する（Ymobileからも自動参照される）。

**ファイル名が `_利用料金明細.pdf` のままで金額が入らない**
GAS の「PDFから金額を取得・ファイル名更新」を実行する。Drive APIへのアップロード直後は反映に数秒かかる場合があるのでしばらく待ってから再実行する。

**「PDFから金額を取得・ファイル名更新」で `Drive is not defined` エラー**
GAS エディタで「サービス」→「Drive API」を追加していない。Apps Script の「サービスを追加」から Drive API (v2) を有効にする。

**OCRで金額が取得できない（「金額取得失敗」件数が増える）**
GAS のログ（表示 → ログ）に出力されたOCRテキストを確認し、キャリアの請求書レイアウト変更で正規表現パターンが合わなくなっていないか確認する。

**Drive APIでストレージ超過エラー（storageQuotaExceeded）**
サービスアカウント認証時に発生する。OAuth2認証（`client_secrets.json`）に切り替えると、ユーザー自身のDriveにアップロードされるため解消する。

### SoftBank固有

**SMS認証で端末名が表示されない**
認証情報シートの「運用端末」列が空欄。サイドバーで保存し直すと回線管理表から自動設定される。

### Y!mobile固有

**ログインが通らない**
Y!mobileはSoftBank ID（id.my.ymobile.jp）と共通の認証基盤を使用している。パスワードはMy SoftBankと同じIDで管理されている場合がある。

---

## 技術仕様

詳細な技術仕様（処理フロー・認証フロー・スプレッドシート構成・環境変数など）は [SoftBank/仕様書.md](SoftBank/仕様書.md) を参照。
