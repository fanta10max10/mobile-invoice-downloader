# 携帯領収書管理

携帯キャリアの料金明細PDFを自動ダウンロードしてGoogle Driveに保存するツール集。

---

## 概要

My SoftBank / My Y!mobile / My au / My UQ mobile / My docomo にPlaywrightで自動ログインし、指定月の請求書PDFをダウンロードしてGoogle Drive APIで直接アップロードするPythonスクリプト群。

- 複数回線の一括処理
- SMS認証（セッション再利用でスキップ可能）
- Google Drive APIによるPDF直接アップロード（OAuth2優先、サービスアカウントフォールバック）
- 解約済回線のダウンロード対応（グレーアウト表示、SoftBank ID / au IDでログイン）
- 電子帳簿保存法準拠のファイル命名

---

## 対応キャリア

| キャリア | PDFの種類 | 認証方式 |
|---|---|---|
| SoftBank | 電話番号別 / 一括 / 機種別 | SoftBank ID（優先）または電話番号 + SMS認証 |
| Y!mobile | 電話番号別のみ | SoftBank ID（優先）または電話番号 + SMS認証 |
| au | 請求書 / 領収書 / 支払証明書 | au ID + 2段階認証 |
| UQ mobile | 請求書 / 領収書 / 支払証明書 | au ID + 2段階認証 |
| docomo | 一括請求 / 利用内訳 | dアカウント + 2段階認証（SMS 6桁） |

認証情報シートのキャリア列から対象キャリアを自動判定するため、コマンド引数でのキャリア指定は不要。

---

## セットアップ

### 1. Python仮想環境

```bash
# このフォルダ（携帯領収書管理/）で実行
python3 -m venv .venv
source .venv/bin/activate  # Windows: .venv\Scripts\activate
```

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
2. JSONをダウンロードして `ダウンロードツール/client_secrets.json` として配置
3. 初回実行時にブラウザが開くのでGoogleアカウントで許可する
4. `ダウンロードツール/drive_oauth_token.json` にトークンが自動保存される（次回以降不要）

> `client_secrets.json` をPDF保存先Driveフォルダにアップロードしておくと、スクリプトが自動でダウンロードして配置する。

#### サービスアカウント認証（フォールバック）

`client_secrets.json` が存在しない場合に自動的に使用される。

1. サービスアカウントを作成してJSONキーを `ダウンロードツール/service_account.json` として配置
2. スプレッドシートをサービスアカウントの `client_email` に「閲覧者」として共有する
3. PDF保存先のDriveフォルダをサービスアカウントの `client_email` に「編集者」として共有する

#### 機密ファイルの一元管理

`service_account.json` と `client_secrets.json` は **`ダウンロードツール/` フォルダに置くだけでよい**。

### 4. スプレッドシートの準備

2つのスプレッドシートを使用する:
- **認証情報管理シート** — 設定・認証情報・PDFリンク管理（GASを配置する側）
- **回線管理スプレッドシート** — 月別の電話番号・解約済・運用端末等のデータ（別スプレッドシート）

`ダウンロードツール/setup_spreadsheet.gs` を認証情報管理シートのApps Scriptに貼り付けて以下の手順を実行する：

1. `setupSheet()` を実行 → 設定・認証情報・SoftBankリンク・Ymobileリンク・ダウンロード履歴 シートを自動作成
2. 設定シートの「回線管理スプレッドシート」に回線管理表のURLを入力
3. 設定シートの「パスワード」にログインパスワードを入力（全番号共通）
4. メニュー「携帯領収書管理 ツール」→「ダウンロード対象の電話番号を管理」でサイドバーを開き対象を選択

#### シート構成（認証情報管理シート）

| シート名 | 役割 |
|---|---|
| 設定 | PDF保存先・パスワード（SB/YM用・au/UQ用・docomo用）・暗証番号・対象月・回線管理スプシURL |
| 認証情報 | ダウンロード対象の電話番号・キャリア・PDFの種類・運用端末 |
| SoftBankリンク | SoftBank PDF月別リンク（GAS自動更新） |
| Ymobileリンク | Y!mobile PDF月別リンク（GAS自動更新） |
| auリンク | au PDF月別リンク（GAS自動更新） |
| UQmobileリンク | UQ mobile PDF月別リンク（GAS自動更新） |
| docomoリンク | docomo PDF月別リンク（GAS自動更新） |
| ダウンロード履歴 | Pythonスクリプトがダウンロード完了時に自動記録（日時・キャリア・電話番号・対象月・ファイル名・結果） |

#### 認証情報シート（ダウンロード対象の電話番号）

メニュー「携帯領収書管理 ツール」→「ダウンロード対象の電話番号を管理」でHTMLサイドバーを開くと、回線管理スプレッドシートの月別シートから電話番号を動的に読み込み、ダウンロード対象をチェックボックスで選択できる。

| 電話番号 | キャリア | PDFの種類 | 運用端末 | 状態 | ログインID |
|---|---|---|---|---|---|
| 09012345678 | SoftBank | 電話番号別 | iPhoneAir | 契約中 | sb_user01 |
| 08012345678 | Ymobile | 電話番号別 | iPhone16 | 契約中 | |
| 07012345678 | au | 請求書 | GalaxyS25 | 契約中 | user@example.com |
| 06012345678 | UQmobile | 請求書 | AQUOSwish | 契約中 | user@example.com |
| 090XXXXXXXX | SoftBank | 電話番号別 | | 解約済 | sb_user02 |

- パスワードは設定シートで一元管理（認証情報シートにはパスワード列なし）
- `キャリア` 列で SoftBank / Ymobile / au / UQmobile / docomo を自動フィルタリング
- SoftBankは `電話番号別` / `一括` / `機種別` をカンマ区切りで複数指定可
- au/UQは `請求書` / `領収書` / `支払証明書` をカンマ区切りで複数指定可
- docomoは `一括請求`（デフォルト） / `利用内訳`（個別回線）を選択可
- ログインID: SoftBank/Ymobile → SoftBank ID、au/UQ → au ID、docomo → dアカウントID（回線管理スプレッドシートの「ID」列から自動設定）
- SoftBank/Ymobileは解約後SoftBank IDが必須（電話番号でログイン不可）。未設定時はエラー
- 解約済回線もサイドバーで選択すればダウンロード対象になる
- 運用端末・状態はサイドバー保存時に回線管理スプレッドシートから自動設定

#### GASメニュー構成

スプレッドシートを開くと「携帯領収書管理 ツール」メニューが自動追加される。

| メニュー | 機能 |
|---|---|
| 初期セットアップ | `setupSheet()` — シートの初期作成 |
| ダウンロード対象の電話番号を管理 | HTMLサイドバーで対象番号を選択・保存 |
| PDFリンク → 全キャリア一括更新 | DriveのPDFを探索、設定済みセルはスキップしてリンクシートに反映 |
| PDFリンク → SoftBankのみ / Ymobileのみ / auのみ / UQmobileのみ | キャリア別にリンク更新 |
| PDFリンク → 全キャリア一括（強制上書き） | 設定済みセルも含めて全件上書き |
| PDFリンク → SoftBankのみ / Ymobileのみ / auのみ / UQmobileのみ（強制上書き） | キャリア別に強制上書き |


> **金額取得について:** 金額はPythonのダウンロード時にページから自動取得してファイル名に含まれる。GASのOCR金額取得メニューは廃止。

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
| au/UQパスワード | au / UQ mobile 共通のau IDパスワード |
| au暗証番号 | au / UQ mobile の4桁暗証番号（必要な場合） |
| dアカウントパスワード | docomo のdアカウントパスワード |
| 対象月 | ドロップダウン（「自動（前月）」= 前月自動） |

### 5. 環境変数（.env）

`.env` ファイルはスプレッドシートの `.gsheet` ファイルから自動生成される。自動生成されない場合は手動で作成する：

```bash
cp env.example .env
```

`.env` を開いて `SPREADSHEET_URL` にスプレッドシートのURLを貼り付ける（ブラウザのURLをそのままコピーでOK）。

---

## 実行方法

```bash
cd ダウンロードツール/

# 前月分をダウンロード（通常はこれだけ）
python3 download.py

# Drive上の既存PDFのファイル名に金額を反映（ダウンロードは行わない）
python3 update_amounts.py
# または
python3 download.py --update-amounts

# 特定の月を指定（環境変数は設定シートより優先）
TARGET_MONTH=202602 python3 download.py

# ブラウザを表示してデバッグ
HEADLESS=false python3 download.py

# 接続テスト（ダウンロードは行わない）
DRY_RUN=true python3 download.py
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
    echo '123' > /tmp/softbank_security_code.txt   ← SoftBankの場合
    echo '123' > /tmp/ymobile_security_code.txt    ← Y!mobileの場合
    echo '123456' > /tmp/docomo_security_code.txt  ← docomoの場合（6桁）
============================================================
```

- **ターミナル直接実行時**: そのまま3桁を入力してEnterを押す
- **バックグラウンド実行時**: ファイル経由でコードを渡す（キャリアに応じてファイル名が異なる）
  - SoftBank: `echo '854' > /tmp/softbank_security_code.txt`
  - Y!mobile: `echo '854' > /tmp/ymobile_security_code.txt`
- 入力待機時間は `SECURITY_CODE_TIMEOUT`（デフォルト60秒）で変更可能

### セッションの再利用

認証成功後、セッション情報がOSの一時フォルダに電話番号ごとに保存される。次回実行時にセッションが有効であれば、SMS認証がスキップされる。

| OS | セッションファイルの場所 |
|---|---|
| Mac/Linux | `/tmp/softbank_session_{電話番号}.json`（SoftBank）<br>`/tmp/ymobile_session_{電話番号}.json`（Y!mobile） |
| Windows | `%TEMP%\softbank_session_{電話番号}.json`（SoftBank）<br>`%TEMP%\ymobile_session_{電話番号}.json`（Y!mobile） |

認証エラーが解消しない場合はセッションファイルを削除して再実行する：

```bash
# Mac/Linux
rm /tmp/softbank_session_*.json /tmp/ymobile_session_*.json

# Windows
del %TEMP%\softbank_session_*.json %TEMP%\ymobile_session_*.json
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
| `SECURITY_CODE_TIMEOUT` | `60` | SMSセキュリティ番号の入力待機時間（秒） |
| `DRY_RUN` | `false` | `true` で接続テスト付きの検証モード（Drive・ページアクセスを確認、ダウンロードは行わない） |
| `RETRY_PHONES` | （なし） | 特定番号のみ再実行（カンマ区切り。例: `09012345678,08012345678`） |
| `DRIVE_FALLBACK_PATH` | プロジェクトルート | Drive容量超過時のローカル保存先パス |

---

## フォルダ構成

```
携帯領収書管理/
├── README.md                       # このファイル
├── 2026/                           # 領収書PDF（年別フォルダ）
│   └── 02/
│       └── SoftBank/
│           └── *.pdf
├── ダウンロードツール/              # スクリプト・設定一式
│   ├── download.py                 # 統合エントリポイント（--update-amounts で金額更新）
│   ├── update_amounts.py           # 金額更新の個別スクリプト（download.pyのラッパー）
│   ├── shared_utils.py             # 全ロジック集約（CarrierConfig/BillingContext/DriveContext等）
│   ├── setup_spreadsheet.gs        # 携帯領収書管理スプシ用GAS（リンクシート更新のみ）
│   ├── requirements.txt            # Python依存ライブラリ
│   ├── env.example                 # .env のテンプレート
│   ├── .env                        # 自動生成・gitignore対象
│   ├── service_account.json        # GCPサービスアカウントキー（機密・gitignore対象）
│   ├── client_secrets.json         # OAuth2クライアントシークレット（機密・gitignore対象）
│   ├── drive_oauth_token.json      # OAuth2トークン（自動生成・機密・gitignore対象）
│   ├── CLAUDE.md
│   ├── 仕様書.md
│   └── bin/                        # CLAUDE.md自動更新フック
└── 認証情報管理シート.gsheet        # スプレッドシートショートカット（gitignore対象）
```

---

## 機密ファイル一覧

| ファイル | 配置場所 | 生成方法 | 備考 |
|---|---|---|---|
| `service_account.json` | `ダウンロードツール/` | 手動配置 | |
| `client_secrets.json` | `ダウンロードツール/` | 手動配置 | PDF保存先フォルダに置けば自動ダウンロード |
| `drive_oauth_token.json` | `ダウンロードツール/` | 自動生成 | 初回OAuth2認証後に生成される |
| `.env` | `ダウンロードツール/` | 自動生成 | `.gsheet` ファイルから `SPREADSHEET_URL` を自動生成 |

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
PDF保存先フォルダに `client_secrets.json` を置いておくと自動でダウンロードして配置する。または手動で `ダウンロードツール/client_secrets.json` に配置する。

**ファイル名が `_利用料金明細.pdf` のままで金額が入らない**
`python3 update_amounts.py` を実行すると、Drive上のPDFからPyMuPDFでテキストを抽出して金額を取得し、ファイル名を `_○○円(税抜).pdf` に更新する。ダウンロード時にも自動取得するが、取得できなかった場合はこのコマンドで後から更新可能。

**Drive APIでストレージ超過エラー（storageQuotaExceeded）**
サービスアカウント認証時に発生する。OAuth2認証（`client_secrets.json`）に切り替えると、ユーザー自身のDriveにアップロードされるため解消する。

### SoftBank固有

**SMS認証で端末名が表示されない**
認証情報シートの「運用端末」列が空欄。サイドバーで保存し直すと回線管理表から自動設定される。

### Y!mobile固有

**ログインが通らない**
Y!mobileはSoftBank ID（id.my.ymobile.jp）と共通の認証基盤を使用している。パスワードはMy SoftBankと同じIDで管理されている場合がある。

### au / UQ mobile固有

**au IDログインが通らない**
au IDはメールアドレスまたは電話番号。設定シートの「au/UQパスワード」にau IDのパスワードが正しく設定されているか確認する。

**暗証番号の入力を求められる**
WEB de 請求書の閲覧に4桁の暗証番号が必要な場合がある。設定シートの「au暗証番号」に設定する。

**2段階認証でワンタイムURLが届く**
au IDの2段階認証にはSMS確認コード方式とワンタイムURL方式がある。ワンタイムURL方式の場合、SMSに届いたURLをタップして認証を完了する必要がある（タイムアウトまでに操作）。

### docomo固有

**dアカウントログインが通らない**
dアカウントIDはメールアドレスまたは電話番号。設定シートの「dアカウントパスワード」にdアカウントのパスワードが正しく設定されているか確認する。

**2段階認証でSMSコードが届かない**
dアカウントの2段階認証はSMSまたはメールで6桁のセキュリティコードが届く。SMS/メールが届かない場合はdアカウントの設定を確認する。

**パスキー認証を要求される**
2026年5月頃にdアカウントはパスキー認証に統一される予定。パスキーが強制された場合、現在のID+パスワード方式では認証できなくなる。

**Web料金明細サービスが未契約**
利用内訳のダウンロードには「Web料金明細サービス」の契約が必要（無料）。My docomoから事前に申し込む。

---

## 技術仕様

詳細な技術仕様（処理フロー・認証フロー・スプレッドシート構成・環境変数など）は [ダウンロードツール/仕様書.md](ダウンロードツール/仕様書.md) を参照。
