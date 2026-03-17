# My SoftBank 料金明細PDF自動ダウンロード

セットアップ・実行方法・トラブルシューティングは **[../README.md](../README.md)** を参照。

## SoftBank固有の設定

- スクリプト: `mysoftbank_billing.py`
- PDFの種類: `電話番号別` / `一括` / `機種別`（アカウントシートのPDFの種類列にカンマ区切りで指定）
- 機密ファイル配置場所: `SoftBank/`（Y!mobileスクリプトもここを自動参照）
- 仕様書: [仕様書.md](仕様書.md)

## 実行

```bash
cd SoftBank/
python3 mysoftbank_billing.py
```
