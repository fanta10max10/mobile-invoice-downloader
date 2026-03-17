# My Y!mobile 料金明細PDF自動ダウンロード

セットアップ・実行方法・トラブルシューティングは **[../README.md](../README.md)** を参照。

## Y!mobile固有の設定

- スクリプト: `myyumobile_billing.py`
- PDFの種類: `電話番号別`のみ
- 機密ファイル: `../SoftBank/` を自動検索（Ymobileフォルダへの配置も可）
- 認証基盤: SoftBank IDと共通（id.my.ymobile.jp）

## 実行

```bash
cd Ymobile/
python3 myyumobile_billing.py
```
