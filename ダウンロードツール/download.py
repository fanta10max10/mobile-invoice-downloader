#!/usr/bin/env python3
"""
携帯領収書管理 統合ダウンロードスクリプト

使い方:
  python3 download.py                  # 通常ダウンロード
  python3 download.py --update-amounts # Drive上のPDFファイル名に金額を反映
"""

import argparse
import logging
import re
import sys
import tempfile
from pathlib import Path

from shared_utils import (
    CarrierConfig, create_billing_context, run_main,
    extract_amount_from_pdf, resolve_save_path,
    get_gspread_client, open_sheet,
)

log = logging.getLogger(__name__)

SOFTBANK_CONFIG = CarrierConfig(
    carrier_name="SoftBank",
    display_name="My SoftBank",
    auth_domain="id.my.softbank.jp",
    login_url="https://my.softbank.jp/msb/d/webLink/doSend/MSB010000",
    wco_base="https://bl11.my.softbank.jp/wco",
    certificate_url="https://bl11.my.softbank.jp/wco/certificate/WCO250",
    bill_pdf_url="https://bl11.my.softbank.jp/wco/external/goBillInfoPdf",
    code_file_prefix="softbank",
    session_file_prefix="softbank",
    temp_dir_prefix="softbank_pdf_",
)

YMOBILE_CONFIG = CarrierConfig(
    carrier_name="Ymobile",
    display_name="My Y!mobile",
    auth_domain="id.my.ymobile.jp",
    login_url="https://my.ymobile.jp/muc/d/webLink/doSend/WCO010023",
    wco_base="https://bl61.my.ymobile.jp/wco",
    certificate_url="https://bl61.my.ymobile.jp/wco/certificate/WCO250",
    bill_pdf_url="https://bl61.my.ymobile.jp/wco/external/goBillInfoPdf",
    code_file_prefix="ymobile",
    session_file_prefix="ymobile",
    temp_dir_prefix="ymobile_pdf_",
)

AU_CONFIG = CarrierConfig(
    carrier_name="au",
    display_name="My au",
    auth_domain="connect.auone.jp",
    login_url="https://id.auone.jp/index.html",
    company_name="KDDI株式会社",
    carrier_family="au",
    au_billing_top_url="https://my.au.com/aus/hc-cs/lic/LIC0020001.hc",
    au_pin_setting_name="au暗証番号",
    password_setting_name="au/UQパスワード",
    code_file_prefix="au",
    session_file_prefix="au",
    temp_dir_prefix="au_pdf_",
)

UQ_CONFIG = CarrierConfig(
    carrier_name="UQmobile",
    display_name="My UQ mobile",
    auth_domain="connect.auone.jp",
    login_url="https://id.auone.jp/index.html",
    company_name="KDDI株式会社",
    carrier_family="au",
    au_billing_top_url="https://my.au.com/aus/hc-cs/lic/LIC0020001.hc",
    au_pin_setting_name="au暗証番号",
    password_setting_name="au/UQパスワード",
    code_file_prefix="uqmobile",
    session_file_prefix="uqmobile",
    temp_dir_prefix="uqmobile_pdf_",
)

ALL_CARRIERS = [SOFTBANK_CONFIG, YMOBILE_CONFIG, AU_CONFIG, UQ_CONFIG]


def update_amounts():
    """Drive上の _利用料金明細.pdf を探してPDFから金額取得→リネーム。
    対象月のキャリアフォルダのみ探索するため高速。
    """
    script_dir = Path(__file__).resolve().parent
    ctx = create_billing_context(ALL_CARRIERS[0], script_dir=script_dir)
    resolve_save_path(ctx)  # DriveContext を初期化

    if not ctx.drive_ctx:
        log.error("Drive APIモードでのみ実行可能です。PDF保存先フォルダにDriveのURLを設定してください。")
        sys.exit(1)

    service = ctx.drive_ctx.service
    root_id = ctx.drive_ctx.base_folder_id

    # 対象月を取得
    from shared_utils import get_target_month
    year, month = get_target_month(ctx)
    if not ctx.target_month:
        try:
            gc = get_gspread_client()
            sh = open_sheet(gc, ctx.spreadsheet_id)
            ws = sh.worksheet("設定")
            from datetime import datetime
            for row in ws.get_all_records():
                if str(row.get("設定名", "")).strip() == "対象月":
                    raw_val = row.get("値", "")
                    val = str(raw_val).strip()
                    if re.match(r"^\d{6}$", val):
                        year, month = val[:4], val[4:6]
                    elif m2 := re.match(r"^(\d{4})年(\d+)月$", val):
                        year, month = m2.group(1), m2.group(2).zfill(2)
                    elif isinstance(raw_val, datetime):
                        year, month = str(raw_val.year), str(raw_val.month).zfill(2)
                    break
        except Exception:
            pass
    log.info(f"対象月: {year}年{month}月")

    # root/YYYY/MM 直下のキャリアフォルダだけ探索
    targets = []
    year_folders = service.files().list(
        q=f"name='{year}' and '{root_id}' in parents and mimeType='application/vnd.google-apps.folder' and trashed=false",
        fields="files(id)"
    ).execute().get("files", [])
    for yf in year_folders:
        month_folders = service.files().list(
            q=f"name='{month}' and '{yf['id']}' in parents and mimeType='application/vnd.google-apps.folder' and trashed=false",
            fields="files(id)"
        ).execute().get("files", [])
        for mf in month_folders:
            carrier_folders = service.files().list(
                q=f"'{mf['id']}' in parents and mimeType='application/vnd.google-apps.folder' and trashed=false",
                fields="files(id,name)"
            ).execute().get("files", [])
            for cf in carrier_folders:
                pdfs = service.files().list(
                    q=f"'{cf['id']}' in parents and mimeType='application/pdf' and trashed=false",
                    fields="files(id,name)"
                ).execute().get("files", [])
                # 請求書本体のみ（支払証明書・領収書・一括・機種別は除外）
                exclude_suffixes = ("_支払証明書.pdf", "_領収書.pdf", "_一括.pdf", "_機種別.pdf")
                for pdf in pdfs:
                    name = pdf["name"]
                    if any(name.endswith(s) for s in exclude_suffixes):
                        continue
                    if name.endswith(".pdf"):
                        targets.append(pdf)

    log.info(f"対象PDF: {len(targets)}件")
    if not targets:
        log.info("金額未取得のPDFはありません。")
        return

    # キャリア名→会社名マップ
    company_map = {}
    for config in ALL_CARRIERS:
        company_map[config.carrier_name] = config.company_name

    updated = 0
    skipped = 0
    failed = 0
    for f in targets:
        name = f["name"]
        # ファイル名パース: 新形式 YYYYMM_会社名_carrier_phone_*.pdf
        m = re.match(r"(\d{6})_.+?_(SoftBank|Ymobile|au|UQmobile)_(\d+)_.+\.pdf", name)
        if not m:
            # 旧形式 YYYYMM_carrier_phone_*.pdf
            m = re.match(r"(\d{6})_(SoftBank|Ymobile|au|UQmobile)_(\d+)_.+\.pdf", name)
        if not m:
            continue
        ym, carrier, phone = m.group(1), m.group(2), m.group(3)
        is_au = carrier in ("au", "UQmobile")
        company = company_map.get(carrier, "")

        content = service.files().get_media(fileId=f["id"]).execute()
        tmp = Path(tempfile.mktemp(suffix=".pdf"))
        tmp.write_bytes(content)
        amount = extract_amount_from_pdf(tmp, phone if is_au else "")
        tmp.unlink()

        if amount:
            new_name = f"{ym}_{company}_{carrier}_{phone}_{amount}.pdf"
            if new_name == name:
                skipped += 1
                continue
            service.files().update(fileId=f["id"], body={"name": new_name}).execute()
            log.info(f"  ✅ {name} → {new_name}")
            updated += 1
        else:
            log.warning(f"  ❌ {name}: 金額取得失敗")
            failed += 1

    log.info("=" * 50)
    log.info(f"金額更新完了: {updated}件更新 / {skipped}件変更なし / {failed}件失敗")
    log.info("=" * 50)


def main():
    parser = argparse.ArgumentParser(description="携帯領収書管理 統合ダウンロードスクリプト")
    parser.add_argument("--update-amounts", action="store_true",
                        help="Drive上のPDFファイル名に金額を反映（ダウンロードは行わない）")
    args = parser.parse_args()

    if args.update_amounts:
        update_amounts()
        return

    script_dir = Path(__file__).resolve().parent
    all_results = []
    for config in ALL_CARRIERS:
        ctx = create_billing_context(config, script_dir=script_dir)
        carrier_results = run_main(ctx) or []
        all_results.extend(carrier_results)

    # 全キャリア最終サマリー
    if all_results:
        status_labels = {"success": "✅ 成功", "skipped": "⏭️ ダウンロード済み", "failed": "❌ 失敗"}
        log.info("=" * 50)
        log.info("処理結果サマリー:")
        for carrier, phone, result in all_results:
            log.info(f"  [{carrier}] {phone}: {status_labels.get(result, result)}")
        n_success = sum(1 for *_, r in all_results if r == "success")
        n_skipped = sum(1 for *_, r in all_results if r == "skipped")
        n_failed = sum(1 for *_, r in all_results if r == "failed")
        log.info(f"  合計: {n_success} 件成功 / {n_skipped} 件スキップ / {n_failed} 件失敗")
        log.info("=" * 50)
        if n_failed > 0:
            sys.exit(1)


if __name__ == "__main__":
    main()
