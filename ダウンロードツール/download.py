#!/usr/bin/env python3
"""
携帯領収書管理 統合ダウンロードスクリプト

使い方:
  python3 download.py
"""

import logging
import sys
from pathlib import Path

from shared_utils import CarrierConfig, create_billing_context, run_main

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
    carrier_family="au",
    au_billing_top_url="https://my.au.com/aus/hc-cs/lic/LIC0020001.hc",
    au_pin_setting_name="au暗証番号",
    password_setting_name="au/UQパスワード",
    code_file_prefix="uqmobile",
    session_file_prefix="uqmobile",
    temp_dir_prefix="uqmobile_pdf_",
)

ALL_CARRIERS = [SOFTBANK_CONFIG, YMOBILE_CONFIG, AU_CONFIG, UQ_CONFIG]


def main():
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
