#!/usr/bin/env python3
"""
携帯領収書管理 統合ダウンロードスクリプト

使い方:
  python3 download.py softbank     # SoftBankのみ
  python3 download.py ymobile      # Y!mobileのみ
  python3 download.py all          # 両方実行
"""

import sys
from pathlib import Path

from shared_utils import CarrierConfig, create_billing_context, run_main

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

CARRIERS = {
    "softbank": SOFTBANK_CONFIG,
    "sb": SOFTBANK_CONFIG,
    "ymobile": YMOBILE_CONFIG,
    "ym": YMOBILE_CONFIG,
}


def main():
    if len(sys.argv) < 2 or sys.argv[1] in ("-h", "--help"):
        print(__doc__.strip())
        print(f"\n指定可能なキャリア: {', '.join(CARRIERS.keys())}, all")
        sys.exit(0)

    arg = sys.argv[1].lower()
    script_dir = Path(__file__).resolve().parent

    if arg == "all":
        targets = [SOFTBANK_CONFIG, YMOBILE_CONFIG]
    elif arg in CARRIERS:
        targets = [CARRIERS[arg]]
    else:
        print(f"エラー: 不明なキャリア '{sys.argv[1]}'")
        print(f"指定可能: {', '.join(CARRIERS.keys())}, all")
        sys.exit(1)

    for config in targets:
        ctx = create_billing_context(config, script_dir=script_dir)
        run_main(ctx)


if __name__ == "__main__":
    main()
