/**
 * My SoftBank 料金明細ダウンロード用スプレッドシート セットアップスクリプト
 *
 * 使い方:
 *   1. Google スプレッドシートを新規作成
 *   2. 拡張機能 → Apps Script を開く
 *   3. このコードを貼り付けて保存
 *   4. setupSheet() を実行（初回は権限承認が必要）
 *   5. シートにアカウント情報を入力
 *   6. publishAsCsv() を実行して「ウェブに公開」の手順を表示
 *
 * pdf_type の値:
 *   phone        ... 電話番号別PDF（デフォルト）
 *   bulk         ... 一括印刷用PDF
 *   device       ... 機種別PDF
 *   phone,bulk   ... 電話番号別 + 一括（カンマ区切りで複数指定可）
 *   phone,bulk,device ... すべてダウンロード
 */

// pdf_type のドロップダウン選択肢
const PDF_TYPE_OPTIONS = [
  "phone",
  "bulk",
  "device",
  "phone,bulk",
  "phone,device",
  "bulk,device",
  "phone,bulk,device",
];

/**
 * スプレッドシートの初期セットアップ
 */
function setupSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // シート名を設定
  let sheet = ss.getSheetByName("accounts");
  if (!sheet) {
    sheet = ss.getActiveSheet();
    sheet.setName("accounts");
  }

  // ヘッダー行
  const headers = ["phone_number", "password", "PDF保存先", "pdf_type"];
  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setValues([headers]);

  // ヘッダーの書式設定
  headerRange
    .setFontWeight("bold")
    .setBackground("#4285F4")
    .setFontColor("#FFFFFF")
    .setHorizontalAlignment("center");

  // サンプルデータ
  const sample = [
    [
      "090-XXXX-XXXX",
      "password123",
      "https://drive.google.com/drive/folders/XXXXXXXX",
      "phone",
    ],
  ];
  sheet.getRange(2, 1, sample.length, sample[0].length).setValues(sample);

  // 列幅の調整
  sheet.setColumnWidth(1, 180);  // phone_number
  sheet.setColumnWidth(2, 180);  // password
  sheet.setColumnWidth(3, 450);  // PDF保存先
  sheet.setColumnWidth(4, 200);  // pdf_type

  // phone_number列を書式なしテキストに（先頭0が消えないように）
  sheet.getRange("A:A").setNumberFormat("@");

  // pdf_type列にドロップダウンを設定（2行目以降）
  const lastRow = Math.max(sheet.getLastRow(), 100);
  const pdfTypeRange = sheet.getRange(2, 4, lastRow - 1, 1);
  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(PDF_TYPE_OPTIONS, true)
    .setAllowInvalid(true)  // カンマ区切り直接入力も許可
    .setHelpText(
      "ダウンロードするPDFの種類を選択してください。\n" +
      "複数選択する場合はカンマ区切りで直接入力（例: phone,bulk）\n\n" +
      "phone  … 電話番号別PDF\n" +
      "bulk   … 一括印刷用PDF\n" +
      "device … 機種別PDF"
    )
    .build();
  pdfTypeRange.setDataValidation(rule);

  // 列の注釈（ツールチップ）
  sheet.getRange("A1").setNote(
    "SoftBank ID（携帯電話番号）。ハイフンOK（スクリプトが自動で除去）\n例: 090-XXXX-XXXX"
  );
  sheet.getRange("B1").setNote("My SoftBankのログインパスワード");
  sheet.getRange("C1").setNote(
    "PDFの保存先。以下どちらでも可:\n" +
    "・Google DriveのフォルダURL（https://drive.google.com/drive/folders/...）\n" +
    "  → drive_path_map.txt でローカルパスに変換される\n" +
    "・Macのローカル絶対パス（/Users/...）\n\n" +
    "1行目に書けば全アカウント共通。2行目以降は空欄でOK。"
  );
  sheet.getRange("D1").setNote(
    "ダウンロードするPDFの種類（カンマ区切りで複数指定可）\n\n" +
    "phone  … 電話番号別PDF（デフォルト）\n" +
    "bulk   … 一括印刷用PDF\n" +
    "device … 機種別PDF\n\n" +
    "例: phone,bulk → 電話番号別と一括の両方をDL"
  );

  // 不要なシートを削除
  const sheets = ss.getSheets();
  for (const s of sheets) {
    if (s.getName() !== "accounts" && sheets.length > 1) {
      ss.deleteSheet(s);
    }
  }

  // スプレッドシート名を設定
  ss.rename("MySoftBank_アカウント管理");

  SpreadsheetApp.getUi().alert(
    "セットアップ完了",
    "シート「accounts」を作成しました。\n\n" +
    "1. サンプル行を実際のアカウント情報に書き換えてください\n" +
    "2. pdf_type 列でダウンロードするPDFの種類を選択してください（デフォルト: phone）\n" +
    "3. 完了後、publishAsCsv() を実行してCSV URLを取得してください",
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}


/**
 * 既存シートに pdf_type 列を追加する（既にsetupSheet済みのシート用）
 */
function addPdfTypeColumn() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("accounts");
  if (!sheet) {
    SpreadsheetApp.getUi().alert("エラー", "「accounts」シートが見つかりません。", SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }

  // ヘッダー行を確認してpdf_type列が既にあるか確認
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  if (headers.includes("pdf_type")) {
    SpreadsheetApp.getUi().alert("情報", "pdf_type 列は既に存在します。", SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }

  // 最終列の次に追加
  const newCol = sheet.getLastColumn() + 1;
  sheet.getRange(1, newCol).setValue("pdf_type")
    .setFontWeight("bold")
    .setBackground("#4285F4")
    .setFontColor("#FFFFFF")
    .setHorizontalAlignment("center");
  sheet.setColumnWidth(newCol, 200);

  // 既存データ行にデフォルト値 "phone" を設定
  const lastRow = sheet.getLastRow();
  if (lastRow >= 2) {
    sheet.getRange(2, newCol, lastRow - 1, 1).setValue("phone");
  }

  // ドロップダウンを設定
  const pdfTypeRange = sheet.getRange(2, newCol, Math.max(lastRow - 1, 99), 1);
  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(PDF_TYPE_OPTIONS, true)
    .setAllowInvalid(true)
    .setHelpText(
      "ダウンロードするPDFの種類を選択してください。\n" +
      "複数選択する場合はカンマ区切りで直接入力（例: phone,bulk）\n\n" +
      "phone  … 電話番号別PDF\n" +
      "bulk   … 一括印刷用PDF\n" +
      "device … 機種別PDF"
    )
    .build();
  pdfTypeRange.setDataValidation(rule);

  sheet.getRange(1, newCol).setNote(
    "ダウンロードするPDFの種類（カンマ区切りで複数指定可）\n\n" +
    "phone  … 電話番号別PDF（デフォルト）\n" +
    "bulk   … 一括印刷用PDF\n" +
    "device … 機種別PDF\n\n" +
    "例: phone,bulk → 電話番号別と一括の両方をDL"
  );

  SpreadsheetApp.getUi().alert(
    "完了",
    "pdf_type 列を追加しました。\n既存行にはデフォルト値「phone」を設定済みです。\n\nスプレッドシートを再公開してCSV URLを更新してください。",
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}


/**
 * 「ウェブに公開」の手順をダイアログで表示する
 * ※ GASからプログラムで「ウェブに公開」はできないため、手順を案内する
 */
function publishAsCsv() {
  const expectedUrl =
    `https://docs.google.com/spreadsheets/d/e/2PACX-***/pub?gid=0&single=true&output=csv`;

  const html = HtmlService.createHtmlOutput(`
    <style>
      body { font-family: sans-serif; font-size: 14px; padding: 12px; }
      h3 { margin-top: 0; }
      ol li { margin-bottom: 8px; }
      code { background: #f0f0f0; padding: 2px 6px; border-radius: 3px; font-size: 13px; }
      .warn { color: #d93025; font-weight: bold; }
      .url-box { background: #f0f0f0; padding: 8px; border-radius: 4px;
                 word-break: break-all; font-family: monospace; font-size: 12px; }
    </style>
    <h3>CSV URLの取得手順</h3>
    <ol>
      <li>このスプレッドシートのメニュー →<br>
          <b>ファイル → 共有 → ウェブに公開</b></li>
      <li>公開対象を <b>「accounts」シート</b> に変更</li>
      <li>形式を <b>「カンマ区切り形式（.csv）」</b> に変更</li>
      <li><b>「公開」</b> ボタンをクリック</li>
      <li>表示されたURLをコピーして <code>MySoftBank_アカウント管理スプシURL.rtf</code> に貼り付ける</li>
    </ol>
    <p>URL形式の例:</p>
    <div class="url-box">${expectedUrl}</div>
    <br>
    <p class="warn">
      ⚠ このURLを知っている人は誰でもアクセスできます。<br>
      パスワードを含むため、URLの共有は厳禁です。<br>
      不要になったら「ウェブに公開」を必ず停止してください。
    </p>
  `)
    .setWidth(520)
    .setHeight(440);

  SpreadsheetApp.getUi().showModalDialog(html, "CSV URLの取得手順");
}


/**
 * メニューにカスタムメニューを追加
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("MySoftBank ツール")
    .addItem("初期セットアップ", "setupSheet")
    .addSeparator()
    .addItem("pdf_type 列を追加（既存シート用）", "addPdfTypeColumn")
    .addSeparator()
    .addItem("CSV URL 取得手順を表示", "publishAsCsv")
    .addToUi();
}
