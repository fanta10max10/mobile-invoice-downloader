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
 * PDFの種類 の値:
 *   電話番号別              ... 電話番号別PDF（デフォルト）
 *   一括                    ... 一括印刷用PDF
 *   機種別                  ... 機種別PDF
 *   電話番号別,一括         ... 電話番号別 + 一括（カンマ区切りで複数指定可）
 *   電話番号別,一括,機種別  ... すべてダウンロード
 */

// PDFの種類 のドロップダウン選択肢
const PDF_TYPE_OPTIONS = [
  "電話番号別",
  "一括",
  "機種別",
  "電話番号別,一括",
  "電話番号別,機種別",
  "一括,機種別",
  "電話番号別,一括,機種別",
];

const SHEET_NAME = "アカウント";

/**
 * スプレッドシートの初期セットアップ
 */
function setupSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // シート名を設定
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.getActiveSheet();
    sheet.setName(SHEET_NAME);
  }

  // ヘッダー行
  const headers = ["電話番号", "パスワード", "PDF保存先フォルダ", "PDFの種類"];
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
      "パスワードをここに入力",
      "https://drive.google.com/drive/folders/XXXXXXXX",
      "電話番号別",
    ],
  ];
  sheet.getRange(2, 1, sample.length, sample[0].length).setValues(sample);

  // 列幅の調整
  sheet.setColumnWidth(1, 180);  // 電話番号
  sheet.setColumnWidth(2, 200);  // パスワード
  sheet.setColumnWidth(3, 450);  // PDF保存先フォルダ
  sheet.setColumnWidth(4, 220);  // PDFの種類

  // 電話番号列を書式なしテキストに（先頭0が消えないように）
  sheet.getRange("A:A").setNumberFormat("@");

  // PDFの種類列にドロップダウンを設定（2行目以降）
  const lastRow = Math.max(sheet.getLastRow(), 100);
  const pdfTypeRange = sheet.getRange(2, 4, lastRow - 1, 1);
  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(PDF_TYPE_OPTIONS, true)
    .setAllowInvalid(true)  // カンマ区切り直接入力も許可
    .setHelpText(
      "ダウンロードするPDFの種類を選択してください。\n" +
      "複数選択する場合はカンマ区切りで直接入力（例: 電話番号別,一括）\n\n" +
      "電話番号別 … 電話番号別PDF\n" +
      "一括       … 一括印刷用PDF\n" +
      "機種別     … 機種別PDF"
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
    "電話番号別 … 電話番号別PDF（デフォルト）\n" +
    "一括       … 一括印刷用PDF\n" +
    "機種別     … 機種別PDF\n\n" +
    "例: 電話番号別,一括 → 電話番号別と一括の両方をDL"
  );

  // 不要なシートを削除
  const sheets = ss.getSheets();
  for (const s of sheets) {
    if (s.getName() !== SHEET_NAME && sheets.length > 1) {
      ss.deleteSheet(s);
    }
  }

  // スプレッドシート名を設定
  ss.rename("MySoftBank_アカウント管理");

  SpreadsheetApp.getUi().alert(
    "セットアップ完了",
    `シート「${SHEET_NAME}」を作成しました。\n\n` +
    "1. サンプル行を実際のアカウント情報に書き換えてください\n" +
    "2. 「PDFの種類」列でダウンロードするPDFの種類を選択してください（デフォルト: 電話番号別）\n" +
    "3. 完了後、publishAsCsv() を実行してCSV URLを取得してください",
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}


/**
 * 既存シートに「PDFの種類」列を追加する（既にsetupSheet済みのシート用）
 */
function addPdfTypeColumn() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    SpreadsheetApp.getUi().alert("エラー", `「${SHEET_NAME}」シートが見つかりません。`, SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }

  // ヘッダー行を確認して「PDFの種類」列が既にあるか確認
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  if (headers.includes("PDFの種類")) {
    SpreadsheetApp.getUi().alert("情報", "「PDFの種類」列は既に存在します。", SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }

  // 最終列の次に追加
  const newCol = sheet.getLastColumn() + 1;
  sheet.getRange(1, newCol).setValue("PDFの種類")
    .setFontWeight("bold")
    .setBackground("#4285F4")
    .setFontColor("#FFFFFF")
    .setHorizontalAlignment("center");
  sheet.setColumnWidth(newCol, 220);

  // 既存データ行にデフォルト値「電話番号別」を設定
  const lastRow = sheet.getLastRow();
  if (lastRow >= 2) {
    sheet.getRange(2, newCol, lastRow - 1, 1).setValue("電話番号別");
  }

  // ドロップダウンを設定
  const pdfTypeRange = sheet.getRange(2, newCol, Math.max(lastRow - 1, 99), 1);
  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(PDF_TYPE_OPTIONS, true)
    .setAllowInvalid(true)
    .setHelpText(
      "ダウンロードするPDFの種類を選択してください。\n" +
      "複数選択する場合はカンマ区切りで直接入力（例: 電話番号別,一括）\n\n" +
      "電話番号別 … 電話番号別PDF\n" +
      "一括       … 一括印刷用PDF\n" +
      "機種別     … 機種別PDF"
    )
    .build();
  pdfTypeRange.setDataValidation(rule);

  sheet.getRange(1, newCol).setNote(
    "ダウンロードするPDFの種類（カンマ区切りで複数指定可）\n\n" +
    "電話番号別 … 電話番号別PDF（デフォルト）\n" +
    "一括       … 一括印刷用PDF\n" +
    "機種別     … 機種別PDF\n\n" +
    "例: 電話番号別,一括 → 電話番号別と一括の両方をDL"
  );

  SpreadsheetApp.getUi().alert(
    "完了",
    "「PDFの種類」列を追加しました。\n既存行にはデフォルト値「電話番号別」を設定済みです。\n\nスプレッドシートを再公開してCSV URLを更新してください。",
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
      <li>公開対象を <b>「アカウント」シート</b> に変更</li>
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
 * Google Drive を探索してPDFリンクをスプレッドシートに設定する
 *
 * 探索するフォルダ構造:
 *   {PDF保存先フォルダ}/YYYY/MM/SoftBank/*.pdf  ← 新形式
 *   {PDF保存先フォルダ}/YYYY/MM/*.pdf           ← 旧形式（後方互換）
 *
 * 見つかったPDFはE列以降の月列（例: 2026年2月）にハイパーリンクとして記録される。
 * 月列は日付昇順（左が古い）に自動整列される。
 */
function updatePdfLinks() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    SpreadsheetApp.getUi().alert("エラー", `「${SHEET_NAME}」シートが見つかりません。`, SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }

  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const phoneColIdx = headers.indexOf("電話番号");
  const folderColIdx = headers.indexOf("PDF保存先フォルダ");
  if (phoneColIdx === -1 || folderColIdx === -1) {
    SpreadsheetApp.getUi().alert("エラー", "「電話番号」または「PDF保存先フォルダ」列が見つかりません。", SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }

  // ルートフォルダURLを取得（最初の非空URL）
  let rootFolderUrl = "";
  for (let i = 1; i < data.length; i++) {
    const val = String(data[i][folderColIdx] || "").trim();
    if (val.startsWith("https://drive.google.com/")) {
      rootFolderUrl = val;
      break;
    }
  }
  if (!rootFolderUrl) {
    SpreadsheetApp.getUi().alert("エラー", "「PDF保存先フォルダ」にGoogle DriveのフォルダURLが設定されていません。\n例: https://drive.google.com/drive/folders/XXXX", SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }

  // フォルダID抽出
  const folderMatch = rootFolderUrl.match(/\/folders\/([a-zA-Z0-9_-]+)/);
  if (!folderMatch) {
    SpreadsheetApp.getUi().alert("エラー", "フォルダURLからIDを取得できませんでした。", SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }
  let rootFolder;
  try {
    rootFolder = DriveApp.getFolderById(folderMatch[1]);
  } catch (e) {
    SpreadsheetApp.getUi().alert("エラー", `フォルダへのアクセスに失敗しました: ${e.message}`, SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }

  // 電話番号 → 行番号マップ（ハイフン除去済み）
  const phoneToRow = {};
  for (let i = 1; i < data.length; i++) {
    const phone = String(data[i][phoneColIdx] || "").replace(/[-\s]/g, "").trim();
    if (phone) phoneToRow[phone] = i + 1; // 1-indexed
  }

  // YYYY/MM/SoftBank/ および YYYY/MM/ 直下のPDFを収集
  // エントリ形式: { year, month, phone, file }
  const pdfEntries = [];
  const yearFolders = rootFolder.getFolders();
  while (yearFolders.hasNext()) {
    const yearFolder = yearFolders.next();
    if (!/^\d{4}$/.test(yearFolder.getName())) continue;

    const monthFolders = yearFolder.getFolders();
    while (monthFolders.hasNext()) {
      const monthFolder = monthFolders.next();
      if (!/^\d{2}$/.test(monthFolder.getName())) continue;

      // 新形式: SoftBank サブフォルダ内
      const subFolders = monthFolder.getFolders();
      while (subFolders.hasNext()) {
        const sub = subFolders.next();
        if (sub.getName() === "SoftBank") {
          _collectSoftBankPdfs(sub, pdfEntries);
        }
      }
      // 旧形式: 月フォルダ直下（キャリアフォルダ追加前との互換）
      _collectSoftBankPdfs(monthFolder, pdfEntries);
    }
  }

  if (pdfEntries.length === 0) {
    SpreadsheetApp.getUi().alert("情報", "SoftBankのPDFファイルが見つかりませんでした。\nGoogle Driveの同期が完了しているか確認してください。", SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }

  // 必要な月列を確保（右→左の順で挿入してインデックスのずれを防ぐ）
  const neededHeaders = [...new Set(pdfEntries.map(e => `${e.year}年${parseInt(e.month)}月`))];
  neededHeaders.sort((a, b) => _monthHeaderToNum(b) - _monthHeaderToNum(a)); // 降順（右→左に挿入）
  for (const header of neededHeaders) {
    _ensureMonthColumn(sheet, header);
  }

  // ハイパーリンクを書き込む（同月複数PDFは2件目以降をセルのメモに追記）
  let updatedCount = 0;
  const processed = new Set();
  for (const entry of pdfEntries) {
    const rowNum = phoneToRow[entry.phone];
    if (!rowNum) continue;

    const colNum = _getMonthColumnNum(sheet, entry.year, entry.month);
    if (!colNum) continue;

    const key = `${entry.phone}_${entry.year}${entry.month}`;
    if (processed.has(key)) {
      // 2件目以降: メモに追記
      const cell = sheet.getRange(rowNum, colNum);
      const note = cell.getNote() || "";
      cell.setNote((note ? note + "\n" : "") + entry.file.getName());
      continue;
    }
    processed.add(key);

    const url = entry.file.getUrl();
    const label = `${parseInt(entry.month)}月`;
    sheet.getRange(rowNum, colNum).setFormula(`=HYPERLINK("${url}","${label}")`);
    updatedCount++;
  }

  SpreadsheetApp.getUi().alert(
    "完了",
    `${pdfEntries.length} 件のPDFを確認し、${updatedCount} 件のリンクを設定しました。`,
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}


/**
 * フォルダ内のSoftBank命名規則PDFを収集するヘルパー
 * ファイル名形式: YYYYMM_SoftBank_{phone}_*.pdf
 */
function _collectSoftBankPdfs(folder, results) {
  const files = folder.getFiles();
  while (files.hasNext()) {
    const file = files.next();
    if (file.getMimeType() !== "application/pdf") continue;
    const m = file.getName().match(/^(\d{4})(\d{2})_SoftBank_(\d+)/);
    if (!m) continue;
    results.push({ year: m[1], month: m[2], phone: m[3], file });
  }
}


/**
 * "2025年2月" → 202502 に変換（数値ソート用）
 */
function _monthHeaderToNum(header) {
  const m = String(header).match(/^(\d{4})年(\d+)月$/);
  if (!m) return 0;
  return parseInt(m[1]) * 100 + parseInt(m[2]);
}


const MONTH_COL_START = 5; // E列（1-indexed）


/**
 * 月列が存在しなければ日付順の位置に挿入する
 */
function _ensureMonthColumn(sheet, monthHeader) {
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  if (headers.includes(monthHeader)) return; // 既に存在

  const targetNum = _monthHeaderToNum(monthHeader);
  let insertAt = sheet.getLastColumn() + 1; // デフォルトは末尾

  for (let i = MONTH_COL_START - 1; i < headers.length; i++) {
    const colNum = _monthHeaderToNum(String(headers[i]));
    if (colNum > 0 && targetNum < colNum) {
      insertAt = i + 1; // 1-indexed
      sheet.insertColumnBefore(insertAt);
      break;
    }
  }

  // ヘッダーを緑色で設定
  sheet.getRange(1, insertAt)
    .setValue(monthHeader)
    .setFontWeight("bold")
    .setBackground("#34A853")
    .setFontColor("#FFFFFF")
    .setHorizontalAlignment("center");
  sheet.setColumnWidth(insertAt, 80);
}


/**
 * 月列の列番号（1-indexed）を返す。存在しなければ null。
 */
function _getMonthColumnNum(sheet, year, month) {
  const header = `${year}年${parseInt(month)}月`;
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const idx = headers.indexOf(header);
  return idx === -1 ? null : idx + 1;
}


/**
 * メニューにカスタムメニューを追加
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("MySoftBank ツール")
    .addItem("初期セットアップ", "setupSheet")
    .addSeparator()
    .addItem("「PDFの種類」列を追加（既存シート用）", "addPdfTypeColumn")
    .addSeparator()
    .addItem("PDFリンクを更新", "updatePdfLinks")
    .addSeparator()
    .addItem("CSV URL 取得手順を表示", "publishAsCsv")
    .addToUi();
}
