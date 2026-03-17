/**
 * 認証情報管理シート セットアップスクリプト
 *
 * 【このスクリプトについて】
 * SoftBank・Y!mobile 両キャリアで1つのスプレッドシートを共用するための統合スクリプト。
 * 以下のシートを作成・初期化します:
 *
 *   設定         ... PDF保存先フォルダ、対象月（SoftBank・Ymobile共通）
 *   認証情報     ... 電話番号・パスワード・キャリア・PDFの種類（ユーザーが入力）
 *   SoftBankリンク  ... SoftBank PDFの月別リンク管理（GAS自動更新）
 *   Ymobileリンク   ... Ymobile PDFの月別リンク管理（GAS自動更新）
 *
 * 【使い方】
 *   1. 認証情報管理シートを開く
 *   2. 拡張機能 → Apps Script を開く
 *   3. このコードを貼り付けて保存
 *   4. setupSheet() を実行（初回は権限承認が必要）
 *   5. 「設定」シートにPDF保存先フォルダURLと対象月を入力
 *   6. 「認証情報」シートに電話番号・パスワード・キャリアを入力
 *
 * 【PDFリンク更新】
 *   「携帯領収書管理 ツール」メニュー →「SoftBank PDFリンクを更新」/「Ymobile PDFリンクを更新」
 *
 * 【PDFの種類 の値】
 *   電話番号別              ... 電話番号別PDF（デフォルト）
 *   一括                    ... 一括印刷用PDF
 *   機種別                  ... 機種別PDF（SoftBankのみ）
 *   電話番号別,一括         ... カンマ区切りで複数指定可
 */

// ────────────────────────────────────────────────
//  定数
// ────────────────────────────────────────────────

const SETTINGS_SHEET_NAME = "設定";
const AUTH_SHEET_NAME = "認証情報";
const SOFTBANK_LINK_SHEET_NAME = "SoftBankリンク";
const YMOBILE_LINK_SHEET_NAME = "Ymobileリンク";

// 月列はD列から開始（電話番号 / PDFの種類 列: A=電話番号 B=PDFの種類）
const MONTH_COL_START = 3; // C列（1-indexed）

const PDF_TYPE_OPTIONS = [
  "電話番号別",
  "一括",
  "機種別",
  "電話番号別,一括",
  "電話番号別,機種別",
  "一括,機種別",
  "電話番号別,一括,機種別",
];


// ────────────────────────────────────────────────
//  初期セットアップ
// ────────────────────────────────────────────────

/**
 * スプレッドシートの初期セットアップ
 * 設定・認証情報・SoftBankリンク・Ymobileリンク の各シートを作成する
 */
function setupSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  ss.rename("認証情報管理シート");
  setupSettingsSheet_(ss);
  setupAuthSheet_(ss);
  setupLinkSheet_(ss, SOFTBANK_LINK_SHEET_NAME, "SoftBank");
  setupLinkSheet_(ss, YMOBILE_LINK_SHEET_NAME, "Ymobile");
  refreshPhoneDropdowns();

  SpreadsheetApp.getUi().alert(
    "セットアップ完了",
    "以下のシートを作成・初期化しました。\n\n" +
    "【設定シート】\n" +
    "  「PDF保存先フォルダ」にGoogle DriveのフォルダURLを入力してください\n" +
    "  「対象月」はドロップダウンで選択してください（「自動（前月）」= 前月を自動選択）\n\n" +
    "【認証情報シート】\n" +
    "  キャリア列（SoftBank / Ymobile）を選択すると、\n" +
    "  電話番号列に月別管理シートの番号がドロップダウンで表示されます\n\n" +
    "【SoftBankリンク / Ymobileリンク シート】\n" +
    "  PDFダウンロード後、メニューからリンク更新を実行するとPDFへのリンクが記録されます",
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}


/**
 * 設定シートを作成・初期化する（内部用）
 */
function setupSettingsSheet_(ss) {
  let sheet = ss.getSheetByName(SETTINGS_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SETTINGS_SHEET_NAME);
    // 先頭に移動（既存の月シートより前に配置）
    ss.setActiveSheet(sheet);
    ss.moveActiveSheet(1);
  }

  if (!sheet.getRange(1, 1).getValue()) {
    const headerRange = sheet.getRange(1, 1, 1, 2);
    headerRange.setValues([["設定名", "値"]]);
    headerRange
      .setFontWeight("bold")
      .setBackground("#FF6D00")
      .setFontColor("#FFFFFF")
      .setHorizontalAlignment("center");
  }

  sheet.setColumnWidth(1, 200);
  sheet.setColumnWidth(2, 500);

  // 認証情報管理シートURL（自動設定・.env の SPREADSHEET_URL に使用）
  _upsertSettingRow_(sheet, "認証情報管理シート",
    ss.getUrl(),
    "このスプレッドシートのURL。\nSoftBank/.env および Ymobile/.env の SPREADSHEET_URL にコピーしてください。"
  );

  _upsertSettingRow_(sheet, "PDF保存先フォルダ",
    "https://drive.google.com/drive/folders/XXXXXXXX",
    "PDFの保存先フォルダ。以下どちらでも可:\n" +
    "・Google DriveのフォルダURL（推奨）\n" +
    "  https://drive.google.com/drive/folders/...\n" +
    "  → Drive APIで直接アップロード（要: サービスアカウントを編集者として共有）\n" +
    "・ローカル絶対パス（/Users/... または C:\\...）"
  );
  _upsertSettingRow_(sheet, "対象月",
    "自動（前月）",
    "ダウンロードする月を選択してください\n「自動（前月）」= 実行時の前月を自動選択"
  );

  _setTargetMonthValidation_(sheet);
}


/**
 * 認証情報シートを作成・初期化する（内部用）
 * 列: 電話番号 | パスワード | キャリア | PDFの種類
 */
function setupAuthSheet_(ss) {
  let sheet = ss.getSheetByName(AUTH_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(AUTH_SHEET_NAME);
    ss.setActiveSheet(sheet);
    ss.moveActiveSheet(2);
  }

  // ヘッダー行（空の場合のみ書き込む）
  if (!sheet.getRange(1, 1).getValue()) {
    const headers = ["電話番号", "パスワード", "キャリア", "PDFの種類"];
    const headerRange = sheet.getRange(1, 1, 1, headers.length);
    headerRange.setValues([headers]);
    headerRange
      .setFontWeight("bold")
      .setBackground("#4285F4")
      .setFontColor("#FFFFFF")
      .setHorizontalAlignment("center");
  }

  sheet.setColumnWidth(1, 180);  // 電話番号
  sheet.setColumnWidth(2, 220);  // パスワード
  sheet.setColumnWidth(3, 120);  // キャリア
  sheet.setColumnWidth(4, 220);  // PDFの種類

  // 電話番号列を書式なしテキストに（先頭0が消えないように）
  sheet.getRange("A:A").setNumberFormat("@");

  // キャリア列にドロップダウンを設定
  const lastRow = Math.max(sheet.getLastRow(), 100);
  const carrierRange = sheet.getRange(2, 3, lastRow - 1, 1);
  const carrierRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(["SoftBank", "Ymobile"], true)
    .setAllowInvalid(false)
    .build();
  carrierRange.setDataValidation(carrierRule);

  // PDFの種類列にドロップダウンを設定
  const pdfTypeRange = sheet.getRange(2, 4, lastRow - 1, 1);
  const pdfTypeRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(PDF_TYPE_OPTIONS, true)
    .setAllowInvalid(true)
    .setHelpText("ダウンロードするPDFの種類を選択。複数はカンマ区切りで直接入力。")
    .build();
  pdfTypeRange.setDataValidation(pdfTypeRule);

  // サンプルデータ（シートが空の場合のみ）
  if (sheet.getLastRow() <= 1) {
    sheet.getRange(2, 1, 2, 4).setValues([
      ["090-XXXX-XXXX", "SoftBankのパスワード", "SoftBank", "電話番号別"],
      ["080-XXXX-XXXX", "Y!mobileのパスワード", "Ymobile", "電話番号別"],
    ]);
  }

  // ツールチップ
  sheet.getRange("A1").setNote("電話番号。ハイフンOK（スクリプトが自動で除去）");
  sheet.getRange("B1").setNote("各キャリアのログインパスワード");
  sheet.getRange("C1").setNote("キャリア名（SoftBank または Ymobile）");
  sheet.getRange("D1").setNote(
    "ダウンロードするPDFの種類（カンマ区切りで複数指定可）\n\n" +
    "電話番号別 … 電話番号別PDF（デフォルト）\n" +
    "一括       … 一括印刷用PDF\n" +
    "機種別     … 機種別PDF（SoftBankのみ）"
  );
}


/**
 * PDFリンク管理シート（SoftBankリンク / Ymobileリンク）を作成・初期化する
 * 列: 電話番号 | PDFの種類 | 2026年1月 | 2026年2月 | ...
 */
function setupLinkSheet_(ss, sheetName, carrierLabel) {
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
  }

  if (!sheet.getRange(1, 1).getValue()) {
    const headers = ["電話番号", "PDFの種類"];
    const color = carrierLabel === "SoftBank" ? "#4285F4" : "#FF6D00";
    const headerRange = sheet.getRange(1, 1, 1, headers.length);
    headerRange.setValues([headers]);
    headerRange
      .setFontWeight("bold")
      .setBackground(color)
      .setFontColor("#FFFFFF")
      .setHorizontalAlignment("center");
  }

  sheet.setColumnWidth(1, 180);
  sheet.setColumnWidth(2, 220);
  sheet.getRange("A:A").setNumberFormat("@");

  sheet.getRange("A1").setNote(`${carrierLabel}の電話番号（ハイフンOK）`);
  sheet.getRange("B1").setNote("PDFの種類（認証情報シートと合わせてください）");
}


// ────────────────────────────────────────────────
//  PDFリンク更新
// ────────────────────────────────────────────────

/** SoftBank PDFリンクを更新する */
function updateSoftBankLinks() {
  _updatePdfLinks_(SOFTBANK_LINK_SHEET_NAME, "SoftBank");
}

/** Ymobile PDFリンクを更新する */
function updateYmobileLinks() {
  _updatePdfLinks_(YMOBILE_LINK_SHEET_NAME, "Ymobile");
}

/**
 * Google Drive を探索してPDFリンクをシートに設定する（内部用）
 *
 * 探索するフォルダ構造:
 *   {設定シートのPDF保存先}/YYYY/MM/{carrier}/*.pdf  ← 新形式
 *   {設定シートのPDF保存先}/YYYY/MM/*.pdf            ← 旧形式（後方互換）
 */
function _updatePdfLinks_(sheetName, carrier) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    SpreadsheetApp.getUi().alert("エラー", `「${sheetName}」シートが見つかりません。setupSheet() を実行してください。`, SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }

  const rootFolderUrl = _getSettingValue_("PDF保存先フォルダ");
  if (!rootFolderUrl || !rootFolderUrl.startsWith("https://drive.google.com/")) {
    SpreadsheetApp.getUi().alert(
      "エラー",
      "「設定」シートの「PDF保存先フォルダ」にGoogle DriveのフォルダURLが設定されていません。",
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    return;
  }

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
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const phoneColIdx = headers.indexOf("電話番号");
  if (phoneColIdx === -1) {
    SpreadsheetApp.getUi().alert("エラー", "「電話番号」列が見つかりません。", SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }
  const phoneToRow = {};
  for (let i = 1; i < data.length; i++) {
    const phone = String(data[i][phoneColIdx] || "").replace(/[-\s]/g, "").trim();
    if (phone) phoneToRow[phone] = i + 1;
  }

  // PDFを収集
  const pdfEntries = [];
  const yearFolders = rootFolder.getFolders();
  while (yearFolders.hasNext()) {
    const yearFolder = yearFolders.next();
    if (!/^\d{4}$/.test(yearFolder.getName())) continue;

    const monthFolders = yearFolder.getFolders();
    while (monthFolders.hasNext()) {
      const monthFolder = monthFolders.next();
      if (!/^\d{2}$/.test(monthFolder.getName())) continue;

      // 新形式: キャリア名サブフォルダ内
      const subFolders = monthFolder.getFolders();
      while (subFolders.hasNext()) {
        const sub = subFolders.next();
        if (sub.getName() === carrier) {
          _collectCarrierPdfs_(sub, carrier, pdfEntries);
        }
      }
      // 旧形式: 月フォルダ直下（後方互換）
      _collectCarrierPdfs_(monthFolder, carrier, pdfEntries);
    }
  }

  if (pdfEntries.length === 0) {
    SpreadsheetApp.getUi().alert("情報", `${carrier}のPDFファイルが見つかりませんでした。\nGoogle Driveの同期が完了しているか確認してください。`, SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }

  // 必要な月列を確保
  const neededHeaders = [...new Set(pdfEntries.map(e => `${e.year}年${parseInt(e.month)}月`))];
  neededHeaders.sort((a, b) => _monthHeaderToNum_(b) - _monthHeaderToNum_(a));
  for (const header of neededHeaders) {
    _ensureMonthColumn_(sheet, header);
  }

  // ハイパーリンクを書き込む
  let updatedCount = 0;
  const processed = new Set();
  for (const entry of pdfEntries) {
    const rowNum = phoneToRow[entry.phone];
    if (!rowNum) continue;

    const colNum = _getMonthColumnNum_(sheet, entry.year, entry.month);
    if (!colNum) continue;

    const key = `${entry.phone}_${entry.year}${entry.month}`;
    if (processed.has(key)) {
      const cell = sheet.getRange(rowNum, colNum);
      const note = cell.getNote() || "";
      cell.setNote((note ? note + "\n" : "") + entry.file.getName());
      continue;
    }
    processed.add(key);

    const url = entry.file.getUrl();
    const amountMatch = entry.file.getName().match(/_(\d+)円\.pdf$/);
    const label = amountMatch
      ? `${parseInt(entry.month)}月 ${Number(amountMatch[1]).toLocaleString()}円`
      : `${parseInt(entry.month)}月`;
    sheet.getRange(rowNum, colNum).setFormula(`=HYPERLINK("${url}","${label}")`);
    updatedCount++;
  }

  SpreadsheetApp.getUi().alert(
    "完了",
    `${pdfEntries.length} 件のPDFを確認し、${updatedCount} 件のリンクを設定しました。`,
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}


// ────────────────────────────────────────────────
//  PDFから金額を取得・ファイル名更新
// ────────────────────────────────────────────────

/** SoftBank PDFから金額を取得してファイル名・リンクを更新する */
function scanAndUpdateSoftBankAmounts() {
  _scanAndUpdatePdfAmounts_(SOFTBANK_LINK_SHEET_NAME, "SoftBank");
}

/** Ymobile PDFから金額を取得してファイル名・リンクを更新する */
function scanAndUpdateYmobileAmounts() {
  _scanAndUpdatePdfAmounts_(YMOBILE_LINK_SHEET_NAME, "Ymobile");
}

/**
 * ファイル名に金額が入っていないPDFをOCRで解析し、
 * ファイル名とスプレッドシートのリンクラベルを更新する（内部用）
 *
 * 【事前準備】拡張機能 → Apps Script → サービス から
 *   「Drive API」(v2) を追加すること。
 */
function _scanAndUpdatePdfAmounts_(sheetName, carrier) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    SpreadsheetApp.getUi().alert("エラー", `「${sheetName}」シートが見つかりません。`, SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }

  const rootFolderUrl = _getSettingValue_("PDF保存先フォルダ");
  if (!rootFolderUrl || !rootFolderUrl.startsWith("https://drive.google.com/")) {
    SpreadsheetApp.getUi().alert("エラー", "「設定」シートの「PDF保存先フォルダ」にGoogle DriveのフォルダURLが設定されていません。", SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }

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

  const pdfEntries = [];
  const yearFolders = rootFolder.getFolders();
  while (yearFolders.hasNext()) {
    const yearFolder = yearFolders.next();
    if (!/^\d{4}$/.test(yearFolder.getName())) continue;
    const monthFolders = yearFolder.getFolders();
    while (monthFolders.hasNext()) {
      const monthFolder = monthFolders.next();
      if (!/^\d{2}$/.test(monthFolder.getName())) continue;
      const subFolders = monthFolder.getFolders();
      while (subFolders.hasNext()) {
        const sub = subFolders.next();
        if (sub.getName() === carrier) _collectCarrierPdfs_(sub, carrier, pdfEntries);
      }
      _collectCarrierPdfs_(monthFolder, carrier, pdfEntries);
    }
  }

  const targets = pdfEntries.filter(e => e.file.getName().endsWith("_利用料金明細.pdf"));
  if (targets.length === 0) {
    SpreadsheetApp.getUi().alert("情報", "金額が未取得のPDFはありませんでした。", SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }

  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const phoneColIdx = headers.indexOf("電話番号");
  const phoneToRow = {};
  for (let i = 1; i < data.length; i++) {
    const phone = String(data[i][phoneColIdx] || "").replace(/[-\s]/g, "").trim();
    if (phone) phoneToRow[phone] = i + 1;
  }

  let updatedCount = 0;
  let failedCount = 0;
  for (const entry of targets) {
    const amount = _extractAmountFromPdf_(entry.file);
    if (!amount) {
      failedCount++;
      Logger.log(`金額取得失敗: ${entry.file.getName()}`);
      continue;
    }

    const newName = entry.file.getName().replace("_利用料金明細.pdf", `_${amount}円.pdf`);
    entry.file.setName(newName);

    const rowNum = phoneToRow[entry.phone];
    if (rowNum) {
      const colNum = _getMonthColumnNum_(sheet, entry.year, entry.month);
      if (colNum) {
        const url = entry.file.getUrl();
        const label = `${parseInt(entry.month)}月 ${Number(amount).toLocaleString()}円`;
        sheet.getRange(rowNum, colNum).setFormula(`=HYPERLINK("${url}","${label}")`);
      }
    }
    updatedCount++;
  }

  SpreadsheetApp.getUi().alert(
    "完了",
    `${targets.length} 件を処理しました。\n` +
    `  更新: ${updatedCount} 件\n` +
    `  金額取得失敗: ${failedCount} 件`,
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}


// ────────────────────────────────────────────────
//  内部ヘルパー関数
// ────────────────────────────────────────────────

function _getSettingValue_(key) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SETTINGS_SHEET_NAME);
  if (!sheet) return null;
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === key) {
      return String(data[i][1]).trim() || null;
    }
  }
  return null;
}


/**
 * フォルダ内のキャリア命名規則PDFを収集する
 * ファイル名形式: YYYYMM_{carrier}_{phone}_*.pdf
 */
function _collectCarrierPdfs_(folder, carrier, results) {
  const files = folder.getFiles();
  while (files.hasNext()) {
    const file = files.next();
    if (file.getMimeType() !== "application/pdf") continue;
    const pattern = new RegExp(`^(\\d{4})(\\d{2})_${carrier}_(\\d+)`);
    const m = file.getName().match(pattern);
    if (!m) continue;
    results.push({ year: m[1], month: m[2], phone: m[3], file });
  }
}


function _monthHeaderToNum_(header) {
  const m = String(header).match(/^(\d{4})年(\d+)月$/);
  if (!m) return 0;
  return parseInt(m[1]) * 100 + parseInt(m[2]);
}


function _ensureMonthColumn_(sheet, monthHeader) {
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  if (headers.some(h => String(h).trim() === monthHeader)) return;

  const targetNum = _monthHeaderToNum_(monthHeader);
  let insertAt = sheet.getLastColumn() + 1;

  for (let i = MONTH_COL_START - 1; i < headers.length; i++) {
    const colNum = _monthHeaderToNum_(String(headers[i]));
    if (colNum > 0 && targetNum < colNum) {
      insertAt = i + 1;
      sheet.insertColumnBefore(insertAt);
      break;
    }
  }

  const headerCell = sheet.getRange(1, insertAt);
  headerCell.setNumberFormat("@");
  headerCell
    .setValue(monthHeader)
    .setFontWeight("bold")
    .setBackground("#34A853")
    .setFontColor("#FFFFFF")
    .setHorizontalAlignment("center");
  sheet.setColumnWidth(insertAt, 80);
}


function _getMonthColumnNum_(sheet, year, month) {
  const header = `${year}年${parseInt(month)}月`;
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const idx = headers.findIndex(h => String(h).trim() === header);
  return idx === -1 ? null : idx + 1;
}


function _upsertSettingRow_(sheet, key, defaultValue, noteText) {
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === key) return;
  }
  const newRow = sheet.getLastRow() + 1;
  sheet.getRange(newRow, 1).setValue(key);
  const valCell = sheet.getRange(newRow, 2);
  valCell.setValue(defaultValue);
  if (noteText) valCell.setNote(noteText);
}


function _setTargetMonthValidation_(sheet) {
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() !== "対象月") continue;

    const options = ["自動（前月）"];
    const now = new Date();
    for (let m = 0; m < 13; m++) {
      const d = new Date(now.getFullYear(), now.getMonth() - m, 1);
      options.push(`${d.getFullYear()}年${d.getMonth() + 1}月`);
    }

    const valCell = sheet.getRange(i + 1, 2);
    const rule = SpreadsheetApp.newDataValidation()
      .requireValueInList(options, true)
      .setAllowInvalid(false)
      .setHelpText("ダウンロードする月を選択。「自動（前月）」なら実行時の前月を自動選択します。")
      .build();
    valCell.setDataValidation(rule);

    if (!valCell.getValue()) {
      valCell.setValue("自動（前月）");
    }
    return;
  }
}


/**
 * PDFファイルをOCRで読み取り、請求金額（数値文字列）を返す。
 * 取得できなければ null を返す。
 * Drive API v2 (Advanced Google Services) が必要。
 */
function _extractAmountFromPdf_(file) {
  let docFileId = null;
  try {
    const blob = file.getBlob();
    const docFile = Drive.Files.insert(
      { title: "tmp_ocr_" + file.getName() },
      blob,
      { convert: true }
    );
    docFileId = docFile.id;
    const doc = DocumentApp.openById(docFileId);
    const text = doc.getBody().getText();

    const patterns = [
      /(?<![小合])計[^\d]*([\d,]+)/,
      /小計[^\d]*([\d,]+)/,
    ];
    for (const pattern of patterns) {
      const m = text.match(pattern);
      if (m) {
        const amount = m[1].replace(/,/g, "");
        if (/^\d+$/.test(amount)) return amount;
      }
    }
    Logger.log(`金額パターン不一致。OCRテキスト先頭200文字: ${text.substring(0, 200)}`);
    return null;
  } catch (e) {
    Logger.log(`OCRエラー (${file.getName()}): ${e.message}`);
    return null;
  } finally {
    if (docFileId) {
      try { Drive.Files.remove(docFileId); } catch (_) {}
    }
  }
}


// ────────────────────────────────────────────────
//  認証情報シート：電話番号ドロップダウン動的更新
//
//  仕組み:
//    _PhoneDropdown シート（非表示）に各キャリアの電話番号を列ごとに格納し、
//    認証情報シートの電話番号列は ONE_OF_RANGE バリデーションで参照する。
//    キャリア列（C列）変更時に onEdit が該当行のバリデーションを更新する。
// ────────────────────────────────────────────────

const PHONE_DROPDOWN_SHEET = "_PhoneDropdown";

/**
 * セルが編集されたとき自動実行。
 * 認証情報シートのキャリア列（C列）が変更されたら
 * 同行の電話番号列（A列）のドロップダウンを更新する。
 */
function onEdit(e) {
  if (!e) return;
  const sheet = e.range.getSheet();
  if (sheet.getName() !== AUTH_SHEET_NAME) return;
  const col = e.range.getColumn();
  const row = e.range.getRow();
  if (col !== 3 || row <= 1) return;
  _setPhoneDropdownForRow_(sheet, row, String(e.value || "").trim());
}


/**
 * 月別管理シートを読み直して _PhoneDropdown を更新し、
 * 認証情報シート全行のバリデーションを再設定する。
 * メニューから実行するか、月別データ更新後に実行する。
 */
function refreshPhoneDropdowns() {
  _refreshPhoneDropdownData_();

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const authSheet = ss.getSheetByName(AUTH_SHEET_NAME);
  if (!authSheet) return;

  const data = authSheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    const carrier = String(data[i][2] || "").trim();
    _setPhoneDropdownForRow_(authSheet, i + 1, carrier);
  }

  SpreadsheetApp.getUi().alert(
    "完了",
    "電話番号ドロップダウンを更新しました。",
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}


/**
 * 指定行の電話番号セル（A列）のバリデーションを _PhoneDropdown の
 * 該当キャリア列を参照する ONE_OF_RANGE に設定する。
 */
function _setPhoneDropdownForRow_(sheet, row, carrier) {
  const phoneCell = sheet.getRange(row, 1);
  if (!carrier) {
    phoneCell.clearDataValidations();
    return;
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const pdSheet = ss.getSheetByName(PHONE_DROPDOWN_SHEET);
  if (!pdSheet) {
    phoneCell.clearDataValidations();
    return;
  }

  const headers = pdSheet.getRange(1, 1, 1, pdSheet.getLastColumn()).getValues()[0];
  const colIdx = headers.findIndex(h => String(h).trim() === carrier);
  if (colIdx === -1) {
    phoneCell.clearDataValidations();
    return;
  }

  const lastRow = Math.max(pdSheet.getLastRow(), 2);
  const phoneRange = pdSheet.getRange(2, colIdx + 1, lastRow - 1, 1);
  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInRange(phoneRange, true)
    .setAllowInvalid(true)
    .setHelpText(`${carrier}の契約中番号（_PhoneDropdown シートより）`)
    .build();
  phoneCell.setDataValidation(rule);
}


/**
 * 月別管理シートを解析して _PhoneDropdown シートを更新する。
 * 各キャリアの電話番号（ハイフンなし）を列に並べて書き込む。
 * キャリア名 → 列のマッピングは1行目のヘッダーで管理する。
 */
function _refreshPhoneDropdownData_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // 月別シートを最新順に検索
  const monthSheets = ss.getSheets().filter(ws => /.*\d+月$/.test(ws.getName()));
  if (monthSheets.length === 0) return;

  monthSheets.sort((a, b) => _parseMonthSheetNum_(a.getName()) - _parseMonthSheetNum_(b.getName()));
  let targetSheet = null;
  for (let i = monthSheets.length - 1; i >= 0; i--) {
    if (monthSheets[i].getLastRow() > 1) { targetSheet = monthSheets[i]; break; }
  }
  if (!targetSheet) return;

  Logger.log(`[_refreshPhoneDropdownData_] 使用シート: ${targetSheet.getName()}`);

  const rows = targetSheet.getDataRange().getValues();
  const carrierPhones = {}; // { "SoftBank": [...], "Ymobile": [...] }
  let cols = null;

  for (let i = 0; i < rows.length; i++) {
    const row = rows[i];
    // ヘッダー行を検出（「電話番号」列を含む行）
    if (cols === null) {
      const idx = row.findIndex(c => String(c).trim() === "電話番号");
      if (idx !== -1) {
        cols = {};
        row.forEach((c, j) => { const k = String(c).trim(); if (k) cols[k] = j; });
      }
      continue;
    }

    const phone = String(row[cols["電話番号"]] || "").replace(/[-\s]/g, "").trim();
    if (!phone || !/^\d{10,13}$/.test(phone)) continue;

    const ci = cols["解約済"];
    if (ci !== undefined && String(row[ci] || "").toUpperCase() === "TRUE") continue;

    const ki = cols["キャリア"];
    const carrier = ki !== undefined ? String(row[ki] || "").trim() : "";
    if (!carrier) continue;

    if (!carrierPhones[carrier]) carrierPhones[carrier] = [];
    if (!carrierPhones[carrier].includes(phone)) carrierPhones[carrier].push(phone);
  }

  Logger.log(`[_refreshPhoneDropdownData_] 収集結果: ${JSON.stringify(Object.fromEntries(Object.entries(carrierPhones).map(([k,v]) => [k, v.length])))}`);

  // _PhoneDropdown シートを取得（なければ作成）
  let pdSheet = ss.getSheetByName(PHONE_DROPDOWN_SHEET);
  if (!pdSheet) {
    pdSheet = ss.insertSheet(PHONE_DROPDOWN_SHEET);
    pdSheet.hideSheet();
  }
  pdSheet.clearContents();

  // キャリア順（SoftBank → Ymobile → その他）で書き込む
  const orderedCarriers = ["SoftBank", "Ymobile",
    ...Object.keys(carrierPhones).filter(c => c !== "SoftBank" && c !== "Ymobile").sort()
  ].filter(c => carrierPhones[c]);

  const maxLen = Math.max(...orderedCarriers.map(c => carrierPhones[c].length));
  const output = [orderedCarriers];
  for (let r = 0; r < maxLen; r++) {
    output.push(orderedCarriers.map(c => carrierPhones[c][r] || ""));
  }

  pdSheet.getRange(1, 1, output.length, orderedCarriers.length).setValues(output);
  Logger.log(`[_refreshPhoneDropdownData_] _PhoneDropdown 更新完了`);
}


/**
 * 月シート名を YYYYMM 数値に変換（ソート用）。
 * "(2025)12月" → 202512, "3月" → 202603（現在年）
 */
function _parseMonthSheetNum_(name) {
  const m1 = name.match(/\((\d{4})\)(\d+)月/);
  if (m1) return parseInt(m1[1]) * 100 + parseInt(m1[2]);
  const m2 = name.match(/(\d+)月$/);
  if (m2) return new Date().getFullYear() * 100 + parseInt(m2[1]);
  return 0;
}


/**
 * 電話番号ドロップダウンの動作確認ダイアログを表示する。
 */
function debugPhoneDropdown() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const pdSheet = ss.getSheetByName(PHONE_DROPDOWN_SHEET);
  if (!pdSheet) {
    SpreadsheetApp.getUi().alert("エラー", `「${PHONE_DROPDOWN_SHEET}」シートが見つかりません。\n「電話番号ドロップダウンを更新」を実行してください。`, SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }

  const rows = pdSheet.getDataRange().getValues();
  const headers = rows[0] || [];
  const lines = headers.map((h, col) => {
    const phones = rows.slice(1).map(r => r[col]).filter(v => v);
    return `${h} (${phones.length}件): ${phones.join(", ") || "なし"}`;
  });

  SpreadsheetApp.getUi().alert(
    "電話番号ドロップダウン デバッグ",
    `_PhoneDropdown シートの内容:\n\n${lines.join("\n\n")}`,
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}


// ────────────────────────────────────────────────
//  メニュー
// ────────────────────────────────────────────────

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("携帯領収書管理 ツール")
    .addItem("初期セットアップ", "setupSheet")
    .addSeparator()
    .addItem("電話番号ドロップダウンを更新", "refreshPhoneDropdowns")
    .addItem("電話番号ドロップダウン デバッグ", "debugPhoneDropdown")
    .addSeparator()
    .addItem("SoftBank PDFリンクを更新", "updateSoftBankLinks")
    .addItem("Ymobile PDFリンクを更新", "updateYmobileLinks")
    .addSeparator()
    .addItem("SoftBank PDFから金額を取得・ファイル名更新", "scanAndUpdateSoftBankAmounts")
    .addItem("Ymobile PDFから金額を取得・ファイル名更新", "scanAndUpdateYmobileAmounts")
    .addToUi();
}
