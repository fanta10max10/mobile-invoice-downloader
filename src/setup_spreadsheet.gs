/**
 * My SoftBank 料金明細ダウンロード用スプレッドシート セットアップスクリプト
 *
 * 使い方:
 *   1. Google スプレッドシートを新規作成
 *   2. 拡張機能 → Apps Script を開く
 *   3. このコードを貼り付けて保存
 *   4. setupSheet() を実行（初回は権限承認が必要）
 *   5. 「設定」シートにPDF保存先フォルダURLと対象月を入力
 *   6. 「アカウント」シートにアカウント情報を入力
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
const SETTINGS_SHEET_NAME = "設定";

// 月列はD列から開始（アカウント列: A=電話番号 B=パスワード C=PDFの種類）
const MONTH_COL_START = 4; // D列（1-indexed）


// ────────────────────────────────────────────────
//  初期セットアップ
// ────────────────────────────────────────────────

/**
 * スプレッドシートの初期セットアップ
 * アカウントシートと設定シートを両方作成する
 */
function setupSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // ── アカウントシートを作成 ──
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.getActiveSheet();
    sheet.setName(SHEET_NAME);
  }

  // ヘッダー行（PDF保存先フォルダは設定シートへ移動）
  const headers = ["電話番号", "パスワード", "PDFの種類"];
  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setValues([headers]);

  // ヘッダーの書式設定
  headerRange
    .setFontWeight("bold")
    .setBackground("#4285F4")
    .setFontColor("#FFFFFF")
    .setHorizontalAlignment("center");

  // サンプルデータ（シートが空の場合のみ書き込む。既存データは上書きしない）
  if (sheet.getLastRow() <= 1) {
    sheet.getRange(2, 1, 1, 3).setValues([["090-XXXX-XXXX", "パスワードをここに入力", "電話番号別"]]);
  }

  // 列幅の調整
  sheet.setColumnWidth(1, 180);  // 電話番号
  sheet.setColumnWidth(2, 200);  // パスワード
  sheet.setColumnWidth(3, 220);  // PDFの種類

  // 電話番号列を書式なしテキストに（先頭0が消えないように）
  sheet.getRange("A:A").setNumberFormat("@");

  // PDFの種類列にドロップダウンを設定（2行目以降）
  const lastRow = Math.max(sheet.getLastRow(), 100);
  const pdfTypeRange = sheet.getRange(2, 3, lastRow - 1, 1);
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
    "ダウンロードするPDFの種類（カンマ区切りで複数指定可）\n\n" +
    "電話番号別 … 電話番号別PDF（デフォルト）\n" +
    "一括       … 一括印刷用PDF\n" +
    "機種別     … 機種別PDF\n\n" +
    "例: 電話番号別,一括 → 電話番号別と一括の両方をDL"
  );

  // ── 設定シートを作成 ──
  setupSettingsSheet_(ss);

  // ── 不要なシートを削除（アカウントと設定は保持） ──
  const sheets = ss.getSheets();
  const keepNames = new Set([SHEET_NAME, SETTINGS_SHEET_NAME]);
  for (const s of sheets) {
    if (!keepNames.has(s.getName()) && sheets.length > keepNames.size) {
      ss.deleteSheet(s);
    }
  }

  // スプレッドシート名を設定
  ss.rename("MySoftBank_アカウント管理");

  SpreadsheetApp.getUi().alert(
    "セットアップ完了",
    "シートを作成しました。\n\n" +
    "【設定シート】\n" +
    "  「PDF保存先フォルダ」にGoogle DriveのフォルダURLを入力してください\n" +
    "  「対象月」はドロップダウンで選択してください（「自動（前月）」= 前月を自動選択）\n\n" +
    "【アカウントシート】\n" +
    "  サンプル行を実際のアカウント情報に書き換えてください",
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}


/**
 * 設定シートを作成・初期化する（内部用）
 * 既存の設定値は上書きしない。未登録の項目のみ追加する。
 */
function setupSettingsSheet_(ss) {
  let sheet = ss.getSheetByName(SETTINGS_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SETTINGS_SHEET_NAME);
  }

  // ヘッダー行（空の場合のみ書き込む）
  if (!sheet.getRange(1, 1).getValue()) {
    const headerRange = sheet.getRange(1, 1, 1, 2);
    headerRange.setValues([["設定名", "値"]]);
    headerRange
      .setFontWeight("bold")
      .setBackground("#FF6D00")
      .setFontColor("#FFFFFF")
      .setHorizontalAlignment("center");
  }

  // 列幅
  sheet.setColumnWidth(1, 200);
  sheet.setColumnWidth(2, 500);

  // 設定項目を追加（既存行は上書きしない）
  _upsertSettingRow_(sheet, "PDF保存先フォルダ",
    "https://drive.google.com/drive/folders/XXXXXXXX",
    "PDFの保存先フォルダ。以下どちらでも可:\n" +
    "・Google DriveのフォルダURL（https://drive.google.com/drive/folders/...）\n" +
    "  → drive_path_map.txt でローカルパスに変換される\n" +
    "・Macのローカル絶対パス（/Users/...）"
  );
  _upsertSettingRow_(sheet, "対象月",
    "自動（前月）",
    "ダウンロードする月を選択してください\n「自動（前月）」= 実行時の前月を自動選択"
  );

  // 「対象月」の値セルにドロップダウンを設定（常に最新の月リストで更新）
  _setTargetMonthValidation_(sheet);
}


/**
 * 「対象月」行の値セルに直近13ヶ月のドロップダウンを設定する
 */
function _setTargetMonthValidation_(sheet) {
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() !== "対象月") continue;

    // 直近13ヶ月の選択肢を生成（現在月〜13ヶ月前）
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

    // 未設定の場合はデフォルト値をセット
    if (!valCell.getValue()) {
      valCell.setValue("自動（前月）");
    }
    return;
  }
}


/**
 * 設定シートに指定キーの行がなければ追加する。既存行は上書きしない。
 */
function _upsertSettingRow_(sheet, key, defaultValue, noteText) {
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === key) return; // 既存行があればスキップ
  }
  const newRow = sheet.getLastRow() + 1;
  sheet.getRange(newRow, 1).setValue(key);
  const valCell = sheet.getRange(newRow, 2);
  valCell.setValue(defaultValue);
  if (noteText) valCell.setNote(noteText);
}


// ────────────────────────────────────────────────
//  移行ヘルパー（既存シート用）
// ────────────────────────────────────────────────

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
    "「PDFの種類」列を追加しました。\n既存行にはデフォルト値「電話番号別」を設定済みです。",
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}


// ────────────────────────────────────────────────
//  PDFリンク更新
// ────────────────────────────────────────────────

/**
 * Google Drive を探索してPDFリンクをスプレッドシートに設定する
 *
 * 探索するフォルダ構造:
 *   {設定シートのPDF保存先}/YYYY/MM/SoftBank/*.pdf  ← 新形式
 *   {設定シートのPDF保存先}/YYYY/MM/*.pdf           ← 旧形式（後方互換）
 *
 * 見つかったPDFはD列以降の月列（例: 2026年2月）にハイパーリンクとして記録される。
 * 月列は日付昇順（左が古い）に自動整列される。
 */
function updatePdfLinks() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    SpreadsheetApp.getUi().alert("エラー", `「${SHEET_NAME}」シートが見つかりません。`, SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }

  // 設定シートからPDF保存先フォルダURLを取得
  const rootFolderUrl = _getSettingValue_("PDF保存先フォルダ");
  if (!rootFolderUrl || !rootFolderUrl.startsWith("https://drive.google.com/")) {
    SpreadsheetApp.getUi().alert(
      "エラー",
      "「設定」シートの「PDF保存先フォルダ」にGoogle DriveのフォルダURLが設定されていません。\n" +
      "例: https://drive.google.com/drive/folders/XXXX",
      SpreadsheetApp.getUi().ButtonSet.OK
    );
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
          _collectSoftBankPdfs_(sub, pdfEntries);
        }
      }
      // 旧形式: 月フォルダ直下（後方互換）
      _collectSoftBankPdfs_(monthFolder, pdfEntries);
    }
  }

  if (pdfEntries.length === 0) {
    SpreadsheetApp.getUi().alert("情報", "SoftBankのPDFファイルが見つかりませんでした。\nGoogle Driveの同期が完了しているか確認してください。", SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }

  // 必要な月列を確保（右→左の順で挿入してインデックスのずれを防ぐ）
  const neededHeaders = [...new Set(pdfEntries.map(e => `${e.year}年${parseInt(e.month)}月`))];
  neededHeaders.sort((a, b) => _monthHeaderToNum_(b) - _monthHeaderToNum_(a)); // 降順（右→左に挿入）
  for (const header of neededHeaders) {
    _ensureMonthColumn_(sheet, header);
  }

  // ハイパーリンクを書き込む（同月複数PDFは2件目以降をセルのメモに追記）
  let updatedCount = 0;
  const processed = new Set();
  for (const entry of pdfEntries) {
    const rowNum = phoneToRow[entry.phone];
    if (!rowNum) continue;

    const colNum = _getMonthColumnNum_(sheet, entry.year, entry.month);
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


// ────────────────────────────────────────────────
//  内部ヘルパー関数
// ────────────────────────────────────────────────

/**
 * 設定シートから指定キーの値を返す。見つからなければ null。
 */
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
 * フォルダ内のSoftBank命名規則PDFを収集する
 * ファイル名形式: YYYYMM_SoftBank_{phone}_*.pdf
 */
function _collectSoftBankPdfs_(folder, results) {
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
function _monthHeaderToNum_(header) {
  const m = String(header).match(/^(\d{4})年(\d+)月$/);
  if (!m) return 0;
  return parseInt(m[1]) * 100 + parseInt(m[2]);
}


/**
 * 月列が存在しなければ日付順の位置に挿入する
 */
function _ensureMonthColumn_(sheet, monthHeader) {
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  if (headers.includes(monthHeader)) return; // 既に存在

  const targetNum = _monthHeaderToNum_(monthHeader);
  let insertAt = sheet.getLastColumn() + 1; // デフォルトは末尾

  for (let i = MONTH_COL_START - 1; i < headers.length; i++) {
    const colNum = _monthHeaderToNum_(String(headers[i]));
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
function _getMonthColumnNum_(sheet, year, month) {
  const header = `${year}年${parseInt(month)}月`;
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const idx = headers.indexOf(header);
  return idx === -1 ? null : idx + 1;
}


// ────────────────────────────────────────────────
//  メニュー
// ────────────────────────────────────────────────

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
    .addToUi();
}
