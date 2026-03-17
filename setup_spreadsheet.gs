/**
 * 認証情報管理シート セットアップスクリプト
 *
 * 【スプレッドシート構成】
 *   認証情報管理シート（このスプレッドシート）:
 *     設定           ... PDF保存先・パスワード・対象月・回線管理スプシURL
 *     認証情報       ... ダウンロード対象の電話番号（サイドバーで管理）
 *     SoftBankリンク ... ダウンロード済みPDFへのリンク一覧
 *     Ymobileリンク  ... 同上
 *   回線管理スプレッドシート（別スプレッドシート）:
 *     月別シート     ... 電話番号・解約済・運用端末等の管理データ
 *
 * 【使い方】
 *   1. このコードをApps Scriptに貼り付けて保存
 *   2. setupSheet() を実行（初回は権限承認が必要）
 *   3. 設定シートに回線管理スプシURL・PDF保存先・パスワードを入力
 *   4. メニュー「ダウンロード対象の電話番号を管理」で対象番号を選択・保存
 *   5. Pythonスクリプトを実行してPDFをダウンロード
 *   6. メニュー「PDFリンクを更新」でリンクシートに反映
 */

// ────────────────────────────────────────────────
//  定数
// ────────────────────────────────────────────────

const SETTINGS_SHEET_NAME = "設定";
const AUTH_SHEET_NAME = "認証情報";
const SOFTBANK_LINK_SHEET_NAME = "SoftBankリンク";
const YMOBILE_LINK_SHEET_NAME = "Ymobileリンク";

const MONTH_COL_START = 3; // リンクシートの月列開始位置（C列）

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

function setupSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  setupSettingsSheet_(ss);
  setupAuthSheet_(ss);
  setupLinkSheet_(ss, SOFTBANK_LINK_SHEET_NAME, "SoftBank");
  setupLinkSheet_(ss, YMOBILE_LINK_SHEET_NAME, "Ymobile");

  // 旧シートの削除
  for (const name of ["_PhoneDropdown", "回線管理表"]) {
    const old = ss.getSheetByName(name);
    if (old) ss.deleteSheet(old);
  }

  SpreadsheetApp.getUi().alert(
    "セットアップ完了",
    "設定・認証情報・リンクシートを作成しました。\n\n" +
    "1. 設定シートに「回線管理スプレッドシート」のURLを入力\n" +
    "2. 設定シートに「パスワード」を入力\n" +
    "3. メニュー「ダウンロード対象の電話番号を管理」で対象を選択",
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}


function setupSettingsSheet_(ss) {
  let sheet = ss.getSheetByName(SETTINGS_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SETTINGS_SHEET_NAME);
    ss.setActiveSheet(sheet);
    ss.moveActiveSheet(1);
  }
  if (!sheet.getRange(1, 1).getValue()) {
    sheet.getRange(1, 1, 1, 2).setValues([["設定名", "値"]])
      .setFontWeight("bold").setBackground("#FF6D00").setFontColor("#FFFFFF").setHorizontalAlignment("center");
  }
  sheet.setColumnWidth(1, 200);
  sheet.setColumnWidth(2, 500);

  _upsertSettingRow_(sheet, "このスプレッドシートURL", ss.getUrl(),
    "Pythonの SPREADSHEET_URL にコピーしてください。");
  _upsertSettingRow_(sheet, "回線管理スプレッドシート", "",
    "回線管理表のURL（月別シートに電話番号・解約済・運用端末がある別スプレッドシート）。");
  _upsertSettingRow_(sheet, "PDF保存先フォルダ", "https://drive.google.com/drive/folders/XXXXXXXX",
    "Google DriveのフォルダURL、またはローカル絶対パス。");
  _upsertSettingRow_(sheet, "パスワード", "",
    "SoftBank / Y!mobile 共通のログインパスワード。");
  _upsertSettingRow_(sheet, "対象月", "自動（前月）",
    "ダウンロードする月。「自動（前月）」= 実行時の前月。");
  _setTargetMonthValidation_(sheet);
}


/**
 * 認証情報シート: 電話番号 | キャリア | PDFの種類 | 運用端末
 * - サイドバーから選択した電話番号が書き込まれる
 * - パスワードは設定シートで一元管理（この列には持たない）
 * - 運用端末はサイドバー保存時に回線管理表から自動設定
 */
function setupAuthSheet_(ss) {
  let sheet = ss.getSheetByName(AUTH_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(AUTH_SHEET_NAME);
    ss.setActiveSheet(sheet);
    ss.moveActiveSheet(2);
  }

  // ヘッダーを正しい4列に設定（旧形式からの移行対応）
  const correctHeaders = ["電話番号", "キャリア", "PDFの種類", "運用端末"];
  const currentHeaders = sheet.getRange(1, 1, 1, Math.max(sheet.getLastColumn(), 1)).getValues()[0]
    .map(h => String(h || "").trim());

  // パスワード列が残っていたら削除
  const pwIdx = currentHeaders.indexOf("パスワード");
  if (pwIdx !== -1) {
    sheet.deleteColumn(pwIdx + 1);
  }

  // ヘッダーが正しくなければ上書き
  const h0 = sheet.getRange(1, 1, 1, Math.max(sheet.getLastColumn(), 1)).getValues()[0]
    .map(v => String(v || "").trim());
  if (h0.length < 4 || h0[0] !== "電話番号" || h0[1] !== "キャリア" || h0[3] !== "運用端末") {
    sheet.getRange(1, 1, 1, 4).setValues([correctHeaders])
      .setFontWeight("bold").setBackground("#4285F4").setFontColor("#FFFFFF").setHorizontalAlignment("center");
  }

  sheet.setColumnWidth(1, 180);
  sheet.setColumnWidth(2, 120);
  sheet.setColumnWidth(3, 220);
  sheet.setColumnWidth(4, 160);
  sheet.getRange("A:A").setNumberFormat("@");
}


/**
 * リンクシート: 電話番号 | PDFの種類 | 月列...
 * - 認証情報シートと連動（サイドバー保存時に電話番号を同期）
 * - 「PDFリンクを更新」でDriveのPDFへのハイパーリンクが書き込まれる
 */
function setupLinkSheet_(ss, sheetName, carrierLabel) {
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) sheet = ss.insertSheet(sheetName);
  if (!sheet.getRange(1, 1).getValue()) {
    const color = carrierLabel === "SoftBank" ? "#4285F4" : "#FF6D00";
    sheet.getRange(1, 1, 1, 2).setValues([["電話番号", "名義"]])
      .setFontWeight("bold").setBackground(color).setFontColor("#FFFFFF").setHorizontalAlignment("center");
  }
  sheet.setColumnWidth(1, 180);
  sheet.setColumnWidth(2, 220);
  sheet.getRange("A:A").setNumberFormat("@");
}


// ────────────────────────────────────────────────
//  電話番号管理 HTMLサイドバー
// ────────────────────────────────────────────────

function openPhoneManagerSidebar() {
  const html = HtmlService.createHtmlOutput(_getPhoneManagerHtml_())
    .setTitle("ダウンロード対象の電話番号").setWidth(320);
  SpreadsheetApp.getUi().showSidebar(html);
}

function _getPhoneManagerHtml_() {
  return `<!DOCTYPE html>
<html>
<head>
<style>
  * { box-sizing: border-box; margin: 0; padding: 0; }
  body { font-family: "Google Sans", Arial, sans-serif; font-size: 13px; color: #333; padding: 12px; }
  .tab-bar { display: flex; gap: 4px; margin-bottom: 12px; }
  .tab-btn {
    flex: 1; padding: 8px; border: none; border-radius: 6px 6px 0 0;
    cursor: pointer; font-size: 13px; font-weight: bold;
    background: #e0e0e0; color: #666; transition: all 0.2s;
  }
  .tab-btn.active { color: #fff; }
  .tab-btn[data-carrier="SoftBank"].active { background: #4285F4; }
  .tab-btn[data-carrier="Ymobile"].active { background: #FF6D00; }
  .phone-list { max-height: 55vh; overflow-y: auto; border: 1px solid #ddd; border-radius: 4px; padding: 4px; }
  .phone-item {
    display: flex; align-items: center; gap: 8px; padding: 6px 8px;
    border-bottom: 1px solid #f0f0f0;
  }
  .phone-item:last-child { border-bottom: none; }
  .phone-item:hover { background: #f5f5f5; }
  .phone-item input[type="checkbox"] { width: 16px; height: 16px; cursor: pointer; }
  .phone-number { flex: 1; font-family: monospace; font-size: 13px; }
  .phone-item select { font-size: 11px; padding: 2px; max-width: 110px; }
  .cancelled { color: #999; text-decoration: line-through; }
  .status { margin: 8px 0; padding: 8px; border-radius: 4px; font-size: 12px; display: none; }
  .status.success { display: block; background: #e8f5e9; color: #2e7d32; }
  .status.error { display: block; background: #ffebee; color: #c62828; }
  .status.loading { display: block; background: #e3f2fd; color: #1565c0; }
  .status.warn { display: block; background: #fff3e0; color: #e65100; }
  .btn-bar { margin-top: 12px; display: flex; gap: 8px; }
  .btn { flex: 1; padding: 10px; border: none; border-radius: 6px; font-size: 13px; font-weight: bold; cursor: pointer; }
  .btn:disabled { opacity: 0.5; cursor: not-allowed; }
  .btn-save { background: #4285F4; color: #fff; }
  .btn-all { background: #e0e0e0; color: #333; font-size: 11px; }
  .empty-msg { padding: 20px; text-align: center; color: #999; }
  .summary { font-size: 11px; color: #666; margin-top: 4px; }
</style>
</head>
<body>
  <div class="tab-bar">
    <button class="tab-btn active" data-carrier="SoftBank" onclick="switchTab('SoftBank')">SoftBank</button>
    <button class="tab-btn" data-carrier="Ymobile" onclick="switchTab('Ymobile')">Y!mobile</button>
  </div>
  <div id="status" class="status"></div>
  <div id="phoneList" class="phone-list"><div class="empty-msg">読み込み中...</div></div>
  <div id="summary" class="summary"></div>
  <div class="btn-bar">
    <button class="btn btn-all" onclick="toggleAll()">全選択/解除</button>
    <button class="btn btn-save" id="saveBtn" onclick="save()">保存</button>
  </div>

<script>
  const PDF_TYPES = ${JSON.stringify(PDF_TYPE_OPTIONS)};
  let currentCarrier = "SoftBank";
  let phoneData = {};
  let selections = {};

  function setStatus(msg, type) {
    var el = document.getElementById("status");
    el.textContent = msg; el.className = "status " + type;
  }
  function clearStatus() { document.getElementById("status").className = "status"; }

  function init() {
    setStatus("回線管理表から電話番号を読み込み中...", "loading");
    google.script.run
      .withSuccessHandler(function(r) {
        phoneData = r.phones;
        selections = r.selections;
        // 解約済番号を選択から自動除外
        var removed = [];
        ["SoftBank", "Ymobile"].forEach(function(c) {
          var phones = phoneData[c] || [];
          var sel = selections[c] || {};
          Object.keys(sel).forEach(function(ph) {
            var info = phones.find(function(p) { return p.phone === ph; });
            if (!info || info.cancelled) {
              delete sel[ph];
              removed.push(ph);
            }
          });
        });
        if (removed.length > 0) {
          setStatus("解約済 " + removed.length + "件を選択から除外しました", "warn");
        } else {
          clearStatus();
        }
        render();
      })
      .withFailureHandler(function(err) { setStatus("読み込みエラー: " + err.message, "error"); })
      .getPhoneManagerData();
  }

  function switchTab(carrier) {
    currentCarrier = carrier;
    document.querySelectorAll(".tab-btn").forEach(function(b) {
      b.classList.toggle("active", b.dataset.carrier === carrier);
    });
    render();
  }

  function render() {
    var list = document.getElementById("phoneList");
    var phones = phoneData[currentCarrier] || [];
    var sel = selections[currentCarrier] || {};

    if (phones.length === 0) {
      list.innerHTML = '<div class="empty-msg">' + currentCarrier + ' の電話番号が見つかりません。<br>設定シートの「回線管理スプレッドシート」URLを確認してください。</div>';
      updateSummary(); return;
    }

    var html = "";
    phones.forEach(function(p) {
      var checked = sel[p.phone] ? "checked" : "";
      var pdfType = (sel[p.phone] && sel[p.phone].pdfType) || "電話番号別";
      var cls = p.cancelled ? "cancelled" : "";
      var label = p.device ? p.phone + " (" + p.device + ")" : p.phone;
      if (p.cancelled) label += " [解約済]";

      html += '<div class="phone-item">'
        + '<input type="checkbox" data-phone="' + p.phone + '" ' + checked
        + (p.cancelled ? " disabled" : "") + ' onchange="onCheck(this)">'
        + '<span class="phone-number ' + cls + '">' + label + '</span>';
      if (!p.cancelled) {
        html += '<select data-phone="' + p.phone + '" onchange="onPdfType(this)">';
        PDF_TYPES.forEach(function(t) {
          html += '<option' + (t === pdfType ? ' selected' : '') + '>' + t + '</option>';
        });
        html += '</select>';
      }
      html += '</div>';
    });
    list.innerHTML = html;
    updateSummary();
  }

  function onCheck(el) {
    if (!selections[currentCarrier]) selections[currentCarrier] = {};
    if (el.checked) {
      var selEl = el.parentElement.querySelector("select");
      selections[currentCarrier][el.dataset.phone] = { pdfType: selEl ? selEl.value : "電話番号別" };
    } else {
      delete selections[currentCarrier][el.dataset.phone];
    }
    updateSummary();
  }

  function onPdfType(el) {
    if (selections[currentCarrier] && selections[currentCarrier][el.dataset.phone]) {
      selections[currentCarrier][el.dataset.phone].pdfType = el.value;
    }
  }

  function toggleAll() {
    var phones = phoneData[currentCarrier] || [];
    var sel = selections[currentCarrier] || {};
    var active = phones.filter(function(p) { return !p.cancelled; });
    var allChecked = active.every(function(p) { return sel[p.phone]; });
    if (!selections[currentCarrier]) selections[currentCarrier] = {};
    if (allChecked) {
      active.forEach(function(p) { delete selections[currentCarrier][p.phone]; });
    } else {
      active.forEach(function(p) {
        if (!selections[currentCarrier][p.phone])
          selections[currentCarrier][p.phone] = { pdfType: "電話番号別" };
      });
    }
    render();
  }

  function updateSummary() {
    document.getElementById("summary").textContent =
      "選択中: SoftBank " + Object.keys(selections["SoftBank"] || {}).length +
      "件 / Ymobile " + Object.keys(selections["Ymobile"] || {}).length + "件";
  }

  function save() {
    var btn = document.getElementById("saveBtn");
    btn.disabled = true;
    setStatus("保存中...", "loading");
    google.script.run
      .withSuccessHandler(function(msg) { setStatus(msg, "success"); btn.disabled = false; })
      .withFailureHandler(function(err) { setStatus("保存エラー: " + err.message, "error"); btn.disabled = false; })
      .savePhoneSelections(selections);
  }

  init();
</script>
</body>
</html>`;
}


// ── サーバー側: サイドバーAPI ──

function getPhoneManagerData() {
  return {
    phones: _getAllPhonesFromMonthSheets_(),
    selections: _getCurrentSelections_(),
  };
}


/**
 * サイドバーからの保存。
 * 認証情報シートとリンクシートの両方を更新する。
 * 解約済の番号は自動的に除外される。
 */
function savePhoneSelections(selections) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const allPhones = _getAllPhonesFromMonthSheets_();

  // 解約済番号セットを作成
  const cancelledSet = new Set();
  for (const carrier of ["SoftBank", "Ymobile"]) {
    for (const p of (allPhones[carrier] || [])) {
      if (p.cancelled) cancelledSet.add(p.phone);
    }
  }

  // 運用端末・名義マップ
  const deviceMap = {}, nameMap = {};
  for (const carrier of ["SoftBank", "Ymobile"]) {
    for (const p of (allPhones[carrier] || [])) {
      if (p.device) deviceMap[p.phone] = p.device;
      if (p.name) nameMap[p.phone] = p.name;
    }
  }

  // ── 認証情報シート ──
  let authSheet = ss.getSheetByName(AUTH_SHEET_NAME);
  if (!authSheet) {
    authSheet = ss.insertSheet(AUTH_SHEET_NAME);
    setupAuthSheet_(ss);
  }

  if (authSheet.getLastRow() > 1) {
    authSheet.getRange(2, 1, authSheet.getLastRow() - 1, authSheet.getLastColumn()).clearContent();
  }

  // 書き込み（解約済を除外）
  const authRows = [];
  const linkData = { SoftBank: [], Ymobile: [] };
  for (const carrier of ["SoftBank", "Ymobile"]) {
    const sel = selections[carrier] || {};
    for (const phone of Object.keys(sel)) {
      if (cancelledSet.has(phone)) continue;
      const pdfType = sel[phone].pdfType || "電話番号別";
      authRows.push([phone, carrier, pdfType, deviceMap[phone] || ""]);
      linkData[carrier].push({ phone, name: nameMap[phone] || "" });
    }
  }
  if (authRows.length > 0) {
    authSheet.getRange(2, 1, authRows.length, 4).setValues(authRows);
  }
  authSheet.getRange("A:A").setNumberFormat("@");

  // ── リンクシートの電話番号を同期 ──
  _syncLinkSheetPhones_(ss, SOFTBANK_LINK_SHEET_NAME, linkData.SoftBank);
  _syncLinkSheetPhones_(ss, YMOBILE_LINK_SHEET_NAME, linkData.Ymobile);

  const skipped = Object.values(selections).reduce((n, s) =>
    n + Object.keys(s).filter(ph => cancelledSet.has(ph)).length, 0);
  let msg = `保存しました（${authRows.length}件）。`;
  if (skipped > 0) msg += `\n解約済 ${skipped}件を除外しました。`;
  return msg;
}


/**
 * リンクシートの電話番号行を認証情報と同期する。
 * 既存のリンク（月列のデータ）は保持しつつ、電話番号行の追加・削除を行う。
 */
function _syncLinkSheetPhones_(ss, sheetName, phoneList) {
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) return;

  // 現在のリンクシートの電話番号を取得
  const data = sheet.getDataRange().getValues();
  const existingPhones = {};
  for (let i = 1; i < data.length; i++) {
    const phone = String(data[i][0] || "").replace(/[-\s]/g, "").trim();
    if (phone) existingPhones[phone] = i + 1; // 行番号
  }

  const targetPhones = new Set(phoneList.map(p => p.phone));

  // 不要な行を削除（降順で削除して行番号ズレを防ぐ）
  const toDelete = Object.entries(existingPhones)
    .filter(([ph]) => !targetPhones.has(ph))
    .sort((a, b) => b[1] - a[1]);
  for (const [, rowNum] of toDelete) {
    sheet.deleteRow(rowNum);
  }

  // 追加が必要な番号 + 既存行の名義更新
  const currentPhones = new Set(Object.keys(existingPhones).filter(ph => targetPhones.has(ph)));
  const nameByPhone = {};
  for (const p of phoneList) nameByPhone[p.phone] = p.name || "";

  for (const { phone, name } of phoneList) {
    if (!currentPhones.has(phone)) {
      const newRow = sheet.getLastRow() + 1;
      sheet.getRange(newRow, 1).setNumberFormat("@").setValue(phone);
      sheet.getRange(newRow, 2).setValue(name);
    } else {
      // 既存行の名義を更新
      sheet.getRange(existingPhones[phone], 2).setValue(name);
    }
  }
}


// ── 回線管理表の読み込み ──

function _normalizeCarrierName_(text) {
  const s = String(text).trim().toLowerCase();
  if (!s) return null;
  if (s === "softbank" || s.includes("ソフトバンク")) return "SoftBank";
  if (s === "ymobile" || s === "y!mobile" || s.includes("ワイモバイル")) return "Ymobile";
  return null;
}

/**
 * 回線管理スプレッドシートの月別シートから電話番号を収集する。
 * セクション分け構造に対応（キャリアラベル行 → ヘッダー行 → データ行）。
 */
function _getAllPhonesFromMonthSheets_() {
  const result = { SoftBank: [], Ymobile: [] };

  let ss;
  const mgmtUrl = _getSettingValue_("回線管理スプレッドシート");
  if (mgmtUrl) {
    const idMatch = mgmtUrl.match(/\/spreadsheets\/d\/([a-zA-Z0-9_-]+)/);
    if (!idMatch) return result;
    try { ss = SpreadsheetApp.openById(idMatch[1]); } catch (e) { return result; }
  } else {
    ss = SpreadsheetApp.getActiveSpreadsheet();
  }

  const monthSheets = ss.getSheets().filter(ws => /.*\d+月$/.test(ws.getName()));
  if (monthSheets.length === 0) return result;

  monthSheets.sort((a, b) => _parseMonthSheetNum_(a.getName()) - _parseMonthSheetNum_(b.getName()));
  let targetSheet = null;
  for (let i = monthSheets.length - 1; i >= 0; i--) {
    if (monthSheets[i].getLastRow() > 1) { targetSheet = monthSheets[i]; break; }
  }
  if (!targetSheet) return result;

  const rows = targetSheet.getDataRange().getValues();
  let cols = null, sectionCarrier = null;

  for (let i = 0; i < rows.length; i++) {
    const row = rows[i];

    // セクションラベル行
    const labelCarrier = _detectCarrierLabel_(row);
    if (labelCarrier) { sectionCarrier = labelCarrier; cols = null; continue; }

    // ヘッダー行
    const phoneColIdx = row.findIndex(c => String(c).trim() === "電話番号");
    if (phoneColIdx !== -1) {
      cols = {};
      row.forEach((c, j) => { const k = String(c).trim(); if (k) cols[k] = j; });
      continue;
    }

    if (!cols || cols["電話番号"] === undefined) continue;

    const phone = String(row[cols["電話番号"]] || "").replace(/[-\s]/g, "").trim();
    if (!phone || !/^\d{10,13}$/.test(phone)) continue;

    let carrier = cols["キャリア"] !== undefined ? _normalizeCarrierName_(row[cols["キャリア"]]) : null;
    if (!carrier) carrier = sectionCarrier;
    if (!carrier || !result[carrier]) continue;

    const ci = cols["解約済"];
    const cancelled = ci !== undefined && String(row[ci] || "").toUpperCase() === "TRUE";
    const di = cols["運用端末"];
    const device = di !== undefined ? String(row[di] || "").trim() : "";
    const ni = cols["名義"] !== undefined ? cols["名義"] : cols["契約者名"];
    const name = ni !== undefined ? String(row[ni] || "").trim() : "";

    if (!result[carrier].some(p => p.phone === phone)) {
      result[carrier].push({ phone, cancelled, device, name });
    }
  }
  return result;
}

function _detectCarrierLabel_(row) {
  const nonEmpty = row.filter(c => String(c).trim() !== "");
  if (nonEmpty.length === 0 || nonEmpty.length > 3) return null;
  if (nonEmpty.some(c => String(c).trim() === "電話番号")) return null;
  for (const cell of nonEmpty) {
    const c = _normalizeCarrierName_(cell);
    if (c) return c;
  }
  return null;
}

function _getCurrentSelections_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(AUTH_SHEET_NAME);
  const result = { SoftBank: {}, Ymobile: {} };
  if (!sheet || sheet.getLastRow() <= 1) return result;

  const data = sheet.getDataRange().getValues();
  const h = data[0].map(v => String(v).trim());
  const pi = h.indexOf("電話番号"), ci = h.indexOf("キャリア"), ti = h.indexOf("PDFの種類");
  if (pi === -1) return result;

  for (let i = 1; i < data.length; i++) {
    const phone = String(data[i][pi] || "").replace(/[-\s]/g, "").trim();
    const carrier = ci !== -1 ? String(data[i][ci] || "").trim() : "";
    const pdfType = ti !== -1 ? String(data[i][ti] || "").trim() : "電話番号別";
    if (phone && carrier && result[carrier]) {
      result[carrier][phone] = { pdfType: pdfType || "電話番号別" };
    }
  }
  return result;
}

function _parseMonthSheetNum_(name) {
  const m1 = name.match(/\((\d{4})\)(\d+)月/);
  if (m1) return parseInt(m1[1]) * 100 + parseInt(m1[2]);
  const m2 = name.match(/(\d+)月$/);
  if (m2) return new Date().getFullYear() * 100 + parseInt(m2[1]);
  return 0;
}


// ────────────────────────────────────────────────
//  PDFリンク更新
// ────────────────────────────────────────────────

function updateSoftBankLinks() { _updatePdfLinks_(SOFTBANK_LINK_SHEET_NAME, "SoftBank"); }
function updateYmobileLinks() { _updatePdfLinks_(YMOBILE_LINK_SHEET_NAME, "Ymobile"); }

/**
 * Google Driveを探索してリンクシートにPDFリンクを設定する。
 * リンクシートに電話番号がない場合でも、Driveで見つかったPDFの電話番号を自動追加する。
 */
function _updatePdfLinks_(sheetName, carrier) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    SpreadsheetApp.getUi().alert("エラー", `「${sheetName}」が見つかりません。setupSheet()を実行してください。`, SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }

  const rootFolderUrl = _getSettingValue_("PDF保存先フォルダ");
  if (!rootFolderUrl || !rootFolderUrl.startsWith("https://drive.google.com/")) {
    SpreadsheetApp.getUi().alert("エラー", "設定シートの「PDF保存先フォルダ」にDriveのURLを設定してください。", SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }

  const folderMatch = rootFolderUrl.match(/\/folders\/([a-zA-Z0-9_-]+)/);
  if (!folderMatch) { SpreadsheetApp.getUi().alert("エラー", "フォルダURLからIDを取得できません。", SpreadsheetApp.getUi().ButtonSet.OK); return; }
  let rootFolder;
  try { rootFolder = DriveApp.getFolderById(folderMatch[1]); }
  catch (e) { SpreadsheetApp.getUi().alert("エラー", `フォルダアクセス失敗: ${e.message}`, SpreadsheetApp.getUi().ButtonSet.OK); return; }

  // PDFを収集
  const pdfEntries = [];
  const yearFolders = rootFolder.getFolders();
  while (yearFolders.hasNext()) {
    const yf = yearFolders.next();
    if (!/^\d{4}$/.test(yf.getName())) continue;
    const monthFolders = yf.getFolders();
    while (monthFolders.hasNext()) {
      const mf = monthFolders.next();
      if (!/^\d{2}$/.test(mf.getName())) continue;
      const subFolders = mf.getFolders();
      while (subFolders.hasNext()) {
        const sub = subFolders.next();
        if (sub.getName() === carrier) _collectCarrierPdfs_(sub, carrier, pdfEntries);
      }
      _collectCarrierPdfs_(mf, carrier, pdfEntries);
    }
  }

  if (pdfEntries.length === 0) {
    SpreadsheetApp.getUi().alert("情報", `${carrier}のPDFが見つかりません。`, SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }

  // 電話番号→行番号マップ（なければ自動追加）
  let data = sheet.getDataRange().getValues();
  const phoneToRow = {};
  for (let i = 1; i < data.length; i++) {
    const ph = String(data[i][0] || "").replace(/[-\s]/g, "").trim();
    if (ph) phoneToRow[ph] = i + 1;
  }

  // PDFに含まれるがリンクシートにない電話番号を自動追加（名義は認証情報から取得不可のため空欄）
  const newPhones = [...new Set(pdfEntries.map(e => e.phone))].filter(ph => !phoneToRow[ph]);
  for (const ph of newPhones) {
    const row = sheet.getLastRow() + 1;
    sheet.getRange(row, 1).setNumberFormat("@").setValue(ph);
    sheet.getRange(row, 2).setValue("");
    phoneToRow[ph] = row;
  }

  // 月列を確保
  const neededHeaders = [...new Set(pdfEntries.map(e => `${e.year}年${parseInt(e.month)}月`))];
  neededHeaders.sort((a, b) => _monthHeaderToNum_(b) - _monthHeaderToNum_(a));
  for (const h of neededHeaders) _ensureMonthColumn_(sheet, h);

  // リンク書き込み
  let updated = 0;
  const processed = new Set();
  for (const entry of pdfEntries) {
    const rowNum = phoneToRow[entry.phone];
    const colNum = _getMonthColumnNum_(sheet, entry.year, entry.month);
    if (!rowNum || !colNum) continue;

    const key = `${entry.phone}_${entry.year}${entry.month}`;
    if (processed.has(key)) {
      const cell = sheet.getRange(rowNum, colNum);
      cell.setNote((cell.getNote() || "") + (cell.getNote() ? "\n" : "") + entry.file.getName());
      continue;
    }
    processed.add(key);

    const url = entry.file.getUrl();
    const am = entry.file.getName().match(/_(\d+)円\.pdf$/);
    const label = am ? `${parseInt(entry.month)}月 ${Number(am[1]).toLocaleString()}円` : `${parseInt(entry.month)}月`;
    sheet.getRange(rowNum, colNum).setFormula(`=HYPERLINK("${url}","${label}")`);
    updated++;
  }

  SpreadsheetApp.getUi().alert("完了",
    `${pdfEntries.length}件のPDFを確認、${updated}件のリンクを設定しました。` +
    (newPhones.length > 0 ? `\n${newPhones.length}件の電話番号を自動追加しました。` : ""),
    SpreadsheetApp.getUi().ButtonSet.OK);
}


// ────────────────────────────────────────────────
//  PDFから金額を取得・ファイル名更新
// ────────────────────────────────────────────────

function scanAndUpdateSoftBankAmounts() { _scanAndUpdatePdfAmounts_(SOFTBANK_LINK_SHEET_NAME, "SoftBank"); }
function scanAndUpdateYmobileAmounts() { _scanAndUpdatePdfAmounts_(YMOBILE_LINK_SHEET_NAME, "Ymobile"); }

function _scanAndUpdatePdfAmounts_(sheetName, carrier) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) { SpreadsheetApp.getUi().alert("エラー", `「${sheetName}」が見つかりません。`, SpreadsheetApp.getUi().ButtonSet.OK); return; }

  const rootFolderUrl = _getSettingValue_("PDF保存先フォルダ");
  if (!rootFolderUrl || !rootFolderUrl.startsWith("https://drive.google.com/")) {
    SpreadsheetApp.getUi().alert("エラー", "PDF保存先フォルダが未設定です。", SpreadsheetApp.getUi().ButtonSet.OK); return;
  }
  const folderMatch = rootFolderUrl.match(/\/folders\/([a-zA-Z0-9_-]+)/);
  if (!folderMatch) { SpreadsheetApp.getUi().alert("エラー", "フォルダIDを取得できません。", SpreadsheetApp.getUi().ButtonSet.OK); return; }
  let rootFolder;
  try { rootFolder = DriveApp.getFolderById(folderMatch[1]); }
  catch (e) { SpreadsheetApp.getUi().alert("エラー", `フォルダアクセス失敗: ${e.message}`, SpreadsheetApp.getUi().ButtonSet.OK); return; }

  const pdfEntries = [];
  const yfs = rootFolder.getFolders();
  while (yfs.hasNext()) { const yf = yfs.next(); if (!/^\d{4}$/.test(yf.getName())) continue;
    const mfs = yf.getFolders(); while (mfs.hasNext()) { const mf = mfs.next(); if (!/^\d{2}$/.test(mf.getName())) continue;
      const sfs = mf.getFolders(); while (sfs.hasNext()) { const sf = sfs.next(); if (sf.getName() === carrier) _collectCarrierPdfs_(sf, carrier, pdfEntries); }
      _collectCarrierPdfs_(mf, carrier, pdfEntries);
  }}

  const targets = pdfEntries.filter(e => e.file.getName().endsWith("_利用料金明細.pdf"));
  if (targets.length === 0) { SpreadsheetApp.getUi().alert("情報", "金額が未取得のPDFはありませんでした。", SpreadsheetApp.getUi().ButtonSet.OK); return; }

  const data = sheet.getDataRange().getValues();
  const phoneToRow = {};
  for (let i = 1; i < data.length; i++) {
    const ph = String(data[i][0] || "").replace(/[-\s]/g, "").trim();
    if (ph) phoneToRow[ph] = i + 1;
  }

  let updated = 0, failed = 0;
  for (const entry of targets) {
    const amount = _extractAmountFromPdf_(entry.file);
    if (!amount) { failed++; continue; }
    entry.file.setName(entry.file.getName().replace("_利用料金明細.pdf", `_${amount}円.pdf`));
    const rowNum = phoneToRow[entry.phone];
    if (rowNum) {
      const colNum = _getMonthColumnNum_(sheet, entry.year, entry.month);
      if (colNum) sheet.getRange(rowNum, colNum).setFormula(`=HYPERLINK("${entry.file.getUrl()}","${parseInt(entry.month)}月 ${Number(amount).toLocaleString()}円")`);
    }
    updated++;
  }
  SpreadsheetApp.getUi().alert("完了", `${targets.length}件処理: 更新${updated}件、失敗${failed}件`, SpreadsheetApp.getUi().ButtonSet.OK);
}


// ────────────────────────────────────────────────
//  内部ヘルパー
// ────────────────────────────────────────────────

function _getSettingValue_(key) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SETTINGS_SHEET_NAME);
  if (!sheet) return null;
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === key) return String(data[i][1]).trim() || null;
  }
  return null;
}

function _collectCarrierPdfs_(folder, carrier, results) {
  const files = folder.getFiles();
  while (files.hasNext()) {
    const file = files.next();
    if (file.getMimeType() !== "application/pdf") continue;
    const m = file.getName().match(new RegExp(`^(\\d{4})(\\d{2})_${carrier}_(\\d+)`));
    if (m) results.push({ year: m[1], month: m[2], phone: m[3], file });
  }
}

function _monthHeaderToNum_(h) {
  const m = String(h).match(/^(\d{4})年(\d+)月$/);
  return m ? parseInt(m[1]) * 100 + parseInt(m[2]) : 0;
}

function _ensureMonthColumn_(sheet, monthHeader) {
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  if (headers.some(h => String(h).trim() === monthHeader)) return;
  const targetNum = _monthHeaderToNum_(monthHeader);
  let insertAt = sheet.getLastColumn() + 1;
  for (let i = MONTH_COL_START - 1; i < headers.length; i++) {
    const n = _monthHeaderToNum_(String(headers[i]));
    if (n > 0 && targetNum < n) { insertAt = i + 1; sheet.insertColumnBefore(insertAt); break; }
  }
  sheet.getRange(1, insertAt).setNumberFormat("@").setValue(monthHeader)
    .setFontWeight("bold").setBackground("#34A853").setFontColor("#FFFFFF").setHorizontalAlignment("center");
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
  for (let i = 1; i < data.length; i++) { if (String(data[i][0]).trim() === key) return; }
  const r = sheet.getLastRow() + 1;
  sheet.getRange(r, 1).setValue(key);
  const v = sheet.getRange(r, 2);
  v.setValue(defaultValue);
  if (noteText) v.setNote(noteText);
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
    const cell = sheet.getRange(i + 1, 2);
    cell.setDataValidation(SpreadsheetApp.newDataValidation()
      .requireValueInList(options, true).setAllowInvalid(false).build());
    if (!cell.getValue()) cell.setValue("自動（前月）");
    return;
  }
}

function _extractAmountFromPdf_(file) {
  let docFileId = null;
  try {
    const docFile = Drive.Files.insert({ title: "tmp_ocr_" + file.getName() }, file.getBlob(), { convert: true });
    docFileId = docFile.id;
    const text = DocumentApp.openById(docFileId).getBody().getText();
    for (const pat of [/(?<![小合])計[^\d]*([\d,]+)/, /小計[^\d]*([\d,]+)/]) {
      const m = text.match(pat);
      if (m) { const a = m[1].replace(/,/g, ""); if (/^\d+$/.test(a)) return a; }
    }
    return null;
  } catch (e) { return null; }
  finally { if (docFileId) try { Drive.Files.remove(docFileId); } catch (_) {} }
}


// ────────────────────────────────────────────────
//  メニュー
// ────────────────────────────────────────────────

function onOpen() {
  SpreadsheetApp.getUi().createMenu("携帯領収書管理 ツール")
    .addItem("初期セットアップ", "setupSheet")
    .addSeparator()
    .addItem("ダウンロード対象の電話番号を管理", "openPhoneManagerSidebar")
    .addSeparator()
    .addItem("SoftBank PDFリンクを更新", "updateSoftBankLinks")
    .addItem("Ymobile PDFリンクを更新", "updateYmobileLinks")
    .addSeparator()
    .addItem("SoftBank PDFから金額を取得・ファイル名更新", "scanAndUpdateSoftBankAmounts")
    .addItem("Ymobile PDFから金額を取得・ファイル名更新", "scanAndUpdateYmobileAmounts")
    .addToUi();
}
