/**
 * 認証情報管理シート セットアップスクリプト
 *
 * 【スプレッドシート構成】
 *   認証情報管理シート（このスプレッドシート）:
 *     設定           ... PDF保存先・パスワード・対象月・回線管理スプシURL
 *     認証情報       ... ダウンロード対象の電話番号（サイドバーで管理）
 *     SoftBankリンク ... ダウンロード済みPDFへのリンク一覧
 *     Ymobileリンク  ... 同上
 *     auリンク       ... 同上
 *     UQmobileリンク ... 同上
 *     docomoリンク   ... 同上
 *   回線管理スプレッドシート（別スプレッドシート）:
 *     月別シート     ... 電話番号・解約済・運用端末等の管理データ
 *
 * 【使い方】
 *   1. このコードをApps Scriptに貼り付けて保存
 *   2. setupSheet() を実行（初回は権限承認が必要）
 *   3. 設定シートに回線管理スプシURL・PDF保存先・パスワードを入力
 *   4. メニュー「携帯領収書管理 ツール」→「ダウンロード対象の電話番号を管理」で対象番号を選択・保存
 *   5. Pythonスクリプトを実行してPDFをダウンロード
 *   6. メニュー「携帯領収書管理 ツール」→「PDFリンク」→「全キャリア一括更新」でリンクシートに反映
 */

// ────────────────────────────────────────────────
//  定数
// ────────────────────────────────────────────────

const SETTINGS_SHEET_NAME = "設定";
const AUTH_SHEET_NAME = "認証情報";
const SOFTBANK_LINK_SHEET_NAME = "SoftBankリンク";
const YMOBILE_LINK_SHEET_NAME = "Ymobileリンク";
const AU_LINK_SHEET_NAME = "auリンク";
const UQMOBILE_LINK_SHEET_NAME = "UQmobileリンク";
const DOCOMO_LINK_SHEET_NAME = "docomoリンク";

const HISTORY_SHEET_NAME = "ダウンロード履歴";
const MONTH_COL_START = 3; // リンクシートの月列開始位置（C列）

// SoftBank / Ymobile 用
const PDF_TYPE_OPTIONS = [
  "電話番号別",
  "一括",
  "機種別",
  "電話番号別,一括",
  "電話番号別,機種別",
  "一括,機種別",
  "電話番号別,一括,機種別",
];

// au / UQmobile 用
const AU_PDF_TYPE_OPTIONS = [
  "請求書",
  "領収書",
  "支払証明書",
  "請求書,領収書",
  "請求書,支払証明書",
  "領収書,支払証明書",
  "請求書,領収書,支払証明書",
];

// docomo 用
const DOCOMO_PDF_TYPE_OPTIONS = [
  "一括請求",
  "利用内訳",
  "一括請求,利用内訳",
];


// ────────────────────────────────────────────────
//  初期セットアップ
// ────────────────────────────────────────────────

function setupSheet() {
  _clearSettingsCache_();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  setupSettingsSheet_(ss);
  setupAuthSheet_(ss);
  setupLinkSheet_(ss, SOFTBANK_LINK_SHEET_NAME, "SoftBank");
  setupLinkSheet_(ss, YMOBILE_LINK_SHEET_NAME, "Ymobile");
  setupLinkSheet_(ss, AU_LINK_SHEET_NAME, "au");
  setupLinkSheet_(ss, UQMOBILE_LINK_SHEET_NAME, "UQmobile");
  setupLinkSheet_(ss, DOCOMO_LINK_SHEET_NAME, "docomo");
  setupHistorySheet_(ss);

  SpreadsheetApp.getUi().alert(
    "セットアップ完了",
    "設定・認証情報・リンクシートを作成しました。\n\n" +
    "1. 設定シートに「回線管理スプレッドシート」のURLを入力\n" +
    "2. 設定シートに「パスワード」を入力\n" +
    "3. メニュー「携帯領収書管理 ツール」→「ダウンロード対象の電話番号を管理」で対象を選択",
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
    "回線管理スプレッドシートのURL（月別シートに電話番号・解約済・運用端末がある別スプレッドシート）。");
  _upsertSettingRow_(sheet, "PDF保存先フォルダ", "https://drive.google.com/drive/folders/XXXXXXXX",
    "Google DriveのフォルダURL、またはローカル絶対パス。");
  _upsertSettingRow_(sheet, "パスワード", "",
    "SoftBank / Y!mobile 共通のログインパスワード。");
  _upsertSettingRow_(sheet, "au/UQパスワード", "",
    "au / UQ mobile 共通のログインパスワード。");
  _upsertSettingRow_(sheet, "au暗証番号", "",
    "au / UQ mobile の4桁暗証番号（請求書閲覧時に必要な場合あり）。", true);
  _upsertSettingRow_(sheet, "dアカウントパスワード", "",
    "docomo のdアカウントパスワード。");
  _upsertSettingRow_(sheet, "対象月", "自動（前月）",
    "ダウンロードする月。「自動（前月）」= 実行時の前月。");
  _setTargetMonthValidation_(sheet);
}


/**
 * 認証情報シート: 電話番号 | キャリア | PDFの種類 | 運用端末 | 状態 | ログインID
 * - サイドバーから選択した電話番号が書き込まれる
 * - パスワードは設定シートで一元管理（認証情報シートにはパスワード列を持たない）
 * - 運用端末はサイドバー保存時に回線管理スプレッドシートから自動設定
 * - 状態列: 有効回線は「契約中」、解約済回線は「解約済」と表示
 * - ログインID列: 回線管理スプレッドシートの「ID」列から自動設定
 *   SoftBank/Ymobile → SoftBank ID、au/UQ → au ID、docomo → dアカウントID
 */
function setupAuthSheet_(ss) {
  let sheet = ss.getSheetByName(AUTH_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(AUTH_SHEET_NAME);
    ss.setActiveSheet(sheet);
    ss.moveActiveSheet(2);
  }

  // ヘッダーが正しくなければ上書き
  const correctHeaders = ["電話番号", "キャリア", "PDFの種類", "運用端末", "状態", "ログインID"];
  const h0 = sheet.getRange(1, 1, 1, Math.max(sheet.getLastColumn(), 1)).getValues()[0]
    .map(v => String(v || "").trim());
  if (h0.length < 6 || h0[0] !== "電話番号" || h0[5] !== "ログインID") {
    sheet.getRange(1, 1, 1, 6).setValues([correctHeaders])
      .setFontWeight("bold").setBackground("#4285F4").setFontColor("#FFFFFF").setHorizontalAlignment("center");
  }

  sheet.setColumnWidth(1, 180);
  sheet.setColumnWidth(2, 120);
  sheet.setColumnWidth(3, 220);
  sheet.setColumnWidth(4, 160);
  sheet.setColumnWidth(5, 100);
  sheet.setColumnWidth(6, 200);
  sheet.getRange("A:A").setNumberFormat("@");

  // キャリア列にドロップダウンを設定
  const lastRow = Math.max(sheet.getLastRow(), 50);
  const carrierRange = sheet.getRange(2, 2, lastRow - 1, 1);
  carrierRange.setDataValidation(SpreadsheetApp.newDataValidation()
    .requireValueInList(["SoftBank", "Ymobile", "au", "UQmobile", "docomo"], true)
    .setAllowInvalid(false).build());

  // PDFの種類列にドロップダウンを設定（SB/YM用 + au/UQ用 + docomo用を結合）
  const allPdfTypes = [...PDF_TYPE_OPTIONS, ...AU_PDF_TYPE_OPTIONS, ...DOCOMO_PDF_TYPE_OPTIONS];
  const pdfTypeRange = sheet.getRange(2, 3, lastRow - 1, 1);
  pdfTypeRange.setDataValidation(SpreadsheetApp.newDataValidation()
    .requireValueInList(allPdfTypes, true)
    .setAllowInvalid(true).build());
}


/**
 * リンクシート: 電話番号 | 名義 | 月列...
 * - 認証情報シートと連動（サイドバー保存時に電話番号を同期）
 * - 「携帯領収書管理 ツール」→「PDFリンク」でDriveのPDFへのハイパーリンクが書き込まれる
 */
function setupLinkSheet_(ss, sheetName, carrierLabel) {
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) sheet = ss.insertSheet(sheetName);
  // ヘッダー色: キャリアブランドカラー
  const colorMap = {
    SoftBank: "#1a1a1a", Ymobile: "#1a1a1a",  // SB系: 黒
    au: "#E94E1B", UQmobile: "#E94E1B",        // KDDI系: オレンジ
    docomo: "#CC0033",                          // docomo: 赤
  };
  const color = colorMap[carrierLabel] || "#4285F4";
  sheet.getRange(1, 1, 1, 2).setValues([["電話番号", "名義"]])
    .setFontWeight("bold").setBackground(color).setFontColor("#FFFFFF").setHorizontalAlignment("center");
  sheet.setColumnWidth(1, 180);
  sheet.setColumnWidth(2, 220);
  sheet.getRange("A:A").setNumberFormat("@");
}


/**
 * ダウンロード履歴シート: 日時 | キャリア | 電話番号 | 対象月 | ファイル名 | 結果
 * - Pythonスクリプトがダウンロード完了時に書き込む
 */
function setupHistorySheet_(ss) {
  let sheet = ss.getSheetByName(HISTORY_SHEET_NAME);
  if (!sheet) sheet = ss.insertSheet(HISTORY_SHEET_NAME);
  if (!sheet.getRange(1, 1).getValue()) {
    const h = ["日時", "キャリア", "電話番号", "対象月", "ファイル名", "結果"];
    sheet.getRange(1, 1, 1, h.length).setValues([h])
      .setFontWeight("bold").setBackground("#34A853").setFontColor("#FFFFFF").setHorizontalAlignment("center");
  }
  sheet.setColumnWidth(1, 160);
  sheet.setColumnWidth(2, 100);
  sheet.setColumnWidth(3, 140);
  sheet.setColumnWidth(4, 80);
  sheet.setColumnWidth(5, 350);
  sheet.setColumnWidth(6, 80);
  sheet.getRange("C:C").setNumberFormat("@");
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
    flex: 1; padding: 8px 4px; border: none; border-radius: 6px 6px 0 0;
    cursor: pointer; font-size: 13px; font-weight: bold;
    background: #e0e0e0; color: #666; transition: all 0.2s;
    line-height: 1.3;
  }
  .tab-btn.active { color: #fff; }
  .tab-btn[data-group="softbank"].active { background: #1a1a1a; }
  .tab-btn[data-group="kddi"].active { background: #E94E1B; }
  .tab-btn[data-group="docomo"].active { background: #CC0033; }
  .carrier-badge { font-size: 14px; margin-right: 2px; }
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
  .btn-reload { background: #e0e0e0; color: #333; font-size: 11px; }
  .empty-msg { padding: 20px; text-align: center; color: #999; }
  .summary { font-size: 11px; color: #666; margin-top: 4px; }
</style>
</head>
<body>
  <div class="tab-bar">
    <button class="tab-btn active" data-group="softbank" onclick="switchGroup('softbank')">🐶🐱<br>SoftBank系</button>
    <button class="tab-btn" data-group="kddi" onclick="switchGroup('kddi')">🍊👸<br>KDDI系</button>
    <button class="tab-btn" data-group="docomo" onclick="switchGroup('docomo')">🍄<br>docomo</button>
  </div>
  <div id="targetMonth" class="summary" style="font-weight:bold; margin-bottom:4px;"></div>
  <div id="status" class="status"></div>
  <div id="phoneList" class="phone-list"><div class="empty-msg">読み込み中...</div></div>
  <div id="summary" class="summary"></div>
  <div class="btn-bar">
    <button class="btn btn-reload" onclick="reload()">再読込</button>
    <button class="btn btn-all" onclick="toggleAll()">全選択/解除</button>
    <button class="btn btn-save" id="saveBtn" onclick="save()">保存</button>
  </div>

<script>
  const PDF_TYPES = ${JSON.stringify(PDF_TYPE_OPTIONS)};
  const AU_PDF_TYPES = ${JSON.stringify(AU_PDF_TYPE_OPTIONS)};
  const DOCOMO_PDF_TYPES = ${JSON.stringify(DOCOMO_PDF_TYPE_OPTIONS)};
  const CARRIER_ICONS = { SoftBank: "🐶", Ymobile: "🐱", au: "🍊", UQmobile: "👸", docomo: "🍄" };
  const GROUPS = {
    softbank: { carriers: ["SoftBank", "Ymobile"], label: "SoftBank系" },
    kddi:     { carriers: ["au", "UQmobile"],      label: "KDDI系" },
    docomo:   { carriers: ["docomo"],               label: "docomo" },
  };
  function getPdfTypesForCarrier(c) {
    if (c === "au" || c === "UQmobile") return AU_PDF_TYPES;
    if (c === "docomo") return DOCOMO_PDF_TYPES;
    return PDF_TYPES;
  }
  function getDefaultPdfType(c) {
    if (c === "docomo") return "一括請求";
    if (c === "au" || c === "UQmobile") return "請求書,支払証明書";
    return "電話番号別";
  }
  let currentGroup = "softbank";
  let phoneData = {};
  let selections = {};

  function setStatus(msg, type) {
    var el = document.getElementById("status");
    el.textContent = msg; el.className = "status " + type;
  }
  function clearStatus() { document.getElementById("status").className = "status"; }

  function init() {
    setStatus("回線管理スプレッドシートから電話番号を読み込み中...", "loading");
    google.script.run
      .withSuccessHandler(function(r) {
        phoneData = r.phones;
        var savedSelections = r.selections;
        document.getElementById("targetMonth").textContent = "対象月: " + (r.targetMonth || "");
        selections = { SoftBank: {}, Ymobile: {}, au: {}, UQmobile: {}, docomo: {} };
        ["SoftBank", "Ymobile", "au", "UQmobile", "docomo"].forEach(function(c) {
          var phones = phoneData[c] || [];
          var saved = savedSelections[c] || {};
          phones.forEach(function(p) {
            selections[c][p.phone] = { pdfType: (saved[p.phone] && saved[p.phone].pdfType) || getDefaultPdfType(c) };
          });
        });
        clearStatus();
        render();
      })
      .withFailureHandler(function(err) { setStatus("読み込みエラー: " + err.message, "error"); })
      .getPhoneManagerData();
  }

  function switchGroup(group) {
    currentGroup = group;
    document.querySelectorAll(".tab-btn").forEach(function(b) {
      b.classList.toggle("active", b.dataset.group === group);
    });
    render();
  }

  function render() {
    var list = document.getElementById("phoneList");
    var group = GROUPS[currentGroup];
    var allPhones = [];
    group.carriers.forEach(function(c) {
      (phoneData[c] || []).forEach(function(p) {
        allPhones.push({ phone: p.phone, carrier: c, device: p.device, cancelled: p.cancelled });
      });
    });

    if (allPhones.length === 0) {
      list.innerHTML = '<div class="empty-msg">' + group.label + ' の電話番号が見つかりません。<br>設定シートの「回線管理スプレッドシート」URLを確認してください。</div>';
      updateSummary(); return;
    }

    var html = "";
    allPhones.forEach(function(p) {
      var sel = selections[p.carrier] || {};
      var checked = sel[p.phone] ? "checked" : "";
      var pdfType = (sel[p.phone] && sel[p.phone].pdfType) || getDefaultPdfType(p.carrier);
      var icon = CARRIER_ICONS[p.carrier] || "";
      var label = p.device ? p.phone + " (" + p.device + ")" : p.phone;
      if (p.cancelled) label += "（解約済）";
      var itemClass = p.cancelled ? "phone-item cancelled" : "phone-item";
      var pdfTypes = getPdfTypesForCarrier(p.carrier);

      html += '<div class="' + itemClass + '">'
        + '<input type="checkbox" data-phone="' + p.phone + '" data-carrier="' + p.carrier + '" ' + checked
        + ' onchange="onCheck(this)">'
        + '<span class="carrier-badge">' + icon + '</span>'
        + '<span class="phone-number">' + label + '</span>';
      if (pdfTypes.length > 1) {
        html += '<select data-phone="' + p.phone + '" data-carrier="' + p.carrier + '" onchange="onPdfType(this)">';
        pdfTypes.forEach(function(t) {
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
    var carrier = el.dataset.carrier;
    if (!selections[carrier]) selections[carrier] = {};
    if (el.checked) {
      var selEl = el.parentElement.querySelector("select");
      selections[carrier][el.dataset.phone] = { pdfType: selEl ? selEl.value : getDefaultPdfType(carrier) };
    } else {
      delete selections[carrier][el.dataset.phone];
    }
    updateSummary();
  }

  function onPdfType(el) {
    var carrier = el.dataset.carrier;
    if (selections[carrier] && selections[carrier][el.dataset.phone]) {
      selections[carrier][el.dataset.phone].pdfType = el.value;
    }
  }

  function toggleAll() {
    var group = GROUPS[currentGroup];
    var allPhones = [];
    group.carriers.forEach(function(c) {
      (phoneData[c] || []).forEach(function(p) { allPhones.push({ phone: p.phone, carrier: c }); });
    });
    var allChecked = allPhones.every(function(p) { return (selections[p.carrier] || {})[p.phone]; });
    if (allChecked) {
      allPhones.forEach(function(p) { if (selections[p.carrier]) delete selections[p.carrier][p.phone]; });
    } else {
      allPhones.forEach(function(p) {
        if (!selections[p.carrier]) selections[p.carrier] = {};
        if (!selections[p.carrier][p.phone])
          selections[p.carrier][p.phone] = { pdfType: getDefaultPdfType(p.carrier) };
      });
    }
    render();
  }

  function updateSummary() {
    var counts = [];
    ["SoftBank", "Ymobile", "au", "UQmobile", "docomo"].forEach(function(c) {
      var n = Object.keys(selections[c] || {}).length;
      if (n > 0) counts.push(CARRIER_ICONS[c] + " " + n);
    });
    document.getElementById("summary").textContent = "選択中: " + (counts.length > 0 ? counts.join(" / ") : "なし");
  }

  function reload() {
    phoneData = {};
    selections = {};
    document.getElementById("phoneList").innerHTML = '<div class="empty-msg">読み込み中...</div>';
    init();
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
  _clearSettingsCache_();
  try {
    const target = _getTargetMonth_();
    return {
      phones: _getAllPhonesFromMonthSheets_(),
      selections: _getCurrentSelections_(),
      targetMonth: target.year + "年" + target.month + "月",
    };
  } catch (e) {
    Logger.log(`[getPhoneManagerData] エラー: ${e.message}\n${e.stack}`);
    throw new Error(`データ読み込みに失敗: ${e.message}`);
  }
}


/**
 * サイドバーからの保存。
 * 認証情報シートとリンクシートの両方を更新する。
 * 解約済でも選択されていればPDFの種類付きで認証情報シートに書き込む。
 */
function savePhoneSelections(selections) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const allPhones = _getAllPhonesFromMonthSheets_();

    // 解約済番号セット・運用端末・名義マップ
    const cancelledSet = new Set();
    const deviceMap = {}, nameMap = {}, loginIdMap = {};
    const ALL_CARRIERS = ["SoftBank", "Ymobile", "au", "UQmobile", "docomo"];
    for (const carrier of ALL_CARRIERS) {
      for (const p of (allPhones[carrier] || [])) {
        if (p.cancelled) cancelledSet.add(p.phone);
        if (p.device) deviceMap[p.phone] = p.device;
        if (p.name) nameMap[p.phone] = p.name;
        if (p.loginId) loginIdMap[p.phone] = p.loginId;
      }
    }

    // ── 認証情報シート ──
    let authSheet = ss.getSheetByName(AUTH_SHEET_NAME);
    if (!authSheet) {
      authSheet = ss.insertSheet(AUTH_SHEET_NAME);
      setupAuthSheet_(ss);
    }

    // 既存データを全クリア（ヘッダー行以外）
    // getMaxRows()で書式だけの行も含めて確実にクリアする
    const lastRow = Math.max(authSheet.getLastRow(), authSheet.getMaxRows());
    const lastCol = Math.max(authSheet.getLastColumn(), 6);
    Logger.log(`[save] クリア: lastRow=${lastRow}, lastCol=${lastCol}`);
    if (lastRow > 1) {
      const rows = lastRow - 1;
      authSheet.getRange(2, 1, rows, lastCol).clearContent();
      authSheet.getRange(2, 1, rows, lastCol).clearFormat();
      authSheet.getRange(2, 1, rows, lastCol)
        .setFontLine("none").setFontColor(null).setBackground(null).setFontWeight("normal");
      SpreadsheetApp.flush();
    }

    // 選択中の番号（契約中・解約済を分離）
    const activeRows = [];
    const cancelledSelectedRows = [];
    const linkData = { SoftBank: [], Ymobile: [], au: [], UQmobile: [], docomo: [] };

    for (const carrier of ALL_CARRIERS) {
      const sel = selections[carrier] || {};
      for (const phone of Object.keys(sel)) {
        const defaultPdfType = (carrier === "docomo") ? "一括請求" : (carrier === "au" || carrier === "UQmobile") ? "請求書,支払証明書" : "電話番号別";
        const pdfType = sel[phone].pdfType || defaultPdfType;
        const status = cancelledSet.has(phone) ? "解約済" : "契約中";
        if (status === "解約済") {
          cancelledSelectedRows.push([phone, carrier, pdfType, deviceMap[phone] || "", "解約済", loginIdMap[phone] || ""]);
        } else {
          activeRows.push([phone, carrier, pdfType, deviceMap[phone] || "", "契約中", loginIdMap[phone] || ""]);
        }
        linkData[carrier].push({ phone, name: nameMap[phone] || "" });
      }
    }

    const cancelledRows = [...cancelledSelectedRows];

    const allRows = [...activeRows, ...cancelledRows];
    if (allRows.length > 0) {
      authSheet.getRange(2, 1, allRows.length, 6).setValues(allRows);
    }
    // 新データより下のゴミ行・ドロップダウンを確実にクリア
    const newLastRow = allRows.length + 1;
    const maxRows = authSheet.getMaxRows();
    if (maxRows > newLastRow) {
      const surplusRows = maxRows - newLastRow;
      authSheet.getRange(newLastRow + 1, 1, surplusRows, lastCol).clearContent();
      authSheet.getRange(newLastRow + 1, 1, surplusRows, lastCol).clearFormat();
      authSheet.getRange(newLastRow + 1, 1, surplusRows, lastCol).clearDataValidations();
    }
    authSheet.getRange("A:A").setNumberFormat("@");

    // 解約済行をグレーアウト+取り消し線
    if (cancelledRows.length > 0) {
      const startRow = activeRows.length + 2;
      authSheet.getRange(startRow, 1, cancelledRows.length, 6)
        .setFontLine("line-through").setFontColor("#999999").setBackground("#f0f0f0");
      // 解約済行のドロップダウンバリデーションをクリア（空値を許容するため）
      authSheet.getRange(startRow, 2, cancelledRows.length, 1).clearDataValidations();
      authSheet.getRange(startRow, 3, cancelledRows.length, 1).clearDataValidations();
    }

    // ── リンクシートの電話番号を同期（解約済もグレー表示） ──
    _syncLinkSheetPhones_(ss, SOFTBANK_LINK_SHEET_NAME, linkData.SoftBank, cancelledSet);
    _syncLinkSheetPhones_(ss, YMOBILE_LINK_SHEET_NAME, linkData.Ymobile, cancelledSet);
    _syncLinkSheetPhones_(ss, AU_LINK_SHEET_NAME, linkData.au, cancelledSet);
    _syncLinkSheetPhones_(ss, UQMOBILE_LINK_SHEET_NAME, linkData.UQmobile, cancelledSet);
    _syncLinkSheetPhones_(ss, DOCOMO_LINK_SHEET_NAME, linkData.docomo, cancelledSet);

    Logger.log(`[save] 書き込み: active=${activeRows.length}, cancelled=${cancelledSelectedRows.length}, total=${allRows.length}`);
    let msg = `保存しました（契約中${activeRows.length}件`;
    if (cancelledSelectedRows.length > 0) msg += `、解約済${cancelledSelectedRows.length}件`;
    msg += `、合計${allRows.length}行）。`;
    return msg;
  } catch (e) {
    Logger.log(`[savePhoneSelections] エラー: ${e.message}\n${e.stack}`);
    throw new Error(`保存に失敗しました: ${e.message}`);
  }
}


/**
 * リンクシートに電話番号を追加・名義を更新する。
 * 既存行は削除しない（PDFリンク等の月列データを保持するため）。
 * 解約済の行はグレーアウト+取り消し線で視覚的に区別する。
 */
function _syncLinkSheetPhones_(ss, sheetName, phoneList, cancelledSet) {
  try {
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) return;

    // 現在のリンクシートの電話番号を取得
    const data = sheet.getDataRange().getValues();
    const existingPhones = {};
    for (let i = 1; i < data.length; i++) {
      const phone = _normalizePhone_(data[i][0]);
      if (phone) existingPhones[phone] = i + 1;
    }

    // 追加・名義更新（既存行は削除しない）
    for (const { phone, name } of phoneList) {
      if (existingPhones[phone]) {
        sheet.getRange(existingPhones[phone], 2).setValue(name);
      } else {
        const newRow = sheet.getLastRow() + 1;
        sheet.getRange(newRow, 1).setNumberFormat("@").setValue(phone);
        sheet.getRange(newRow, 2).setValue(name);
      }
    }

    // 電話番号・名義列（A,B列）の解約済スタイルを追加のみ（解除はしない）
    // 別の月で解約済だった回線のグレーアウトが消えないようにする
    const refreshedData = sheet.getDataRange().getValues();
    for (let i = 1; i < refreshedData.length; i++) {
      const phone = _normalizePhone_(refreshedData[i][0]);
      if (!phone) continue;
      const row = i + 1;
      if (cancelledSet && cancelledSet.has(phone)) {
        const abRange = sheet.getRange(row, 1, 1, 2);
        abRange.setFontLine("line-through").setFontColor("#999999");
      }
      // cancelledSetに含まれない回線の既存スタイルはそのまま維持
    }
  } catch (e) {
    Logger.log(`[_syncLinkSheetPhones_] ${sheetName} エラー: ${e.message}`);
  }
}


// ── 回線管理スプレッドシートの読み込み ──

function _normalizeCarrierName_(text) {
  const s = String(text).trim().toLowerCase();
  if (!s) return null;
  if (s === "softbank" || s.includes("ソフトバンク")) return "SoftBank";
  if (s === "ymobile" || s === "y!mobile" || s.includes("ワイモバイル")) return "Ymobile";
  if (s === "au" || s.includes("エーユー") || s === "kddi") return "au";
  if (s === "uqmobile" || s === "uq mobile" || s === "uq" || s.includes("ユーキュー")) return "UQmobile";
  if (s === "docomo" || s === "nttdocomo" || s.includes("ドコモ")) return "docomo";
  return null;
}

/**
 * 設定シートの対象月から対象の月番号を取得する。
 * 「自動（前月）」または未設定の場合は前月を返す。
 * 返却値: { year: number, month: number } (例: { year: 2026, month: 1 })
 */
function _getTargetMonth_() {
  const raw = _getSettingValue_("対象月");
  // Date型（スプレッドシートが「2026年1月」をDateに変換する場合）
  if (raw instanceof Date) {
    return { year: raw.getFullYear(), month: raw.getMonth() + 1 };
  }
  const val = String(raw || "").trim();
  // YYYYMM形式
  const m1 = val.match(/^(\d{4})(\d{2})$/);
  if (m1) return { year: parseInt(m1[1]), month: parseInt(m1[2]) };
  // YYYY年MM月形式
  const m2 = val.match(/^(\d{4})年(\d{1,2})月$/);
  if (m2) return { year: parseInt(m2[1]), month: parseInt(m2[2]) };
  // 「自動（前月）」またはその他 → 前月
  const now = new Date();
  let y = now.getFullYear(), m = now.getMonth(); // getMonth() is 0-based
  if (m === 0) { y--; m = 12; }
  return { year: y, month: m };
}

/**
 * 回線管理スプレッドシートの月別シートから電話番号を収集する。
 * 設定シートの対象月に対応する月シートを読み込む。
 * セクション分け構造に対応（キャリアラベル行 → ヘッダー行 → データ行）。
 */
function _getAllPhonesFromMonthSheets_() {
  const result = { SoftBank: [], Ymobile: [], au: [], UQmobile: [], docomo: [] };

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

  // 対象月に対応するシートを特定
  const target = _getTargetMonth_();
  const targetNum = target.year * 100 + target.month;
  monthSheets.sort((a, b) => _parseMonthSheetNum_(a.getName()) - _parseMonthSheetNum_(b.getName()));

  let targetSheet = null;
  for (const sheet of monthSheets) {
    if (_parseMonthSheetNum_(sheet.getName()) === targetNum) {
      targetSheet = sheet;
      break;
    }
  }
  // 対象月シートが見つからない場合はデータあり最新月にフォールバック
  if (!targetSheet) {
    for (let i = monthSheets.length - 1; i >= 0; i--) {
      if (monthSheets[i].getLastRow() > 1) { targetSheet = monthSheets[i]; break; }
    }
  }
  if (!targetSheet || targetSheet.getLastRow() <= 1) return result;

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

    const phone = _normalizePhone_(row[cols["電話番号"]]);
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
    const ii = cols["ID"];
    const loginId = ii !== undefined ? String(row[ii] || "").trim() : "";

    if (!result[carrier].some(p => p.phone === phone)) {
      result[carrier].push({ phone, cancelled, device, name, loginId });
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
  const result = { SoftBank: {}, Ymobile: {}, au: {}, UQmobile: {}, docomo: {} };
  if (!sheet || sheet.getLastRow() <= 1) return result;

  const data = sheet.getDataRange().getValues();
  const h = data[0].map(v => String(v).trim());
  const pi = h.indexOf("電話番号"), ci = h.indexOf("キャリア"), ti = h.indexOf("PDFの種類");
  if (pi === -1) return result;

  for (let i = 1; i < data.length; i++) {
    const phone = _normalizePhone_(data[i][pi]);
    const carrier = ci !== -1 ? String(data[i][ci] || "").trim() : "";
    const defaultPdf = (carrier === "docomo") ? "一括請求" : (carrier === "au" || carrier === "UQmobile") ? "請求書,支払証明書" : "電話番号別";
    const pdfType = ti !== -1 ? String(data[i][ti] || "").trim() : defaultPdf;
    if (phone && carrier && result[carrier]) {
      result[carrier][phone] = { pdfType: pdfType || defaultPdf };
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

function updateSoftBankLinks()      { _updatePdfLinks_(SOFTBANK_LINK_SHEET_NAME, "SoftBank"); }
function updateYmobileLinks()       { _updatePdfLinks_(YMOBILE_LINK_SHEET_NAME, "Ymobile"); }
function updateAuLinks()            { _updatePdfLinks_(AU_LINK_SHEET_NAME, "au"); }
function updateUQmobileLinks()      { _updatePdfLinks_(UQMOBILE_LINK_SHEET_NAME, "UQmobile"); }
function updateDocomoLinks()        { _updatePdfLinks_(DOCOMO_LINK_SHEET_NAME, "docomo"); }
function forceUpdateSoftBankLinks() { _updatePdfLinks_(SOFTBANK_LINK_SHEET_NAME, "SoftBank", false, true); }
function forceUpdateYmobileLinks()  { _updatePdfLinks_(YMOBILE_LINK_SHEET_NAME, "Ymobile", false, true); }
function forceUpdateAuLinks()       { _updatePdfLinks_(AU_LINK_SHEET_NAME, "au", false, true); }
function forceUpdateUQmobileLinks() { _updatePdfLinks_(UQMOBILE_LINK_SHEET_NAME, "UQmobile", false, true); }
function forceUpdateDocomoLinks()   { _updatePdfLinks_(DOCOMO_LINK_SHEET_NAME, "docomo", false, true); }

function updateAllLinks()      { _runAllLinkUpdates_(false); }
function forceUpdateAllLinks() { _runAllLinkUpdates_(true); }

function _runAllLinkUpdates_(force) {
  const sb = _updatePdfLinks_(SOFTBANK_LINK_SHEET_NAME, "SoftBank", true, force);
  const ym = _updatePdfLinks_(YMOBILE_LINK_SHEET_NAME, "Ymobile", true, force);
  const au = _updatePdfLinks_(AU_LINK_SHEET_NAME, "au", true, force);
  const uq = _updatePdfLinks_(UQMOBILE_LINK_SHEET_NAME, "UQmobile", true, force);
  const dc = _updatePdfLinks_(DOCOMO_LINK_SHEET_NAME, "docomo", true, force);
  const totals = [sb, ym, au, uq, dc];
  _showLinkUpdateResult_(
    totals.reduce((s, r) => s + (r.pdfs || 0), 0),
    totals.reduce((s, r) => s + (r.added || 0), 0),
    totals.reduce((s, r) => s + (r.overwritten || 0), 0),
    totals.reduce((s, r) => s + (r.newPhones || 0), 0),
    force
  );
}

/**
 * Google Driveを探索してリンクシートにPDFリンクを設定する。
 * リンクシートに電話番号がない場合でも、Driveで見つかったPDFの電話番号を自動追加する。
 * silent=true のときはポップアップを出さず結果オブジェクトを返す（一括更新用）。
 * force=true のときは設定済みセルも上書きする（通常は設定済みをスキップ）。
 */
function _updatePdfLinks_(sheetName, carrier, silent = false, force = false) {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    ui.alert("エラー", `「${sheetName}」が見つかりません。setupSheet()を実行してください。`, ui.ButtonSet.OK);
    return { pdfs: 0, updated: 0, newPhones: 0 };
  }

  const rootFolderUrl = _getSettingValue_("PDF保存先フォルダ");
  if (!rootFolderUrl || !rootFolderUrl.startsWith("https://drive.google.com/")) {
    ui.alert("エラー", "設定シートの「PDF保存先フォルダ」にDriveのURLを設定してください。", ui.ButtonSet.OK);
    return { pdfs: 0, updated: 0, newPhones: 0 };
  }

  const folderMatch = rootFolderUrl.match(/\/folders\/([a-zA-Z0-9_-]+)/);
  if (!folderMatch) { ui.alert("エラー", "フォルダURLからIDを取得できません。", ui.ButtonSet.OK); return { pdfs: 0, updated: 0, newPhones: 0 }; }
  let rootFolder;
  try { rootFolder = DriveApp.getFolderById(folderMatch[1]); }
  catch (e) { ui.alert("エラー", `フォルダアクセス失敗: ${e.message}`, ui.ButtonSet.OK); return { pdfs: 0, updated: 0, newPhones: 0 }; }

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

  // リンク対象は請求書/電話番号別のみ（領収書・支払証明書・一括・機種別は除外）
  const excludeSuffixes = ["_領収書.pdf", "_支払証明書.pdf", "_一括.pdf", "_機種別.pdf"];
  const linkEntries = pdfEntries.filter(e => {
    const name = e.file.getName();
    return !excludeSuffixes.some(s => name.endsWith(s));
  });

  if (linkEntries.length === 0) {
    if (!silent) SpreadsheetApp.getUi().alert("情報", `${carrier}のPDFが見つかりません。`, SpreadsheetApp.getUi().ButtonSet.OK);
    return { pdfs: 0, updated: 0, newPhones: 0 };
  }

  // 電話番号→行番号マップ（なければ自動追加）
  let data = sheet.getDataRange().getValues();
  const phoneToRow = {};
  for (let i = 1; i < data.length; i++) {
    const ph = _normalizePhone_(data[i][0]);
    if (ph) phoneToRow[ph] = i + 1;
  }

  // PDFに含まれるがリンクシートにない電話番号を自動追加（名義は認証情報から取得不可のため空欄）
  const newPhones = [...new Set(linkEntries.map(e => e.phone))].filter(ph => !phoneToRow[ph]);
  for (const ph of newPhones) {
    const row = sheet.getLastRow() + 1;
    sheet.getRange(row, 1).setNumberFormat("@").setValue(ph);
    sheet.getRange(row, 2).setValue("");
    phoneToRow[ph] = row;
  }

  // 月列を確保
  const neededHeaders = [...new Set(linkEntries.map(e => `${e.year}年${parseInt(e.month)}月`))];
  neededHeaders.sort((a, b) => _monthHeaderToNum_(b) - _monthHeaderToNum_(a));
  for (const h of neededHeaders) _ensureMonthColumn_(sheet, h);

  // リンク書き込み（1電話番号×1月につき1リンク）
  let added = 0, overwritten = 0;
  for (const entry of linkEntries) {
    const rowNum = phoneToRow[entry.phone];
    const colNum = _getMonthColumnNum_(sheet, entry.year, entry.month);
    if (!rowNum || !colNum) continue;

    const cell = sheet.getRange(rowNum, colNum);
    const hasFormula = !!cell.getFormula();
    if (hasFormula && !force) { overwritten++; continue; }
    const url = entry.file.getUrl();
    const am = entry.file.getName().match(/_(\d+)円(?:\(税抜\))?\.pdf$/);
    const label = am ? `${Number(am[1]).toLocaleString()}円` : `${parseInt(entry.month)}月`;
    cell.setFormula(`=HYPERLINK("${url}","${label}")`);
    if (hasFormula) { overwritten++; } else { added++; }
  }

  if (!silent) _showLinkUpdateResult_(linkEntries.length, added, overwritten, newPhones.length, force);
  return { pdfs: linkEntries.length, added, overwritten, newPhones: newPhones.length };
}


// ────────────────────────────────────────────────
//  PDFから金額を取得・ファイル名更新
// ────────────────────────────────────────────────

function scanAndUpdateSoftBankAmounts() { _scanAndUpdatePdfAmounts_(SOFTBANK_LINK_SHEET_NAME, "SoftBank"); }
function scanAndUpdateYmobileAmounts() { _scanAndUpdatePdfAmounts_(YMOBILE_LINK_SHEET_NAME, "Ymobile"); }
function scanAndUpdateAuAmounts() { _scanAndUpdatePdfAmounts_(AU_LINK_SHEET_NAME, "au"); }
function scanAndUpdateUQmobileAmounts() { _scanAndUpdatePdfAmounts_(UQMOBILE_LINK_SHEET_NAME, "UQmobile"); }
function scanAndUpdateDocomoAmounts() { _scanAndUpdatePdfAmounts_(DOCOMO_LINK_SHEET_NAME, "docomo"); }
function scanAndUpdateAllAmounts() {
  const sb = _scanAndUpdatePdfAmounts_(SOFTBANK_LINK_SHEET_NAME, "SoftBank", true);
  const ym = _scanAndUpdatePdfAmounts_(YMOBILE_LINK_SHEET_NAME, "Ymobile", true);
  const au = _scanAndUpdatePdfAmounts_(AU_LINK_SHEET_NAME, "au", true);
  const uq = _scanAndUpdatePdfAmounts_(UQMOBILE_LINK_SHEET_NAME, "UQmobile", true);
  const dc = _scanAndUpdatePdfAmounts_(DOCOMO_LINK_SHEET_NAME, "docomo", true);
  const total = (sb.total || 0) + (ym.total || 0) + (au.total || 0) + (uq.total || 0) + (dc.total || 0);
  const updated = (sb.updated || 0) + (ym.updated || 0) + (au.updated || 0) + (uq.updated || 0) + (dc.updated || 0);
  const failed = (sb.failed || 0) + (ym.failed || 0) + (au.failed || 0) + (uq.failed || 0) + (dc.failed || 0);
  if (total === 0) {
    SpreadsheetApp.getUi().alert("情報", "金額が未取得のPDFはありませんでした。", SpreadsheetApp.getUi().ButtonSet.OK);
  } else {
    SpreadsheetApp.getUi().alert("完了", `${total}件処理: 更新${updated}件、失敗${failed}件`, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

function _scanAndUpdatePdfAmounts_(sheetName, carrier, silent = false) {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) { ui.alert("エラー", `「${sheetName}」が見つかりません。`, ui.ButtonSet.OK); return { total: 0, updated: 0, failed: 0 }; }

  const rootFolderUrl = _getSettingValue_("PDF保存先フォルダ");
  if (!rootFolderUrl || !rootFolderUrl.startsWith("https://drive.google.com/")) {
    ui.alert("エラー", "PDF保存先フォルダが未設定です。", ui.ButtonSet.OK); return { total: 0, updated: 0, failed: 0 };
  }
  const folderMatch = rootFolderUrl.match(/\/folders\/([a-zA-Z0-9_-]+)/);
  if (!folderMatch) { ui.alert("エラー", "フォルダIDを取得できません。", ui.ButtonSet.OK); return { total: 0, updated: 0, failed: 0 }; }
  let rootFolder;
  try { rootFolder = DriveApp.getFolderById(folderMatch[1]); }
  catch (e) { ui.alert("エラー", `フォルダアクセス失敗: ${e.message}`, ui.ButtonSet.OK); return { total: 0, updated: 0, failed: 0 }; }

  const pdfEntries = [];
  const yfs = rootFolder.getFolders();
  while (yfs.hasNext()) { const yf = yfs.next(); if (!/^\d{4}$/.test(yf.getName())) continue;
    const mfs = yf.getFolders(); while (mfs.hasNext()) { const mf = mfs.next(); if (!/^\d{2}$/.test(mf.getName())) continue;
      const sfs = mf.getFolders(); while (sfs.hasNext()) { const sf = sfs.next(); if (sf.getName() === carrier) _collectCarrierPdfs_(sf, carrier, pdfEntries); }
      _collectCarrierPdfs_(mf, carrier, pdfEntries);
  }}

  const targets = pdfEntries.filter(e => e.file.getName().endsWith("_利用料金明細.pdf"));
  if (targets.length === 0) {
    if (!silent) SpreadsheetApp.getUi().alert("情報", "金額が未取得のPDFはありませんでした。", SpreadsheetApp.getUi().ButtonSet.OK);
    return { total: 0, updated: 0, failed: 0 };
  }

  const data = sheet.getDataRange().getValues();
  const phoneToRow = {};
  for (let i = 1; i < data.length; i++) {
    const ph = _normalizePhone_(data[i][0]);
    if (ph) phoneToRow[ph] = i + 1;
  }

  const isAuFamily = (carrier === "au" || carrier === "UQmobile");
  let updated = 0, failed = 0;
  for (const entry of targets) {
    // au/UQはまとめ請求PDFなので電話番号を渡して個別金額を取得
    const amount = _extractAmountFromPdf_(entry.file, isAuFamily ? entry.phone : null);
    if (!amount) { failed++; continue; }
    entry.file.setName(entry.file.getName().replace("_利用料金明細.pdf", `_${amount}円.pdf`));
    const rowNum = phoneToRow[entry.phone];
    if (rowNum) {
      const colNum = _getMonthColumnNum_(sheet, entry.year, entry.month);
      if (colNum) sheet.getRange(rowNum, colNum).setFormula(`=HYPERLINK("${entry.file.getUrl()}","${parseInt(entry.month)}月 ${Number(amount).toLocaleString()}円")`);
    }
    updated++;
  }
  if (!silent) SpreadsheetApp.getUi().alert("完了", `${targets.length}件処理: 更新${updated}件、失敗${failed}件`, SpreadsheetApp.getUi().ButtonSet.OK);
  return { total: targets.length, updated, failed };
}


// ────────────────────────────────────────────────
//  内部ヘルパー
// ────────────────────────────────────────────────

let _settingsCache = null;

function _getSettingValue_(key) {
  if (!_settingsCache) {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SETTINGS_SHEET_NAME);
    if (!sheet) return null;
    const data = sheet.getDataRange().getValues();
    _settingsCache = {};
    for (let i = 1; i < data.length; i++) {
      const k = String(data[i][0]).trim();
      const raw = data[i][1];
      // Date型はそのまま保持（対象月のドロップダウンがDateに変換されるため）
      const v = (raw instanceof Date) ? raw : (String(raw).trim() || null);
      if (k) _settingsCache[k] = v;
    }
  }
  return _settingsCache[key] || null;
}

function _clearSettingsCache_() { _settingsCache = null; }

function _normalizePhone_(phone) {
  let s = String(phone || "").trim();
  s = s.replace(/[０-９]/g, c => String.fromCharCode(c.charCodeAt(0) - 0xFEE0));
  s = s.replace(/[-\s\u2010-\u2015\u2212\uFF0D]/g, "");
  if (s.length === 10 && s[0] !== "0") s = "0" + s;
  return s;
}

function _collectCarrierPdfs_(folder, carrier, results) {
  const files = folder.getFiles();
  while (files.hasNext()) {
    const file = files.next();
    if (file.getMimeType() !== "application/pdf") continue;
    // YYYYMM_会社名_電話番号_...
    const m = file.getName().match(/^(\d{4})(\d{2})_[^_]+_(\d{10,})/);
    if (m) results.push({ year: m[1], month: m[2], phone: m[3], file });
  }
}

function _monthHeaderToNum_(h) {
  const m = String(h).match(/^(\d{4})年(\d+)月$/);
  return m ? parseInt(m[1]) * 100 + parseInt(m[2]) : 0;
}

function _ensureMonthColumn_(sheet, monthHeader) {
  SpreadsheetApp.flush();
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

function _upsertSettingRow_(sheet, key, defaultValue, noteText, forceText = false) {
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === key) {
      // 既存行があってもテキスト書式を適用（先頭0の保持のため）
      if (forceText) sheet.getRange(i + 1, 2).setNumberFormat("@");
      return;
    }
  }
  const r = sheet.getLastRow() + 1;
  sheet.getRange(r, 1).setValue(key);
  const v = sheet.getRange(r, 2);
  if (forceText) v.setNumberFormat("@");
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

function _showLinkUpdateResult_(total, added, skippedOrOverwritten, newPhones, force) {
  const parts = [];
  if (added > 0) parts.push(`新規 ${added}件`);
  if (skippedOrOverwritten > 0) parts.push(force ? `上書き ${skippedOrOverwritten}件` : `設定済み ${skippedOrOverwritten}件（スキップ）`);
  const linkMsg = parts.length > 0 ? parts.join("、") : "変更なし";
  SpreadsheetApp.getUi().alert("完了",
    `${total}件のPDFを確認（${linkMsg}）` +
    (newPhones > 0 ? `\n${newPhones}件の電話番号を自動追加しました。` : ""),
    SpreadsheetApp.getUi().ButtonSet.OK);
}

/**
 * PDFから金額を取得する。
 * SoftBank/Ymobile: 「計」の後の金額を取得（1回線1PDF）
 * au/UQmobile: 電話番号に対応する個別金額を取得（まとめ請求PDF）
 *   PDFの1ページ目に「（内訳）０８０－１４３８－８３４３  ( 1,851 )」形式で記載
 */
function _extractAmountFromPdf_(file, phone) {
  let docFileId = null;
  try {
    // Drive OCR: ocr:true + ocrLanguage で高精度テキスト抽出
    let docFile;
    try {
      docFile = Drive.Files.insert(
        { title: "tmp_ocr_" + file.getName() },
        file.getBlob(),
        { convert: true, ocr: true, ocrLanguage: "ja" }
      );
    } catch (e1) {
      // フォールバック: ocrなし
      docFile = Drive.Files.insert(
        { title: "tmp_ocr_" + file.getName() },
        file.getBlob(),
        { convert: true }
      );
    }
    docFileId = docFile.id;
    // NFKC正規化: 全角数字→半角、全角記号→半角（Apple領収書管理と同じ手法）
    let text = DocumentApp.openById(docFileId).getBody().getText();
    text = String(text).normalize("NFKC");
    // ハイフン類を統一
    text = text.replace(/[‐-‒–—−﹣－]/g, "-");

    // au/UQmobile: 電話番号が指定された場合、その番号の個別金額を取得
    if (phone) {
      const digits = phone.replace(/\D/g, "");
      const formatted = digits.slice(0, 3) + "-" + digits.slice(3, 7) + "-" + digits.slice(7);

      Logger.log(`[OCR] au/UQ searching: ${formatted}`);

      // 電話番号の位置を見つける
      const phoneIdx = text.indexOf(formatted);
      if (phoneIdx === -1) {
        Logger.log(`[OCR] phone not found in text`);
        return null;
      }

      // この電話番号のセクション（次の電話番号パターンまで、または末尾から1000文字）
      const afterPhone = text.substring(phoneIdx);
      // 次の電話番号パターン（XXX-XXXX-XXXX）を探す（自分自身を除く）
      const nextPhoneMatch = afterPhone.substring(formatted.length).match(/\d{3}-\d{4}-\d{4}/);
      const sectionEnd = nextPhoneMatch ? formatted.length + nextPhoneMatch.index : Math.min(afterPhone.length, 1000);
      const section = afterPhone.substring(0, sectionEnd);

      Logger.log(`[OCR] section(200): ${section.substring(0, 200)}`);

      // パターン1: 括弧付き金額 ( 1,851 ) — 一番最後のものを採用
      const pat1Matches = [...section.matchAll(/\(\s*([\d,]+)\s*\)/g)];
      if (pat1Matches.length > 0) {
        const lastMatch = pat1Matches[pat1Matches.length - 1];
        const a = lastMatch[1].replace(/,/g, "");
        if (/^\d+$/.test(a) && a.length <= 7 && parseInt(a) > 0) {
          Logger.log(`[OCR] found bracketed amount: ${a}`);
          return a;
        }
      }

      // パターン2: 「小計」「合計」の後の金額
      const pat2 = section.match(/(?:小計|合計)[^\d]*([\d,]+)/);
      if (pat2) {
        const a = pat2[1].replace(/,/g, "");
        if (/^\d+$/.test(a) && a.length <= 7 && parseInt(a) > 0) {
          Logger.log(`[OCR] found subtotal: ${a}`);
          return a;
        }
      }

      // パターン3: セクション内の最後の有意な金額（3桁以上の数字）
      const allAmounts = [...section.matchAll(/([\d,]{3,})/g)]
        .map(m => m[1].replace(/,/g, ""))
        .filter(a => /^\d+$/.test(a) && parseInt(a) >= 100 && parseInt(a) <= 9999999);
      if (allAmounts.length > 0) {
        const lastAmount = allAmounts[allAmounts.length - 1];
        Logger.log(`[OCR] fallback last amount in section: ${lastAmount}`);
        return lastAmount;
      }

      Logger.log(`[OCR] au/UQ phone=${phone} not found. Near text: ${section.substring(0, 100)}`);
      return null;
    }

    // SoftBank/Ymobile: 「計」の後の金額
    for (const pat of [/(?<![小合])計[^\d]*([\d,]+)/, /小計[^\d]*([\d,]+)/]) {
      const m = text.match(pat);
      if (m) { const a = m[1].replace(/,/g, ""); if (/^\d+$/.test(a)) return a; }
    }
    return null;
  } catch (e) { Logger.log(`[OCR] error: ${e.message}`); return null; }
  finally { if (docFileId) try { Drive.Files.remove(docFileId); } catch (_) {} }
}


// ────────────────────────────────────────────────
//  メニュー
// ────────────────────────────────────────────────

function onOpen() {
  const linkMenu = SpreadsheetApp.getUi().createMenu("PDFリンク")
    .addItem("全キャリア一括更新", "updateAllLinks")
    .addSeparator()
    .addItem("SoftBankのみ", "updateSoftBankLinks")
    .addItem("Ymobileのみ", "updateYmobileLinks")
    .addItem("auのみ", "updateAuLinks")
    .addItem("UQmobileのみ", "updateUQmobileLinks")
    .addItem("docomoのみ", "updateDocomoLinks")
    .addSeparator()
    .addItem("全キャリア一括（強制上書き）", "forceUpdateAllLinks")
    .addItem("SoftBankのみ（強制上書き）", "forceUpdateSoftBankLinks")
    .addItem("Ymobileのみ（強制上書き）", "forceUpdateYmobileLinks")
    .addItem("auのみ（強制上書き）", "forceUpdateAuLinks")
    .addItem("UQmobileのみ（強制上書き）", "forceUpdateUQmobileLinks")
    .addItem("docomoのみ（強制上書き）", "forceUpdateDocomoLinks");

  // 金額取得・ファイル名更新はPythonのダウンロード時に自動取得するためGASメニューから削除
  // au/UQのまとめ請求PDFはOCRでの個別金額取得が不正確なため、Pythonのページ取得に一本化

  SpreadsheetApp.getUi().createMenu("携帯領収書管理 ツール")
    .addItem("初期セットアップ", "setupSheet")
    .addSeparator()
    .addItem("ダウンロード対象の電話番号を管理", "openPhoneManagerSidebar")
    .addSeparator()
    .addSubMenu(linkMenu)
    .addToUi();
}
