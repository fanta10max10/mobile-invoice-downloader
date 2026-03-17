/**
 * 携帯領収書管理スプレッドシート セットアップスクリプト
 *
 * 【このスクリプトについて】
 * SoftBank・Y!mobile 両キャリアで1つのスプレッドシートを共用するための統合スクリプト。
 * 以下のシートを作成・初期化します:
 *
 *   設定           ... PDF保存先フォルダ、対象月、パスワード（全キャリア共通）
 *   認証情報       ... 電話番号・キャリア・PDFの種類（HTMLサイドバーで管理）
 *   SoftBankリンク ... SoftBank PDFの月別リンク管理（GAS自動更新）
 *   Ymobileリンク  ... Ymobile PDFの月別リンク管理（GAS自動更新）
 *
 * 【使い方】
 *   1. 携帯領収書管理スプレッドシートを開く
 *   2. 拡張機能 → Apps Script を開く
 *   3. このコードを貼り付けて保存
 *   4. setupSheet() を実行（初回は権限承認が必要）
 *   5. 「設定」シートにPDF保存先フォルダURL・パスワード・対象月を入力
 *   6. メニューから「ダウンロード対象の電話番号を管理」でサイドバーを開き、対象番号を選択
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

// リンクシートの月列開始位置（A=電話番号 B=PDFの種類 → C列から月）
const MONTH_COL_START = 3;

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

  setupSettingsSheet_(ss);
  setupAuthSheet_(ss);
  setupLinkSheet_(ss, SOFTBANK_LINK_SHEET_NAME, "SoftBank");
  setupLinkSheet_(ss, YMOBILE_LINK_SHEET_NAME, "Ymobile");

  // 旧 _PhoneDropdown シートがあれば削除
  const oldPd = ss.getSheetByName("_PhoneDropdown");
  if (oldPd) ss.deleteSheet(oldPd);

  SpreadsheetApp.getUi().alert(
    "セットアップ完了",
    "以下のシートを作成・初期化しました。\n\n" +
    "【設定シート】\n" +
    "  「PDF保存先フォルダ」にGoogle DriveのフォルダURLを入力\n" +
    "  「パスワード」にログインパスワードを入力（全番号共通）\n" +
    "  「対象月」をドロップダウンで選択\n\n" +
    "【ダウンロード対象の管理】\n" +
    "  メニューから「ダウンロード対象の電話番号を管理」を開いて\n" +
    "  月別シートの電話番号から対象を選択してください\n\n" +
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

  _upsertSettingRow_(sheet, "このスプレッドシートURL",
    ss.getUrl(),
    "このスプレッドシートのURL。\nSoftBank/.env および Ymobile/.env の SPREADSHEET_URL にコピーしてください。"
  );

  _upsertSettingRow_(sheet, "回線管理スプレッドシート",
    "",
    "回線管理表（月別シートに電話番号・解約済・運用端末等がある）のURL。\n" +
    "サイドバーで電話番号を読み込む際に参照します。\n" +
    "空欄の場合はこのスプレッドシート内の月別シートを検索します。"
  );

  _upsertSettingRow_(sheet, "PDF保存先フォルダ",
    "https://drive.google.com/drive/folders/XXXXXXXX",
    "PDFの保存先フォルダ。以下どちらでも可:\n" +
    "・Google DriveのフォルダURL（推奨）\n" +
    "  https://drive.google.com/drive/folders/...\n" +
    "  → Drive APIで直接アップロード（要: サービスアカウントを編集者として共有）\n" +
    "・ローカル絶対パス（/Users/... または C:\\...）"
  );

  _upsertSettingRow_(sheet, "パスワード",
    "",
    "SoftBank / Y!mobile 共通のログインパスワード。\n全電話番号で同じパスワードを使用します。"
  );

  _upsertSettingRow_(sheet, "対象月",
    "自動（前月）",
    "ダウンロードする月を選択してください\n「自動（前月）」= 実行時の前月を自動選択"
  );

  _setTargetMonthValidation_(sheet);
}


/**
 * 認証情報シートを作成・初期化する（内部用）
 * 列: 電話番号 | キャリア | PDFの種類
 * パスワードは設定シートで一元管理するため、このシートには含まない。
 */
function setupAuthSheet_(ss) {
  let sheet = ss.getSheetByName(AUTH_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(AUTH_SHEET_NAME);
    ss.setActiveSheet(sheet);
    ss.moveActiveSheet(2);
  }

  // 旧形式（パスワード列あり）からの移行
  const currentHeaders = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const pwColIdx = currentHeaders.findIndex(h => String(h).trim() === "パスワード");
  if (pwColIdx !== -1) {
    sheet.deleteColumn(pwColIdx + 1);
  }

  // ヘッダー行（空の場合のみ書き込む）
  if (!sheet.getRange(1, 1).getValue()) {
    const headers = ["電話番号", "キャリア", "PDFの種類"];
    const headerRange = sheet.getRange(1, 1, 1, headers.length);
    headerRange.setValues([headers]);
    headerRange
      .setFontWeight("bold")
      .setBackground("#4285F4")
      .setFontColor("#FFFFFF")
      .setHorizontalAlignment("center");
  }

  sheet.setColumnWidth(1, 180);  // 電話番号
  sheet.setColumnWidth(2, 120);  // キャリア
  sheet.setColumnWidth(3, 220);  // PDFの種類

  // 電話番号列を書式なしテキストに（先頭0が消えないように）
  sheet.getRange("A:A").setNumberFormat("@");

  // ツールチップ
  sheet.getRange("A1").setNote("電話番号。ハイフンOK（スクリプトが自動で除去）\nサイドバーから管理できます。");
  sheet.getRange("B1").setNote("キャリア名（SoftBank または Ymobile）");
  sheet.getRange("C1").setNote(
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
//  電話番号管理 HTMLサイドバー
// ────────────────────────────────────────────────

/**
 * サイドバーを開く（メニューから呼び出す）
 */
function openPhoneManagerSidebar() {
  const html = HtmlService.createHtmlOutput(_getPhoneManagerHtml_())
    .setTitle("ダウンロード対象の電話番号")
    .setWidth(320);
  SpreadsheetApp.getUi().showSidebar(html);
}


/**
 * サイドバー用HTML（google.script.run でサーバー関数を呼ぶ）
 */
function _getPhoneManagerHtml_() {
  return `<!DOCTYPE html>
<html>
<head>
<style>
  * { box-sizing: border-box; margin: 0; padding: 0; }
  body { font-family: "Google Sans", Arial, sans-serif; font-size: 13px; color: #333; padding: 12px; }
  h3 { font-size: 14px; margin-bottom: 8px; }
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
    border-bottom: 1px solid #f0f0f0; transition: background 0.15s;
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
  .btn-bar { margin-top: 12px; display: flex; gap: 8px; }
  .btn {
    flex: 1; padding: 10px; border: none; border-radius: 6px;
    font-size: 13px; font-weight: bold; cursor: pointer; transition: opacity 0.2s;
  }
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
  let phoneData = {};       // { carrier: [{ phone, cancelled, device }] }
  let selections = {};      // { carrier: { phone: { checked, pdfType } } }
  let allDataLoaded = false;

  function setStatus(msg, type) {
    const el = document.getElementById("status");
    el.textContent = msg;
    el.className = "status " + type;
  }

  function clearStatus() {
    document.getElementById("status").className = "status";
  }

  // 初期読み込み
  function init() {
    setStatus("月別シートから電話番号を読み込み中...", "loading");
    google.script.run
      .withSuccessHandler(function(result) {
        phoneData = result.phones;       // { SoftBank: [...], Ymobile: [...] }
        selections = result.selections;  // { SoftBank: {phone: {pdfType}}, ... }
        allDataLoaded = true;
        clearStatus();
        render();
      })
      .withFailureHandler(function(err) {
        setStatus("読み込みエラー: " + err.message, "error");
      })
      .getPhoneManagerData();
  }

  function switchTab(carrier) {
    currentCarrier = carrier;
    document.querySelectorAll(".tab-btn").forEach(function(btn) {
      btn.classList.toggle("active", btn.dataset.carrier === carrier);
    });
    render();
  }

  function render() {
    const list = document.getElementById("phoneList");
    const phones = phoneData[currentCarrier] || [];
    const sel = selections[currentCarrier] || {};

    if (phones.length === 0) {
      list.innerHTML = '<div class="empty-msg">月別シートに ' + currentCarrier + ' の電話番号が見つかりません。<br>月別シート（例: 3月）にキャリア・電話番号列があるか確認してください。</div>';
      updateSummary();
      return;
    }

    let html = "";
    phones.forEach(function(p) {
      const checked = sel[p.phone] ? "checked" : "";
      const pdfType = (sel[p.phone] && sel[p.phone].pdfType) || "電話番号別";
      const cls = p.cancelled ? "cancelled" : "";
      const label = p.device ? p.phone + " (" + p.device + ")" : p.phone;

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
      const selEl = el.parentElement.querySelector("select");
      selections[currentCarrier][el.dataset.phone] = {
        pdfType: selEl ? selEl.value : "電話番号別"
      };
    } else {
      delete selections[currentCarrier][el.dataset.phone];
    }
    updateSummary();
  }

  function onPdfType(el) {
    if (!selections[currentCarrier]) selections[currentCarrier] = {};
    if (selections[currentCarrier][el.dataset.phone]) {
      selections[currentCarrier][el.dataset.phone].pdfType = el.value;
    }
  }

  function toggleAll() {
    const phones = phoneData[currentCarrier] || [];
    const sel = selections[currentCarrier] || {};
    const activePhones = phones.filter(function(p) { return !p.cancelled; });
    const allChecked = activePhones.every(function(p) { return sel[p.phone]; });

    if (!selections[currentCarrier]) selections[currentCarrier] = {};
    if (allChecked) {
      activePhones.forEach(function(p) { delete selections[currentCarrier][p.phone]; });
    } else {
      activePhones.forEach(function(p) {
        if (!selections[currentCarrier][p.phone]) {
          selections[currentCarrier][p.phone] = { pdfType: "電話番号別" };
        }
      });
    }
    render();
  }

  function updateSummary() {
    var total = 0;
    ["SoftBank", "Ymobile"].forEach(function(c) {
      total += Object.keys(selections[c] || {}).length;
    });
    document.getElementById("summary").textContent =
      "選択中: SoftBank " + Object.keys(selections["SoftBank"] || {}).length +
      "件 / Ymobile " + Object.keys(selections["Ymobile"] || {}).length + "件";
  }

  function save() {
    var btn = document.getElementById("saveBtn");
    btn.disabled = true;
    setStatus("保存中...", "loading");

    google.script.run
      .withSuccessHandler(function() {
        setStatus("保存しました。", "success");
        btn.disabled = false;
      })
      .withFailureHandler(function(err) {
        setStatus("保存エラー: " + err.message, "error");
        btn.disabled = false;
      })
      .savePhoneSelections(selections);
  }

  init();
</script>
</body>
</html>`;
}


/**
 * サイドバーの初期データを返す（google.script.run から呼ばれる）
 * @return {{ phones: Object, selections: Object }}
 */
function getPhoneManagerData() {
  const phones = _getAllPhonesFromMonthSheets_();
  const selections = _getCurrentSelections_();
  return { phones: phones, selections: selections };
}


/**
 * サイドバーからの保存（google.script.run から呼ばれる）
 * @param {Object} selections  { "SoftBank": { "09012345678": { pdfType: "電話番号別" } }, ... }
 */
function savePhoneSelections(selections) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(AUTH_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(AUTH_SHEET_NAME);
  }

  // ヘッダー確認
  if (!sheet.getRange(1, 1).getValue()) {
    const headers = ["電話番号", "キャリア", "PDFの種類"];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length)
      .setFontWeight("bold")
      .setBackground("#4285F4")
      .setFontColor("#FFFFFF")
      .setHorizontalAlignment("center");
  }

  // 既存データをクリア（ヘッダー行以外）
  if (sheet.getLastRow() > 1) {
    sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).clearContent();
  }

  // 選択されたデータを書き込む
  const rows = [];
  for (const carrier of ["SoftBank", "Ymobile"]) {
    const carrierSel = selections[carrier] || {};
    for (const phone of Object.keys(carrierSel)) {
      rows.push([phone, carrier, carrierSel[phone].pdfType || "電話番号別"]);
    }
  }

  if (rows.length > 0) {
    sheet.getRange(2, 1, rows.length, 3).setValues(rows);
  }

  // 電話番号列をテキスト形式に
  sheet.getRange("A:A").setNumberFormat("@");
}


/**
 * キャリアラベル文字列を正規化する。
 * 「ソフトバンク」→ "SoftBank", 「ワイモバイル」→ "Ymobile" 等。
 * 不明なら null を返す。
 */
function _normalizeCarrierName_(text) {
  const s = String(text).trim().toLowerCase();
  if (!s) return null;
  if (s === "softbank" || s.includes("ソフトバンク")) return "SoftBank";
  if (s === "ymobile" || s === "y!mobile" || s.includes("ワイモバイル")) return "Ymobile";
  return null;
}


/**
 * 月別シートから全キャリアの電話番号を収集する（内部用）
 *
 * 月別シートの構造（複数セクション対応）:
 *   ワイモバイル              ← セクションラベル行（キャリア名）
 *   電話番号 | 解約済 | ...   ← ヘッダー行
 *   080-xxxx | FALSE  | ...   ← データ行
 *   ...
 *   ソフトバンク              ← 次のセクションラベル行
 *   電話番号 | 解約済 | ...   ← ヘッダー行
 *   090-xxxx | FALSE  | ...   ← データ行
 *
 * @return {{ SoftBank: [{phone, cancelled, device}], Ymobile: [...] }}
 */
function _getAllPhonesFromMonthSheets_() {
  const result = { SoftBank: [], Ymobile: [] };

  // 回線管理スプレッドシートを開く（設定があれば外部、なければ自身）
  let ss;
  const mgmtUrl = _getSettingValue_("回線管理スプレッドシート");
  if (mgmtUrl) {
    const idMatch = mgmtUrl.match(/\/spreadsheets\/d\/([a-zA-Z0-9_-]+)/);
    if (!idMatch) {
      Logger.log("[_getAllPhonesFromMonthSheets_] 回線管理スプレッドシートURLからIDを取得できません");
      return result;
    }
    try {
      ss = SpreadsheetApp.openById(idMatch[1]);
    } catch (e) {
      Logger.log(`[_getAllPhonesFromMonthSheets_] 回線管理スプレッドシートを開けません: ${e.message}`);
      return result;
    }
  } else {
    ss = SpreadsheetApp.getActiveSpreadsheet();
  }

  // 月別シートを最新順に検索
  const monthSheets = ss.getSheets().filter(ws => /.*\d+月$/.test(ws.getName()));
  if (monthSheets.length === 0) return result;

  monthSheets.sort((a, b) => _parseMonthSheetNum_(a.getName()) - _parseMonthSheetNum_(b.getName()));

  // 最新のデータがあるシートを使用
  let targetSheet = null;
  for (let i = monthSheets.length - 1; i >= 0; i--) {
    if (monthSheets[i].getLastRow() > 1) { targetSheet = monthSheets[i]; break; }
  }
  if (!targetSheet) return result;

  Logger.log(`[_getAllPhonesFromMonthSheets_] 使用シート: ${targetSheet.getName()}`);

  const rows = targetSheet.getDataRange().getValues();
  let cols = null;           // 現在のセクションの列マッピング
  let sectionCarrier = null; // 現在のセクションのキャリア名

  for (let i = 0; i < rows.length; i++) {
    const row = rows[i];

    // セクションラベル行の検出（セルにキャリア名が含まれるか）
    // 例: ["ワイモバイル", "", "", ...] や ["", "ソフトバンク", ...]
    const labelCarrier = _detectCarrierLabel_(row);
    if (labelCarrier) {
      sectionCarrier = labelCarrier;
      cols = null; // 次のヘッダー行を待つ
      continue;
    }

    // ヘッダー行を検出（「電話番号」を含む行）→ 新セクション開始
    const phoneColIdx = row.findIndex(c => String(c).trim() === "電話番号");
    if (phoneColIdx !== -1) {
      cols = {};
      row.forEach((c, j) => { const k = String(c).trim(); if (k) cols[k] = j; });
      continue;
    }

    // データ行の処理
    if (!cols || cols["電話番号"] === undefined) continue;

    const phone = String(row[cols["電話番号"]] || "").replace(/[-\s]/g, "").trim();
    if (!phone || !/^\d{10,13}$/.test(phone)) continue;

    // キャリア判定: キャリア列があればそこから、なければセクションラベルから
    let carrier = null;
    if (cols["キャリア"] !== undefined) {
      carrier = _normalizeCarrierName_(row[cols["キャリア"]]);
    }
    if (!carrier) carrier = sectionCarrier;
    if (!carrier || !result[carrier]) continue;

    const ci = cols["解約済"];
    const cancelled = ci !== undefined && String(row[ci] || "").toUpperCase() === "TRUE";

    const di = cols["運用端末"];
    const device = di !== undefined ? String(row[di] || "").trim() : "";

    // 重複チェック
    if (!result[carrier].some(p => p.phone === phone)) {
      result[carrier].push({ phone: phone, cancelled: cancelled, device: device });
    }
  }

  Logger.log(`[_getAllPhonesFromMonthSheets_] 結果: SoftBank ${result.SoftBank.length}件, Ymobile ${result.Ymobile.length}件`);
  return result;
}


/**
 * 行からキャリアラベル（セクション見出し）を検出する。
 * 行内の非空セルが少なく（1〜2個）、かつキャリア名を含む場合にキャリア名を返す。
 * データ行やヘッダー行と誤検知しないよう、非空セルが少ない行のみ対象とする。
 */
function _detectCarrierLabel_(row) {
  const nonEmpty = row.filter(c => String(c).trim() !== "");
  if (nonEmpty.length === 0 || nonEmpty.length > 3) return null;

  // 「電話番号」を含む行はヘッダーなのでスキップ
  if (nonEmpty.some(c => String(c).trim() === "電話番号")) return null;

  for (const cell of nonEmpty) {
    const carrier = _normalizeCarrierName_(cell);
    if (carrier) return carrier;
  }
  return null;
}


/**
 * 認証情報シートの現在の選択状態を返す（内部用）
 * @return {{ SoftBank: { phone: { pdfType } }, Ymobile: { ... } }}
 */
function _getCurrentSelections_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(AUTH_SHEET_NAME);
  const result = { SoftBank: {}, Ymobile: {} };
  if (!sheet || sheet.getLastRow() <= 1) return result;

  const data = sheet.getDataRange().getValues();
  const headers = data[0].map(h => String(h).trim());
  const phoneIdx = headers.indexOf("電話番号");
  const carrierIdx = headers.indexOf("キャリア");
  const pdfTypeIdx = headers.indexOf("PDFの種類");
  // 旧形式にパスワード列がある場合もフォールバック
  if (phoneIdx === -1) return result;

  for (let i = 1; i < data.length; i++) {
    const phone = String(data[i][phoneIdx] || "").replace(/[-\s]/g, "").trim();
    if (!phone) continue;
    const carrier = carrierIdx !== -1 ? String(data[i][carrierIdx] || "").trim() : "";
    if (!carrier || !result[carrier]) continue;
    const pdfType = pdfTypeIdx !== -1 ? String(data[i][pdfTypeIdx] || "").trim() : "電話番号別";
    result[carrier][phone] = { pdfType: pdfType || "電話番号別" };
  }

  return result;
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
//  メニュー
// ────────────────────────────────────────────────

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("携帯領収書管理 ツール")
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
