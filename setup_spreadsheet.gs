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
 */

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
  const headers = ["phone_number", "password", "PDF保存先"];
  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setValues([headers]);

  // ヘッダーの書式設定
  headerRange
    .setFontWeight("bold")
    .setBackground("#4285F4")
    .setFontColor("#FFFFFF")
    .setHorizontalAlignment("center");

  // サンプルデータ（プレースホルダー）
  const sample = [
    ["090-1234-5678", "password123", "/Users/yamamoto/Library/CloudStorage/GoogleDrive-xxx/マイドライブ/確定申告系/2026/SoftBank請求書"],
  ];
  sheet.getRange(2, 1, sample.length, sample[0].length).setValues(sample);

  // 列幅の調整
  sheet.setColumnWidth(1, 180);  // phone_number
  sheet.setColumnWidth(2, 180);  // password
  sheet.setColumnWidth(3, 500);  // PDF保存先

  // phone_number列を書式なしテキストに（先頭0が消えないように）
  sheet.getRange("A:A").setNumberFormat("@");

  // シートの保護メモ
  sheet.getRange("A1").setNote(
    "SoftBank ID（携帯電話番号）。ハイフンOK（自動で除去される）\n例: 090-1234-5678"
  );
  sheet.getRange("B1").setNote("My SoftBankのログインパスワード");
  sheet.getRange("C1").setNote(
    "PDFの保存先フォルダのパス。\n" +
    "1行目に書けば全アカウント共通で使われる。\n" +
    "2行目以降は空欄でOK。\n" +
    "保存先を変えたいときはここを書き換えるだけ！"
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
    "2. 完了後、publishAsCsv() を実行してCSV URLを取得してください",
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}


/**
 * 「ウェブに公開」の手順をダイアログで表示する
 * ※ GASからプログラムで「ウェブに公開」はできないため、手順を案内する
 */
function publishAsCsv() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ssId = ss.getId();

  // 「ウェブに公開」後に生成されるURL形式の例
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
      <li>表示されたURLをコピー</li>
    </ol>
    <p>取得したURLを <code>.env</code> の <code>SPREADSHEET_CSV_URL</code> に設定してください。</p>
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
    .setHeight(420);

  SpreadsheetApp.getUi().showModalDialog(html, "CSV URLの取得手順");
}


/**
 * メニューにカスタムメニューを追加
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("MySoftBank ツール")
    .addItem("初期セットアップ", "setupSheet")
    .addItem("CSV URL 取得手順を表示", "publishAsCsv")
    .addToUi();
}
