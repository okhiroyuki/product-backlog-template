/**
 * プロダクトバックログ管理 — テンプレート自動生成（定数・共通UIヘルパー）
 *
 * Apps Script プロジェクトには本ファイルのほか、backlog-sheets.gs / ids.gs /
 * validation.gs を同じプロジェクトに追加する
 * （同一プロジェクト内ではファイルをまたいで関数・var を共有できるため import は不要）。
 *
 * 関数「createBacklogSheet」を実行すると、シート名を確認して 1 つのバックログシートを展開する。
 * 何度でも実行でき、シート名を変えることで 1 つのスプレッドシートに複数のバックログを管理できる。
 * 同名シートが既にある場合は上書きせず、入力規則の適用とIDカウンタ同期のみを行う。
 */

var ID_SHEET_NAME = '🔢 ID管理';

/** バックログの列定義（1始まり）。列構成を変えるときはここを変更する。 */
var BACKLOG_COLUMNS = {
  ID: 1,
  THEME: 2,
  STATUS: 3,
  DOABLE: 4,
  WHO: 5,
  WHAT: 6,
  WHY: 7,
  POINT: 8,
  PRD: 9,
  JIRA: 10,
};
var BACKLOG_COLUMN_COUNT = 10;

/** ステータス・着手可能性・ポイントのプルダウン選択肢 */
var STATUS_OPTIONS = ['Open', 'In Sprint', 'Done', 'Closed'];
var DOABLE_OPTIONS = ['Ready', 'Not Ready'];
var POINT_OPTIONS = ['1', '2', '3', '5', '8'];

/**
 * バックログシートを 1 つ展開する。シート名はダイアログで確認する。
 * シート名を変えて再実行すると、同じブックに複数のバックログを追加できる。
 */
function createBacklogSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.setSpreadsheetTimeZone('Asia/Tokyo');

  const defaultSheet = ss.getSheetByName('シート1');
  if (defaultSheet && ss.getSheets().length === 1) {
    const tmp = ss.insertSheet('_tmp');
    ss.deleteSheet(defaultSheet);
  }

  setupIdSheetHeaderOnly_(ss);

  const sheetName = promptForBacklogSheetName_(ss);
  if (!sheetName) return;

  const existing = ss.getSheetByName(sheetName);
  const isNew = !existing;

  setupBacklogSheet(ss, sheetName, isNew);
  registerSheetIfNeeded_(ss, sheetName);

  SpreadsheetApp.flush();
  applyAllReferenceValidations_(ss);
  SpreadsheetApp.flush();
  syncIdCountersFromBookCore(ss);
  SpreadsheetApp.flush();
  reorderBacklogTabs_(ss);

  ss.setActiveSheet(ss.getSheetByName(sheetName));

  const msg = isNew
    ? '✅ バックログシート「' + sheetName + '」を作成しました！\n\nプルダウン・ID自動採番が使えます。'
    : '✅ 既存のバックログシート「' + sheetName + '」を検出したため、データは保持したまま入力規則とIDカウンタを更新しました。';
  try {
    SpreadsheetApp.getUi().alert(msg);
  } catch (ignore) {
    Logger.log('createBacklogSheet: 完了ダイアログを表示できませんでした。');
  }
}

/** ダイアログでバックログシート名を入力してもらう。キャンセル時は null。 */
function promptForBacklogSheetName_(ss) {
  const ui = SpreadsheetApp.getUi();
  const result = ui.prompt(
    'バックログシートの名前',
    '新しいバックログシートの名前を入力してください。\n（例：レセハブアプリ、管理コンソール）',
    ui.ButtonSet.OK_CANCEL
  );
  if (result.getSelectedButton() !== ui.Button.OK) return null;
  const name = String(result.getResponseText()).trim();
  if (!name) {
    notifyUser_('シート名が空です。', '名前');
    return null;
  }
  if (/[:\\\/\?\*\[\]]/.test(name)) {
    notifyUser_('シート名に使用できない文字（: \\ / ? * [ ]）が含まれています。', '名前');
    return null;
  }
  if (name.length > 100) {
    notifyUser_('シート名は 100 文字以内にしてください。', '名前');
    return null;
  }
  return name;
}

function getOrCreateSheet(ss, name) {
  return ss.getSheetByName(name) || ss.insertSheet(name);
}

/** 値・書式・入力規則をクリアしてテンプレ再展開の前提にする。 */
function resetSheetCellsForTemplate_(sh, maxRows, maxCols) {
  if (!sh) return;
  sh.clearContents();
  sh.clearFormats();
  let rows = Math.min(Math.max(parseInt(maxRows, 10) || 500, 1), sh.getMaxRows());
  let cols = Math.min(Math.max(parseInt(maxCols, 10) || 40, 1), sh.getMaxColumns());
  try {
    sh.getRange(1, 1, rows, cols).clearDataValidations();
  } catch (e) {
    Logger.log('resetSheetCellsForTemplate_(' + sh.getName() + '): ' + (e && e.message ? e.message : e));
  }
}

/** タブ順を整える（バックログシートを先頭に、🔢 ID管理 を最後に）。 */
function reorderBacklogTabs_(ss) {
  let names = getBacklogSheetNames_(ss);
  let pos = 1;
  for (let i = 0; i < names.length; i++) {
    let sh = ss.getSheetByName(names[i]);
    if (sh) {
      ss.setActiveSheet(sh);
      ss.moveActiveSheet(pos);
      pos++;
    }
  }
  let idSh = ss.getSheetByName(ID_SHEET_NAME);
  if (idSh) {
    ss.setActiveSheet(idSh);
    ss.moveActiveSheet(ss.getNumSheets());
  }
}

/**
 * getUi() が使えないコンテキスト（サイドバーからのサーバー呼び出し等）でも落ちない通知。
 * まずダイアログ、無理ならトースト、それも無理なら Logger。
 */
function notifyUser_(message, title) {
  title = title || 'バックログ';
  try {
    SpreadsheetApp.getUi().alert(title ? title + '\n\n' + message : message);
    return;
  } catch (ignore) {}
  try {
    let ss = SpreadsheetApp.getActiveSpreadsheet();
    if (ss) ss.toast(message, title, 12);
  } catch (ignore2) {
    Logger.log('[' + title + '] ' + message);
  }
}

/** Toast で完了を通知する。 */
function toastDone_(message, title) {
  title = title || '完了';
  try {
    let ss = SpreadsheetApp.getActiveSpreadsheet();
    if (ss) ss.toast(message, title, 5);
  } catch (ignore) {
    Logger.log('[' + title + '] ' + message);
  }
}

/** ヘッダー行のスタイル設定。 */
function styleHeader(sheet, row, cols) {
  const range = sheet.getRange(row, 1, 1, cols);
  range
    .setBackground('#1a73e8')
    .setFontColor('#ffffff')
    .setFontWeight('bold')
    .setVerticalAlignment('middle');
  sheet.setFrozenRows(row);
}

/** 列幅を一括設定 */
function setColWidths(sheet, widths) {
  widths.forEach((w, i) => sheet.setColumnWidth(i + 1, w));
}

/** ドロップダウン検証（一覧から選択する入力規則） */
function setDropdown(sheet, row, col, values) {
  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(values, true)
    .setAllowInvalid(false)
    .build();
  sheet.getRange(row, col).setDataValidation(rule);
}
