/**
 * プロダクトバックログ管理 — テンプレート自動生成（定数・共通UIヘルパー）
 *
 * Apps Script プロジェクトには本ファイルのほか、backlog-sheets.gs / epic-sheets.gs /
 * ids.gs / validation.gs を同じプロジェクトに追加する
 * （同一プロジェクト内ではファイルをまたいで関数・var を共有できるため import は不要）。
 *
 * 関数「createBacklogSheet」を実行すると、プロダクト名を確認して
 * バックログ_{名前} と エピック_{名前} の 2 シートを展開する。
 * 何度でも実行でき、名前を変えることで 1 つのスプレッドシートに複数のバックログを管理できる。
 * 同名シートが既にある場合は上書きせず、入力規則の適用とIDカウンタ同期のみを行う。
 */

var ID_SHEET_NAME = '🔢 ID管理';
/** 旧テンプレのブック共通 Epic シート名（後方互換）。 */
var LEGACY_EPIC_SHEET_NAME = 'Epic';
var BACKLOG_SHEET_PREFIX = 'バックログ_';
var EPIC_SHEET_PREFIX = 'エピック_';
/** 旧接頭辞（後方互換）。 */
var LEGACY_BACKLOG_SHEET_PREFIX = 'PBL_';
var LEGACY_EPIC_SHEET_PREFIX = 'Epic_';

/** バックログの列定義（1始まり）。列構成を変えるときはここを変更する。 */
var BACKLOG_COLUMNS = {
  ID: 1,
  EPIC: 2,
  STATUS: 3,
  DOABLE: 4,
  WHO: 5,
  WHAT: 6,
  WHY: 7,
  AC: 8,
  POINT: 9,
  PRD: 10,
  JIRA: 11,
  BUKO: 12,
};
var BACKLOG_COLUMN_COUNT = 12;

/** ヘッダー行のラベル（列順）。列名を変えるときはここを変更する。 */
var BACKLOG_HEADERS = ['ID', 'Epic', 'ステータス', '着手可能性', '誰が', '何をしたい', 'それはなぜか（価値）', '受け入れ条件（AC）', 'ポイント', 'PRD', 'JIRA', '備考'];

/** 列幅（px, 列順）。 */
var BACKLOG_COLUMN_WIDTHS = [90, 140, 100, 100, 120, 340, 340, 340, 70, 90, 110, 200];

/**
 * Epic シートの列定義（1始まり）。
 * ID / Epic名 / PRD / JIRA / 備考。
 */
var EPIC_COLUMNS = {
  ID: 1,
  NAME: 2,
  PRD: 3,
  JIRA: 4,
  BUKO: 5,
};
var EPIC_COLUMN_COUNT = 5;

/** Epic シートのヘッダー行のラベル（列順）。 */
var EPIC_HEADERS = ['ID', 'Epic名', 'PRD', 'JIRA', '備考'];

/** Epic シートの列幅（px, 列順）。 */
var EPIC_COLUMN_WIDTHS = [90, 200, 90, 110, 200];

/** ステータス・着手可能性・ポイントのプルダウン選択肢 */
var STATUS_OPTIONS = ['Open', 'In Sprint', 'Done', 'Closed'];
var DOABLE_OPTIONS = ['Ready', 'Not Ready'];
var POINT_OPTIONS = ['1', '2', '3', '5', '8'];

/** 条件付き書式の背景色マップ（キーはプルダウン選択肢の値）。 */
var DOABLE_COLORS = { 'Ready': '#d9ead3', 'Not Ready': '#fff2cc' };
var STATUS_COLORS = { 'Open': '#fff2cc', 'In Sprint': '#cfe2f3', 'Done': '#d9ead3', 'Closed': '#d9d9d9' };

/** プロダクト名からバックログ / エピック シート名を組み立てる。 */
function buildBacklogSheetName_(baseName) {
  return BACKLOG_SHEET_PREFIX + baseName;
}
function buildEpicSheetName_(baseName) {
  return EPIC_SHEET_PREFIX + baseName;
}

/** エピックシート名かどうか（エピック_… / 旧 Epic_… / 旧「Epic」）。 */
function isEpicSheetName_(sheetName) {
  let n = String(sheetName || '');
  return (
    n === LEGACY_EPIC_SHEET_NAME ||
    n.indexOf(EPIC_SHEET_PREFIX) === 0 ||
    n.indexOf(LEGACY_EPIC_SHEET_PREFIX) === 0
  );
}

/**
 * バックログシート名から接頭辞で導けるエピックシート名を返す。
 * バックログ_foo → エピック_foo。PBL_foo → Epic_foo（旧）。導けなければ空文字。
 */
function derivedEpicSheetName_(backlogSheetName) {
  let n = String(backlogSheetName || '');
  if (n.indexOf(BACKLOG_SHEET_PREFIX) === 0) {
    return EPIC_SHEET_PREFIX + n.substring(BACKLOG_SHEET_PREFIX.length);
  }
  if (n.indexOf(LEGACY_BACKLOG_SHEET_PREFIX) === 0) {
    return LEGACY_EPIC_SHEET_PREFIX + n.substring(LEGACY_BACKLOG_SHEET_PREFIX.length);
  }
  return '';
}

/**
 * 参照用に、バックログシートに対応する既存のエピックシート名を返す。
 * 接頭辞から導けない旧シートは、旧「Epic」が実在するときだけそれを使う。
 * 見つからなければ空文字を返す（呼び出し側で新規作成しないため）。
 */
function epicSheetNameForBacklog_(ss, backlogSheetName) {
  let derived = derivedEpicSheetName_(backlogSheetName);
  if (derived) return derived;
  return ss.getSheetByName(LEGACY_EPIC_SHEET_NAME) ? LEGACY_EPIC_SHEET_NAME : '';
}

/** 既存シートがバックログシートの体裁（A1 が ID）かどうか。 */
function looksLikeBacklogSheet_(sh) {
  if (!sh || sh.getLastRow() < 1 || sh.getLastColumn() < 1) return false;
  return String(sh.getRange(1, 1).getValue()).trim() === BACKLOG_HEADERS[0];
}

/** ブック内のエピックシート名一覧。 */
function getEpicSheetNames_(ss) {
  let sheets = ss.getSheets();
  let names = [];
  for (let i = 0; i < sheets.length; i++) {
    let n = sheets[i].getName();
    if (isEpicSheetName_(n)) names.push(n);
  }
  return names;
}

/**
 * バックログシートを 1 組展開する。プロダクト名はダイアログで確認する。
 * バックログ_{名前} と エピック_{名前} を作成する。名前を変えて再実行すると複数組を追加できる。
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

  const baseName = promptForProductBaseName_(ss);
  if (!baseName) return;

  const backlogName = buildBacklogSheetName_(baseName);
  const epicName = buildEpicSheetName_(baseName);

  const backlogExisting = ss.getSheetByName(backlogName);
  const epicExisting = ss.getSheetByName(epicName);
  const backlogIsNew = !backlogExisting;
  const epicIsNew = !epicExisting;

  setupEpicSheet(ss, epicName, epicIsNew);
  setupBacklogSheet(ss, backlogName, backlogIsNew);
  registerSheetIfNeeded_(ss, backlogName);

  SpreadsheetApp.flush();
  applyAllReferenceValidations_(ss);
  SpreadsheetApp.flush();
  applyReferenceColorFormats_(ss);
  SpreadsheetApp.flush();
  syncIdCountersFromBookCore(ss);
  SpreadsheetApp.flush();
  reorderBacklogTabs_(ss);

  ss.setActiveSheet(ss.getSheetByName(backlogName));

  const msg = backlogIsNew || epicIsNew
    ? '✅ シートを用意しました！\n\n・' + backlogName + (backlogIsNew ? '（新規）' : '（既存・保持）') +
      '\n・' + epicName + (epicIsNew ? '（新規）' : '（既存・保持）') +
      '\n\nプルダウン・ID自動採番が使えます。'
    : '✅ 既存の「' + backlogName + '」「' + epicName + '」を検出したため、データは保持したまま入力規則とIDカウンタを更新しました。';
  try {
    SpreadsheetApp.getUi().alert(msg);
  } catch (ignore) {
    Logger.log('createBacklogSheet: 完了ダイアログを表示できませんでした。');
  }
}

/**
 * ダイアログでプロダクト名（バックログ_ / エピック_ の共通部分）を入力してもらう。キャンセル時は null。
 * 先頭の接頭辞（新旧）は除去する。
 */
function promptForProductBaseName_(ss) {
  const ui = SpreadsheetApp.getUi();
  const result = ui.prompt(
    'プロダクト名',
    '共通の名前を入力してください。\n' +
      'バックログ_名前 と エピック_名前 の 2 シートを作成します。\n' +
      '（例：レセハブアプリ → バックログ_レセハブアプリ / エピック_レセハブアプリ）',
    ui.ButtonSet.OK_CANCEL
  );
  if (result.getSelectedButton() !== ui.Button.OK) return null;
  let name = String(result.getResponseText()).trim();
  const prefixes = [
    BACKLOG_SHEET_PREFIX,
    EPIC_SHEET_PREFIX,
    LEGACY_BACKLOG_SHEET_PREFIX,
    LEGACY_EPIC_SHEET_PREFIX,
  ];
  for (let i = 0; i < prefixes.length; i++) {
    if (name.indexOf(prefixes[i]) === 0) {
      name = name.substring(prefixes[i].length);
      break;
    }
  }
  name = name.trim();
  if (!name) {
    notifyUser_('名前が空です。', '名前');
    return null;
  }
  if (/[:\\\/\?\*\[\]]/.test(name)) {
    notifyUser_('名前に使用できない文字（: \\ / ? * [ ]）が含まれています。', '名前');
    return null;
  }
  const backlogName = buildBacklogSheetName_(name);
  const epicName = buildEpicSheetName_(name);
  if (backlogName.length > 100 || epicName.length > 100) {
    notifyUser_('シート名は 100 文字以内にしてください（接頭辞 バックログ_ / エピック_ を含む）。', '名前');
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

/** タブ順を整える（バックログ → 対応エピックをペアで並べ、最後に 🔢 ID管理）。 */
function reorderBacklogTabs_(ss) {
  let names = getBacklogSheetNames_(ss);
  let placed = {};
  let pos = 1;
  for (let i = 0; i < names.length; i++) {
    let pblSh = ss.getSheetByName(names[i]);
    if (pblSh) {
      ss.setActiveSheet(pblSh);
      ss.moveActiveSheet(pos);
      placed[names[i]] = true;
      pos++;
    }
    let epicName = epicSheetNameForBacklog_(ss, names[i]);
    let epicSh = epicName ? ss.getSheetByName(epicName) : null;
    if (epicSh) {
      ss.setActiveSheet(epicSh);
      ss.moveActiveSheet(pos);
      placed[epicName] = true;
      pos++;
    }
  }
  let epicNames = getEpicSheetNames_(ss);
  for (let i = 0; i < epicNames.length; i++) {
    if (placed[epicNames[i]]) continue;
    let orphan = ss.getSheetByName(epicNames[i]);
    if (!orphan) continue;
    ss.setActiveSheet(orphan);
    ss.moveActiveSheet(pos);
    pos++;
  }
  let idSh = ss.getSheetByName(ID_SHEET_NAME);
  if (idSh) {
    ss.setActiveSheet(idSh);
    ss.moveActiveSheet(ss.getNumSheets());
    try {
      idSh.hideSheet();
    } catch (e) {}
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
