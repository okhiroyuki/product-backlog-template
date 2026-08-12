/** 🔢 ID管理 シートの読み書き・採番ロジック。バックログシートごとに独立した連番（PBL-001〜）を振る。 */

/** 🔢 ID管理：ヘッダのみ（syncIdCountersFromBookCore で中身を埋める）。既存データは保持。 */
function setupIdSheetHeaderOnly_(ss) {
  let sh = getOrCreateSheet(ss, ID_SHEET_NAME);
  if (sh.getLastRow() < 1) {
    sh.getRange(1, 1, 1, 3).setValues([['バックログシート', '最終発番（数値）', '説明']]);
  }
  styleHeader(sh, 1, 3);
  setColWidths(sh, [160, 120, 200]);
  try {
    sh.hideSheet();
  } catch (e) {}
}

/**
 * メニュー「🔢 IDカウンタをブックから再同期」用。
 * 全バックログシートの ID を走査して 🔢 ID管理 を更新する。
 */
function syncIdCountersFromBook() {
  let ss = SpreadsheetApp.getActiveSpreadsheet();
  syncIdCountersFromBookCore(ss);
  if (ss.getSheetByName(ID_SHEET_NAME)) {
    toastDone_('🔢 ID管理をブック内の ID に合わせました', '再同期');
  }
}

function syncIdCountersFromBookCore(ss) {
  let sh = ss.getSheetByName(ID_SHEET_NAME);
  if (!sh) {
    notifyUser_('🔢 ID管理 シートがありません。先に createBacklogSheet を実行してください。', 'ID 管理');
    return;
  }
  let names = getBacklogSheetNames_(ss);
  let rows = [['バックログシート', '最終発番（数値）', '説明']];
  for (let i = 0; i < names.length; i++) {
    rows.push([names[i], scanMaxPblIdFromSheet(ss, names[i]), 'バックログ項目']);
  }
  resetSheetCellsForTemplate_(sh);
  sh.getRange(1, 1, rows.length, 3).setValues(rows);
  styleHeader(sh, 1, 3);
}

/** 🔢 ID管理 に登録されているバックログシート名の一覧を返す。 */
function getBacklogSheetNames_(ss) {
  let sh = ss.getSheetByName(ID_SHEET_NAME);
  if (!sh) return [];
  let lr = sh.getLastRow();
  if (lr < 2) return [];
  let vals = sh.getRange(2, 1, lr - 1, 1).getValues();
  let names = [];
  for (let i = 0; i < vals.length; i++) {
    let t = String(vals[i][0]).trim();
    if (t !== '') names.push(t);
  }
  return names;
}

/** バックログシート名を 🔢 ID管理 に登録する（未登録のときのみ）。 */
function registerSheetIfNeeded_(ss, sheetName) {
  let sh = ss.getSheetByName(ID_SHEET_NAME);
  if (!sh) return;
  let names = getBacklogSheetNames_(ss);
  if (names.indexOf(sheetName) >= 0) return;
  sh.appendRow([sheetName, 0, 'バックログ項目']);
}

/** 指定シートの ID 列から最大の PBL 連番を返す。 */
function scanMaxPblIdFromSheet(ss, sheetName) {
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) return 0;
  let lr = sheet.getLastRow();
  if (lr < 2) return 0;
  let vals = sheet.getRange(2, BACKLOG_COLUMNS.ID, lr - 1, 1).getValues();
  let max = 0;
  for (let i = 0; i < vals.length; i++) {
    let m = String(vals[i][0]).trim().match(/^PBL-(\d+)$/);
    if (m) {
      let n = parseInt(m[1], 10);
      if (n > max) max = n;
    }
  }
  return max;
}

/**
 * ロック付きで、指定バックログシートの連番を +1 し、表示用 ID 文字列（PBL-001）を返す。
 * 🔢 ID管理 の該当シート行を更新する。
 */
function issueNextId(ss, sheetName) {
  let lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    let sh = ss.getSheetByName(ID_SHEET_NAME);
    if (!sh) throw new Error('ID管理シートがありません');

    let data = sh.getDataRange().getValues();
    let row = -1;
    for (let r = 1; r < data.length; r++) {
      if (String(data[r][0]) === sheetName) {
        row = r;
        break;
      }
    }
    if (row < 0) {
      sh.appendRow([sheetName, 0, 'バックログ項目']);
      data = sh.getDataRange().getValues();
      for (let r = 1; r < data.length; r++) {
        if (String(data[r][0]) === sheetName) {
          row = r;
          break;
        }
      }
    }
    let last = Number(data[row][1]);
    if (isNaN(last)) last = 0;
    let next = last + 1;
    sh.getRange(row + 1, 2).setValue(next);

    return formatBacklogId(next);
  } finally {
    lock.releaseLock();
  }
}

function formatBacklogId(num) {
  let n = Number(num);
  if (isNaN(n) || n < 1) throw new Error('不正な連番: ' + num);
  let s = String(n);
  let pad = s.length < 3 ? ('000' + s).slice(-3) : s;
  return 'PBL-' + pad;
}

/**
 * 編集時に自動発火するトリガー。
 * 登録済みのバックログシートのデータ行に ID がなく内容が入っている場合、自動採番する。
 */
function onEdit(e) {
  const range = e.range;
  const sheet = range.getSheet();
  if (sheet.getName() === ID_SHEET_NAME) return;
  if (getBacklogSheetNames_(e.source).indexOf(sheet.getName()) < 0) return;
  const row = range.getRow();
  if (row < 2) return;

  const idCell = sheet.getRange(row, BACKLOG_COLUMNS.ID);
  if (String(idCell.getValue()).trim() !== '') return;

  const rowData = sheet.getRange(row, 1, 1, BACKLOG_COLUMN_COUNT).getValues()[0];
  const hasContent = rowData.some(function (v) {
    return String(v).trim() !== '';
  });
  if (!hasContent) return;

  try {
    const id = issueNextId(e.source, sheet.getName());
    idCell.setValue(id);
    applyRowValidations_(sheet, row);
  } catch (err) {
    Logger.log('onEdit: ID採番に失敗しました。 ' + (err && err.message ? err.message : err));
  }
}
