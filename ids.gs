/** 🔢 ID管理 シートの読み書き・採番ロジック。バックログは PBL-001〜、Epic は EPC-001〜。 */

/** 🔢 ID管理：ヘッダのみ（syncIdCountersFromBookCore で中身を埋める）。既存データは保持。 */
function setupIdSheetHeaderOnly_(ss) {
  let sh = getOrCreateSheet(ss, ID_SHEET_NAME);
  if (sh.getLastRow() < 1) {
    sh.getRange(1, 1, 1, 3).setValues([['シート', '最終発番（数値）', '説明']]);
  }
  styleHeader(sh, 1, 3);
  setColWidths(sh, [160, 120, 200]);
  try {
    sh.hideSheet();
  } catch (e) {}
}

/**
 * メニュー「🔢 IDカウンタをブックから再同期」用。
 * 全バックログシートと Epic シートの ID を走査して 🔢 ID管理 を更新する。
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
  let rows = [['シート', '最終発番（数値）', '説明']];
  let seen = {};
  for (let i = 0; i < names.length; i++) {
    if (!ss.getSheetByName(names[i])) continue;
    rows.push([names[i], scanMaxIdFromSheet(ss, names[i], 'PBL'), 'バックログ項目']);
    seen[names[i]] = true;
    let epicName = epicSheetNameForBacklog_(ss, names[i]);
    if (epicName && ss.getSheetByName(epicName) && !seen[epicName]) {
      rows.push([epicName, scanMaxIdFromSheet(ss, epicName, 'EPC'), 'Epic']);
      seen[epicName] = true;
    }
  }
  let epicNames = getEpicSheetNames_(ss);
  for (let i = 0; i < epicNames.length; i++) {
    if (seen[epicNames[i]]) continue;
    rows.push([epicNames[i], scanMaxIdFromSheet(ss, epicNames[i], 'EPC'), 'Epic']);
  }
  resetSheetCellsForTemplate_(sh);
  sh.getRange(1, 1, rows.length, 3).setValues(rows);
  styleHeader(sh, 1, 3);
}

/** 🔢 ID管理 に登録されているバックログシート名の一覧を返す（Epic・ID管理は除く）。 */
function getBacklogSheetNames_(ss) {
  let sh = ss.getSheetByName(ID_SHEET_NAME);
  if (!sh) return [];
  let lr = sh.getLastRow();
  if (lr < 2) return [];
  let vals = sh.getRange(2, 1, lr, 1).getValues();
  let names = [];
  for (let i = 0; i < vals.length; i++) {
    let t = String(vals[i][0]).trim();
    if (t === '' || t === ID_SHEET_NAME || isEpicSheetName_(t)) continue;
    names.push(t);
  }
  return names;
}

/** バックログシート名を 🔢 ID管理 に登録する（未登録のときのみ）。 */
function registerSheetIfNeeded_(ss, sheetName) {
  if (!sheetName || sheetName === ID_SHEET_NAME || isEpicSheetName_(sheetName)) return;
  let sh = ss.getSheetByName(ID_SHEET_NAME);
  if (!sh) return;
  let names = getBacklogSheetNames_(ss);
  if (names.indexOf(sheetName) >= 0) return;
  sh.appendRow([sheetName, 0, 'バックログ項目']);
}

/**
 * 指定シートの ID 列から、指定プレフィックス（PBL / EPC）の最大連番を返す。
 */
function scanMaxIdFromSheet(ss, sheetName, prefix) {
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) return 0;
  let lr = sheet.getLastRow();
  if (lr < 2) return 0;
  let idCol = isEpicSheetName_(sheetName) ? EPIC_COLUMNS.ID : BACKLOG_COLUMNS.ID;
  let vals = sheet.getRange(2, idCol, lr, 1).getValues();
  let re = new RegExp('^' + prefix + '-(\\d+)$');
  let max = 0;
  for (let i = 0; i < vals.length; i++) {
    let m = String(vals[i][0]).trim().match(re);
    if (m) {
      let n = parseInt(m[1], 10);
      if (n > max) max = n;
    }
  }
  return max;
}

/**
 * ロック付きで、指定シートの連番を +1 し、表示用 ID 文字列を返す。
 * prefix は 'PBL' または 'EPC'。🔢 ID管理 の該当行を更新する。
 */
function issueNextId(ss, sheetName, prefix) {
  prefix = prefix || 'PBL';
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
      let desc = isEpicSheetName_(sheetName) ? 'Epic' : 'バックログ項目';
      sh.appendRow([sheetName, 0, desc]);
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

    return formatPrefixedId_(prefix, next);
  } finally {
    lock.releaseLock();
  }
}

function formatPrefixedId_(prefix, num) {
  let n = Number(num);
  if (isNaN(n) || n < 1) throw new Error('不正な連番: ' + num);
  let s = String(n);
  let pad = s.length < 3 ? ('000' + s).slice(-3) : s;
  return prefix + '-' + pad;
}

function formatBacklogId(num) {
  return formatPrefixedId_('PBL', num);
}

/**
 * 編集時に自動発火するトリガー。
 * 登録済みのバックログ／Epic シートのデータ行に ID がなく内容が入っている場合、自動採番する。
 */
function onEdit(e) {
  const range = e.range;
  const sheet = range.getSheet();
  const sheetName = sheet.getName();
  if (sheetName === ID_SHEET_NAME) return;

  const row = range.getRow();
  if (row < 2) return;

  const isEpic = isEpicSheetName_(sheetName);
  const isBacklog = getBacklogSheetNames_(e.source).indexOf(sheetName) >= 0;
  if (!isEpic && !isBacklog) return;

  const idCol = isEpic ? EPIC_COLUMNS.ID : BACKLOG_COLUMNS.ID;
  const colCount = isEpic ? EPIC_COLUMN_COUNT : BACKLOG_COLUMN_COUNT;
  const prefix = isEpic ? 'EPC' : 'PBL';

  const idCell = sheet.getRange(row, idCol);
  if (String(idCell.getValue()).trim() !== '') return;

  const rowData = sheet.getRange(row, 1, 1, colCount).getValues()[0];
  const hasContent = rowData.some(function (v) {
    return String(v).trim() !== '';
  });
  if (!hasContent) return;

  try {
    const id = issueNextId(e.source, sheetName, prefix);
    idCell.setValue(id);
    if (!isEpic) {
      applyRowValidations_(
        sheet,
        row,
        getEpicNameRange_(e.source, epicSheetNameForBacklog_(e.source, sheetName))
      );
    }
  } catch (err) {
    Logger.log('onEdit: ID採番に失敗しました。 ' + (err && err.message ? err.message : err));
  }
}
