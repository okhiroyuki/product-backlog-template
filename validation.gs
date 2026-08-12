/** ドロップダウン（入力規則）の適用。ステータス・着手可能性・ポイント。 */

/** 登録済みの全バックログシートのデータ行すべてに入力規則を適用する。 */
function applyAllReferenceValidations_(ss) {
  let names = getBacklogSheetNames_(ss);
  for (let i = 0; i < names.length; i++) {
    let sh = ss.getSheetByName(names[i]);
    if (!sh) continue;
    let lr = sh.getLastRow();
    for (let r = 2; r <= lr; r++) {
      applyRowValidations_(sh, r);
    }
  }
}

/** 1 行分の入力規則（ステータス・着手可能性・ポイント）を適用する。 */
function applyRowValidations_(sh, row) {
  setDropdown(sh, row, BACKLOG_COLUMNS.STATUS, STATUS_OPTIONS);
  setDropdown(sh, row, BACKLOG_COLUMNS.DOABLE, DOABLE_OPTIONS);
  setDropdown(sh, row, BACKLOG_COLUMNS.POINT, POINT_OPTIONS);
}
