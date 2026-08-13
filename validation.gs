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

/** 全バックログシートの着手可能性・ステータス列に条件付き書式（背景色マップ）を適用する。 */
function applyReferenceColorFormats_(ss) {
  applyColumnColorFormats_(ss, BACKLOG_COLUMNS.DOABLE, DOABLE_COLORS);
  applyColumnColorFormats_(ss, BACKLOG_COLUMNS.STATUS, STATUS_COLORS);
}

/** 指定列の背景色マップに基づく条件付き書式を全バックログシートへ適用する。 */
function applyColumnColorFormats_(ss, col, colorMap) {
  let names = getBacklogSheetNames_(ss);
  for (let i = 0; i < names.length; i++) {
    let sh = ss.getSheetByName(names[i]);
    if (!sh) continue;
    applyColumnColorFormat_(sh, col, colorMap);
  }
}

/** 1 シート・1 列分の条件付き書式を適用する（値の一致で背景色を変える）。 */
function applyColumnColorFormat_(sh, col, colorMap) {
  const range = sh.getRange(2, col, sh.getMaxRows() - 1, 1);
  const rules = [];
  Object.keys(colorMap).forEach(function (value) {
    rules.push(
      SpreadsheetApp.newConditionalFormatRule()
        .whenTextEqualTo(value)
        .setBackground(colorMap[value])
        .setRanges([range])
        .build()
    );
  });
  sh.setConditionalFormatRules(rules);
}
