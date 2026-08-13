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
  let columnColorMaps = [
    { col: BACKLOG_COLUMNS.DOABLE, colorMap: DOABLE_COLORS },
    { col: BACKLOG_COLUMNS.STATUS, colorMap: STATUS_COLORS },
  ];
  let names = getBacklogSheetNames_(ss);
  for (let i = 0; i < names.length; i++) {
    let sh = ss.getSheetByName(names[i]);
    if (!sh) continue;
    applyColumnColorFormats_(sh, columnColorMaps);
  }
}

/**
 * 1 シートへ複数列分の条件付き書式をまとめて適用する。
 * setConditionalFormatRules はシート全体を置き換えるため、
 * 全列のルールを 1 回で設定して上書き消失を防ぐ。
 */
function applyColumnColorFormats_(sh, columnColorMaps) {
  const rules = [];
  columnColorMaps.forEach(function (entry) {
    const range = sh.getRange(2, entry.col, sh.getMaxRows() - 1, 1);
    Object.keys(entry.colorMap).forEach(function (value) {
      rules.push(
        SpreadsheetApp.newConditionalFormatRule()
          .whenTextEqualTo(value)
          .setBackground(entry.colorMap[value])
          .setRanges([range])
          .build()
      );
    });
  });
  sh.setConditionalFormatRules(rules);
}
