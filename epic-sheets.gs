/** Epic シートのヘッダー・列幅・初期サンプル行の定義。 */

/**
 * 指定名の Epic 管理シートを展開する。
 * isNew=true のときのみ初期サンプル行を入れる。
 */
function setupEpicSheet(ss, sheetName, isNew) {
  let sh = getOrCreateSheet(ss, sheetName);
  if (isNew) {
    resetSheetCellsForTemplate_(sh, 500, EPIC_COLUMN_COUNT);
    applyEpicTemplateLayout_(sh, true);
    seedEpicSampleRow_(sh);
  } else {
    applyEpicTemplateLayout_(sh, false);
  }
  return sh;
}

/**
 * ヘッダー・見出し書式・折り返しなど「テンプレートの見た目」を適用する（非破壊）。
 * 2 行目以降のデータ値は変更しない。既存シートへの再反映にも使う。
 * applyWidths=true のときだけ列幅を設定する（新規作成時）。再反映では触らない。
 */
function applyEpicTemplateLayout_(sh, applyWidths) {
  if (sh.getMaxColumns() < EPIC_COLUMN_COUNT) {
    sh.insertColumnsAfter(sh.getMaxColumns(), EPIC_COLUMN_COUNT - sh.getMaxColumns());
  }
  sh.getRange(1, 1, 1, EPIC_COLUMN_COUNT).setValues([EPIC_HEADERS]);
  styleHeader(sh, 1, EPIC_COLUMN_COUNT);
  if (applyWidths) setColWidths(sh, EPIC_COLUMN_WIDTHS);
  let lastRow = sh.getLastRow();
  let wrapRows = Math.min(Math.max(lastRow - 1, 100), sh.getMaxRows() - 1);
  if (wrapRows > 0) {
    sh.getRange(2, 1, wrapRows, EPIC_COLUMN_COUNT).setWrap(true);
  }
}

/** 初期サンプル行（ダミー値）。 */
function seedEpicSampleRow_(sh) {
  if (sh.getLastRow() < 2) {
    sh.getRange(2, 1, 1, EPIC_COLUMN_COUNT).setValues([
      ['EPC-001', 'サンプルEpic', '', '',
        'Epic 配下のストーリーは対応するバックログ_ シートの Epic 列にこの Epic名を書く。'],
    ]);
  }
}
