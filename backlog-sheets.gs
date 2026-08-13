/** バックログシートのヘッダー・列幅・初期サンプル行の定義。 */

/** 指定名のバックログシートを展開する。isNew=true のときのみ初期サンプル行を入れる。 */
function setupBacklogSheet(ss, sheetName, isNew) {
  let sh = getOrCreateSheet(ss, sheetName);
  if (isNew) {
    resetSheetCellsForTemplate_(sh, 500, BACKLOG_COLUMN_COUNT);
    applyBacklogTemplateLayout_(sh);
    seedBacklogSampleRow_(sh);
  } else {
    applyBacklogTemplateLayout_(sh);
  }
  return sh;
}

/**
 * ヘッダー・見出し書式・列幅・折り返しなど「テンプレートの見た目」を適用する（非破壊）。
 * 2 行目以降のデータ値は変更しない。既存シートへの再反映にも使う。
 */
function applyBacklogTemplateLayout_(sh) {
  sh.getRange(1, 1, 1, BACKLOG_COLUMN_COUNT).setValues([BACKLOG_HEADERS]);
  styleHeader(sh, 1, BACKLOG_COLUMN_COUNT);
  setColWidths(sh, BACKLOG_COLUMN_WIDTHS);
  let lastRow = sh.getLastRow();
  let wrapRows = Math.min(Math.max(lastRow - 1, 100), sh.getMaxRows() - 1);
  if (wrapRows > 0) {
    sh.getRange(2, 1, wrapRows, BACKLOG_COLUMN_COUNT).setWrap(true);
  }
}

/** 初期サンプル行（ダミー値・1 行だけ）。 */
function seedBacklogSampleRow_(sh) {
  if (sh.getLastRow() < 2) {
    sh.getRange(2, 1, 1, BACKLOG_COLUMN_COUNT).setValues([
      ['PBL-001', 'サンプルテーマ', 'Open', 'Not Ready', 'サンプル担当',
        'サンプルの機能を実装したい。',
        'テンプレートの動作を確認したいから。',
        '3', '', '', '詳細はPRDを参照。'],
    ]);
  }
}
