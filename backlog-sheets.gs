/** バックログシートのヘッダー・列幅・初期サンプル行の定義。 */

/** 指定名のバックログシートを展開する。isNew=true のときのみ初期サンプル行を入れる。 */
function setupBacklogSheet(ss, sheetName, isNew) {
  let sh = getOrCreateSheet(ss, sheetName);
  if (isNew) {
    resetSheetCellsForTemplate_(sh, 500, BACKLOG_COLUMN_COUNT);
    applyBacklogTemplateLayout_(sh, true);
    seedBacklogSampleRow_(sh);
  } else {
    applyBacklogTemplateLayout_(sh, false);
  }
  return sh;
}

/**
 * ヘッダー・見出し書式・折り返しなど「テンプレートの見た目」を適用する（非破壊）。
 * 2 行目以降のデータ値は変更しない。既存シートへの再反映にも使う。
 * applyWidths=true のときだけ列幅を設定する（新規作成時）。再反映では触らない。
 */
function applyBacklogTemplateLayout_(sh, applyWidths) {
  migrateBacklogAcColumnIfNeeded_(sh);
  if (sh.getMaxColumns() < BACKLOG_COLUMN_COUNT) {
    sh.insertColumnsAfter(sh.getMaxColumns(), BACKLOG_COLUMN_COUNT - sh.getMaxColumns());
  }
  sh.getRange(1, 1, 1, BACKLOG_COLUMN_COUNT).setValues([BACKLOG_HEADERS]);
  styleHeader(sh, 1, BACKLOG_COLUMN_COUNT);
  if (applyWidths) setColWidths(sh, BACKLOG_COLUMN_WIDTHS);
  let lastRow = sh.getLastRow();
  let wrapRows = Math.min(Math.max(lastRow - 1, 100), sh.getMaxRows() - 1);
  if (wrapRows > 0) {
    sh.getRange(2, 1, wrapRows, BACKLOG_COLUMN_COUNT).setWrap(true);
  }
}

/**
 * 旧テンプレ（価値の次がポイント）から、受け入れ条件（AC）列を挿入する。
 * 再反映時に既存データの列ずれを防ぐ。すでに AC 列があるシートは何もしない。
 */
function migrateBacklogAcColumnIfNeeded_(sh) {
  let lastCol = Math.max(sh.getLastColumn(), 1);
  let headers = sh.getRange(1, 1, 1, lastCol).getValues()[0];
  for (let i = 0; i < headers.length; i++) {
    if (String(headers[i]).trim() === '受け入れ条件（AC）') return;
  }
  if (String(headers[7] || '').trim() !== 'ポイント') return;
  sh.insertColumnAfter(7);
}

/** 初期サンプル行（ダミー値）。本線の項目と、それを支える技術検証の 2 行。 */
function seedBacklogSampleRow_(sh) {
  if (sh.getLastRow() < 2) {
    sh.getRange(2, 1, 2, BACKLOG_COLUMN_COUNT).setValues([
      ['PBL-001', 'サンプルEpic', 'Open', 'Not Ready', 'サンプル担当',
        'サンプルの機能を実装したい。',
        'テンプレートの動作を確認したいから。',
        'テンプレートの列とプルダウンが確認できる。',
        '3', '', '', '詳細はPRDを参照。'],
      ['PBL-002', 'サンプルEpic', 'Open', 'Ready', 'サンプル担当',
        'サンプルの機能に必要な技術検証をしたい。',
        'PBL-001 の実現方針と見積もりを固めたいから。',
        '実現方針と見積もりの根拠がドキュメントに残っている。',
        '2', '', '', '備考に補足を書いてよい。'],
    ]);
  }
}
