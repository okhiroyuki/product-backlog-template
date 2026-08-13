/** バックログシートのヘッダー・列幅・初期サンプル行の定義。 */

/** 指定名のバックログシートを展開する。isNew=true のときのみ初期サンプル行を入れる。 */
function setupBacklogSheet(ss, sheetName, isNew) {
  let sh = getOrCreateSheet(ss, sheetName);
  if (isNew) {
    resetSheetCellsForTemplate_(sh, 500, BACKLOG_COLUMN_COUNT);
    sh.getRange(1, 1, 1, BACKLOG_COLUMN_COUNT).setValues([
      ['ID', 'テーマ', 'ステータス', '着手可能性', '誰が', '何をしたい', 'それはなぜか（価値）', 'ポイント', 'PRD', 'JIRA', '備考'],
    ]);
    styleHeader(sh, 1, BACKLOG_COLUMN_COUNT);
    setColWidths(sh, [90, 140, 100, 100, 120, 340, 340, 70, 90, 110, 200]);
    sh.getRange(2, 1, 100, BACKLOG_COLUMN_COUNT).setWrap(true);
    seedBacklogSampleRow_(sh);
  }
  if (!isNew) {
    sh.getRange(1, BACKLOG_COLUMNS.BUKO).setValue('備考');
    sh.setColumnWidth(BACKLOG_COLUMNS.BUKO, 200);
  }
  return sh;
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
