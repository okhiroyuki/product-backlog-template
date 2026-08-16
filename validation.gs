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

/** 1 行分の入力規則（種別・ステータス・着手可能性・ポイント）を適用する。 */
function applyRowValidations_(sh, row) {
  setDropdown(sh, row, BACKLOG_COLUMNS.TYPE, TYPE_OPTIONS);
  setDropdown(sh, row, BACKLOG_COLUMNS.STATUS, STATUS_OPTIONS);
  setDropdown(sh, row, BACKLOG_COLUMNS.DOABLE, DOABLE_OPTIONS);
  setDropdown(sh, row, BACKLOG_COLUMNS.POINT, POINT_OPTIONS);
}

/**
 * メニュー「🔁 テンプレートを既存シートへ再反映」用。
 * 既存データ（2 行目以降の値）は保持したまま、最新のテンプレートを入れ直す。
 * 反映対象: ヘッダー名・見出し書式・列幅・折り返し・入力規則・条件付き書式。
 * 開いているシートが未登録のバックログ相当なら 🔢 ID管理 に登録して対象に含める。
 */
function refreshBacklogTemplates() {
  let ss = SpreadsheetApp.getActiveSpreadsheet();

  let active = ss.getActiveSheet();
  if (active && active.getName() !== ID_SHEET_NAME) {
    registerSheetIfNeeded_(ss, active.getName());
  }

  let names = getBacklogSheetNames_(ss);
  if (names.length === 0) {
    notifyUser_(
      '🔢 ID管理 にバックログシートが 1 件も登録されていないため、適用対象がありません。\n' +
        '反映したいシートを開いた状態で、もう一度このメニューを実行してください。',
      '再反映'
    );
    return;
  }

  for (let i = 0; i < names.length; i++) {
    let sh = ss.getSheetByName(names[i]);
    if (sh) applyBacklogTemplateLayout_(sh);
  }
  SpreadsheetApp.flush();
  applyAllReferenceValidations_(ss);
  SpreadsheetApp.flush();
  applyReferenceColorFormats_(ss);
  SpreadsheetApp.flush();
  syncIdCountersFromBookCore(ss);
  SpreadsheetApp.flush();

  notifyUser_(
    'テンプレートを再反映しました（データは保持）。\n\n対象シート: ' + names.join('、'),
    '再反映'
  );
}

/** 全バックログシートの種別・着手可能性・ステータス列に条件付き書式（背景色マップ）を適用する。 */
function applyReferenceColorFormats_(ss) {
  let columnColorMaps = [
    { col: BACKLOG_COLUMNS.TYPE, colorMap: TYPE_COLORS },
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
