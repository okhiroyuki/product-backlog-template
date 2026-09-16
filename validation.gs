/** ドロップダウン（入力規則）の適用。Epic・ステータス・着手可能性・ポイント。 */

/** 登録済みの全バックログ／Epic シートのデータ行すべてに入力規則を適用する。 */
function applyAllReferenceValidations_(ss) {
  let names = getBacklogSheetNames_(ss);
  for (let i = 0; i < names.length; i++) {
    let sh = ss.getSheetByName(names[i]);
    if (!sh) continue;
    let epicNameRange = getEpicNameRange_(ss, epicSheetNameForBacklog_(ss, names[i]));
    let lr = sh.getLastRow();
    for (let r = 2; r <= lr; r++) {
      applyRowValidations_(sh, r, epicNameRange);
    }
  }
}

/** バックログ 1 行分の入力規則を適用する。 */
function applyRowValidations_(sh, row, epicNameRange) {
  if (epicNameRange) {
    const epicRule = SpreadsheetApp.newDataValidation()
      .requireValueInRange(epicNameRange, true)
      .setAllowInvalid(false)
      .build();
    sh.getRange(row, BACKLOG_COLUMNS.EPIC).setDataValidation(epicRule);
  }
  setDropdown(sh, row, BACKLOG_COLUMNS.STATUS, STATUS_OPTIONS);
  setDropdown(sh, row, BACKLOG_COLUMNS.DOABLE, DOABLE_OPTIONS);
  setDropdown(sh, row, BACKLOG_COLUMNS.POINT, POINT_OPTIONS);
}

/**
 * 指定のエピックシートの Epic名列を、プルダウンの参照範囲として返す。
 * シートの最終行ではなく最大行まで参照するため、Epic追加後の再反映は不要。
 */
function getEpicNameRange_(ss, epicSheetName) {
  if (!epicSheetName) return null;
  let sh = ss.getSheetByName(epicSheetName);
  if (!sh || sh.getMaxRows() < 2) return null;
  return sh.getRange(2, EPIC_COLUMNS.NAME, sh.getMaxRows() - 1, 1);
}

/**
 * メニュー「🔁 テンプレートを既存シートへ再反映」用。
 * 既存データ（2 行目以降の値）は保持したまま、最新のテンプレートを入れ直す。
 * 反映対象: ヘッダー名・見出し書式・折り返し・入力規則・条件付き書式（列幅は変更しない）。
 * 開いているシートが未登録のバックログ相当なら 🔢 ID管理 に登録して対象に含める。
 * バックログ_ があるのに対応 エピック_ が無い場合は エピック_ を新規作成する。
 */
function refreshBacklogTemplates() {
  let ss = SpreadsheetApp.getActiveSpreadsheet();

  let active = ss.getActiveSheet();
  if (active) {
    let activeName = active.getName();
    let isBacklogLike =
      activeName.indexOf(BACKLOG_SHEET_PREFIX) === 0 ||
      activeName.indexOf(LEGACY_BACKLOG_SHEET_PREFIX) === 0 ||
      looksLikeBacklogSheet_(active);
    if (activeName !== ID_SHEET_NAME && !isEpicSheetName_(activeName) && isBacklogLike) {
      registerSheetIfNeeded_(ss, activeName);
    }
  }

  let names = getBacklogSheetNames_(ss);
  let epicSheetNames = getEpicSheetNames_(ss);
  if (names.length === 0 && epicSheetNames.length === 0) {
    notifyUser_(
      '🔢 ID管理 にバックログシートが 1 件も登録されていないため、適用対象がありません。\n' +
        '反映したいシートを開いた状態で、もう一度このメニューを実行してください。',
      '再反映'
    );
    return;
  }

  for (let i = 0; i < names.length; i++) {
    let sh = ss.getSheetByName(names[i]);
    if (sh) applyBacklogTemplateLayout_(sh, false);
    let epicName = derivedEpicSheetName_(names[i]);
    if (!epicName) continue;
    let epicExisting = ss.getSheetByName(epicName);
    setupEpicSheet(ss, epicName, !epicExisting);
  }
  epicSheetNames = getEpicSheetNames_(ss);
  for (let i = 0; i < epicSheetNames.length; i++) {
    let epicSh = ss.getSheetByName(epicSheetNames[i]);
    if (epicSh) applyEpicTemplateLayout_(epicSh, false);
  }

  SpreadsheetApp.flush();
  applyAllReferenceValidations_(ss);
  SpreadsheetApp.flush();
  applyReferenceColorFormats_(ss);
  SpreadsheetApp.flush();
  syncIdCountersFromBookCore(ss);
  SpreadsheetApp.flush();
  reorderBacklogTabs_(ss);

  let targets = names.slice();
  for (let i = 0; i < epicSheetNames.length; i++) {
    if (targets.indexOf(epicSheetNames[i]) < 0) targets.push(epicSheetNames[i]);
  }
  notifyUser_(
    'テンプレートを再反映しました（データは保持）。\n\n対象シート: ' + targets.join('、'),
    '再反映'
  );
}

/** 全バックログシートの条件付き書式（背景色マップ）を適用する。 */
function applyReferenceColorFormats_(ss) {
  let backlogMaps = [
    { col: BACKLOG_COLUMNS.DOABLE, colorMap: DOABLE_COLORS },
    { col: BACKLOG_COLUMNS.STATUS, colorMap: STATUS_COLORS },
  ];
  let names = getBacklogSheetNames_(ss);
  for (let i = 0; i < names.length; i++) {
    let sh = ss.getSheetByName(names[i]);
    if (!sh) continue;
    applyColumnColorFormats_(sh, backlogMaps);
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
