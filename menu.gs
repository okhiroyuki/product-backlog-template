/** カスタムメニュー（onOpen）。 */

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('バックログ')
    .addItem('🆕 バックログシートを追加（名前を指定）', 'createBacklogSheet')
    .addItem('🔁 テンプレートを既存シートへ再反映', 'refreshBacklogTemplates')
    .addItem('🔢 IDカウンタをブックから再同期', 'syncIdCountersFromBook')
    .addToUi();
}