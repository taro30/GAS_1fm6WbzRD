function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('🚀各種オプション')
    .addItem('前日スケジュール追加', 'addCalendarDayEvents')
    .addItem('日付選択追加', 'addCalendarSelectDayEvents')
    .addItem('最終行取得', 'getLastRow')
    .addToUi();
}