function onOpen(e: GoogleAppsScript.Events.SheetsOnOpen) {
  SpreadsheetApp.getUi()
    .createMenu('テスト チャット')
    .addItem('チャット開始', 'showSidebar_')
    .addToUi()
}

export { onOpen }
