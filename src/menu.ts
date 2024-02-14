function onOpen(e: GoogleAppsScript.Events.SheetsOnOpen) {
  SpreadsheetApp.getUi()
    .createMenu('テスト チャット')
    .addItem('チャット開始', 'showSidebarText_')
    .addToUi()
}

export { onOpen }
