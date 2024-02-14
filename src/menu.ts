function onOpen(e: GoogleAppsScript.Events.SheetsOnOpen) {
  SpreadsheetApp.getUi()
    .createMenu('テスト チャット')
    .addItem('チャット開始', 'showSidebarText_')
    .addItem('チャット開始(マルチモーダル)', 'showSidebarMulti_')
    .addToUi()
}

export { onOpen }
