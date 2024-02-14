import { run } from './gemini'

function showSidebarText_() {
  const ui = HtmlService.createHtmlOutputFromFile('build/sidebar-text')
    .setTitle('My custom sidebar')
    .setWidth(300)
  SpreadsheetApp.getUi().showSidebar(ui)
}

function doPrompt(msg: string): string {
  return run(msg, SpreadsheetApp.getActiveSheet())
}

export { showSidebarText_, doPrompt }
