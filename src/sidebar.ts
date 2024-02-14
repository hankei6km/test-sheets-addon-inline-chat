import { run } from './gemini'

function showSidebar_() {
  const ui = HtmlService.createHtmlOutputFromFile('build/sidebar')
    .setTitle('My custom sidebar')
    .setWidth(300)
  SpreadsheetApp.getUi().showSidebar(ui)
}

function doPrompt(msg: string): string {
  return run(msg, SpreadsheetApp.getActiveSheet())
}

export { showSidebar_, doPrompt }
