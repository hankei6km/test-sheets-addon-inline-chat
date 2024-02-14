import { run } from './gemini'

function showSidebarText_() {
  const ui = HtmlService.createHtmlOutputFromFile('build/sidebar-text')
    .setTitle('My custom sidebar')
    .setWidth(300)
  SpreadsheetApp.getUi().showSidebar(ui)
}

function showSidebarMulti_() {
  const ui = HtmlService.createHtmlOutputFromFile('build/sidebar-multi')
    .setTitle('テストチャット(マルチモーダル)')
    .setWidth(300)
  SpreadsheetApp.getUi().showSidebar(ui)
}

function doPrompt(msg: string): string {
  return run('', msg, SpreadsheetApp.getActiveSheet())
}

function doPromptMulti(imageUrl: string, msg: string): string {
  return run(imageUrl, msg, SpreadsheetApp.getActiveSheet())
}

export { showSidebarText_, showSidebarMulti_, doPrompt, doPromptMulti }
