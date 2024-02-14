function normalizeA1Notation(a1Notaion: string): string {
  // const p = a1Notaion.split('!')
  // p[p.length - 1] = p[p.length - 1].replace(/[^A-Za-z0-9:]/g, '')
  // return p.join('!')

  // シート名は削除される(対応は大変なので今回はスルーする
  return a1Notaion.replace(/[^A-Za-z0-9:]/g, '')
}
function setValuesOrFormulas(
  sheet: GoogleAppsScript.Spreadsheet.Sheet,
  a1Notaion: string,
  valuesOrFormulas: string[][]
): void {
  console.log(valuesOrFormulas)
  const range = (() => {
    const ret = sheet.getRange(normalizeA1Notation(a1Notaion))
    console.log(
      `rows: ${valuesOrFormulas.length}, cols: ${valuesOrFormulas[0].length}`
    )
    if (
      ret.getNumRows() !== valuesOrFormulas.length ||
      ret.getNumColumns() !== valuesOrFormulas[0].length
    ) {
      return sheet.getRange(
        ret.getRow(),
        ret.getColumn(),
        valuesOrFormulas.length,
        valuesOrFormulas[0].length
      )
    }
    return ret
  })()
  const rangeRows = range.getNumRows()
  const rangeCols = range.getNumColumns()
  console.log(`rangeRows: ${rangeRows}, rangeCols: ${rangeCols}`)
  for (let row = 0; row < rangeRows; row++) {
    for (let col = 0; col < rangeCols; col++) {
      const cell = range.getCell(row + 1, col + 1)
      if (
        typeof valuesOrFormulas[row][col] === 'string' &&
        valuesOrFormulas[row][col].startsWith('=')
      ) {
        cell.setFormula(valuesOrFormulas[row][col])
      } else {
        if (
          valuesOrFormulas[row][col] === undefined ||
          valuesOrFormulas[row][col] === null
        ) {
          cell.setValue('')
        } else if (typeof valuesOrFormulas[row][col] === 'object') {
          cell.setValue(JSON.stringify(valuesOrFormulas[row][col]))
        } else {
          cell.setValue(`${valuesOrFormulas[row][col]}`)
        }
      }
    }
  }
}

function insertRowsAfter(
  sheet: GoogleAppsScript.Spreadsheet.Sheet,
  afterPosition: number,
  howMany: number = 1,
  valuesOrFormulas: string[][] = []
): void {
  const h =
    howMany > valuesOrFormulas.length ? howMany : valuesOrFormulas.length
  sheet.insertRowsAfter(afterPosition, h)
  if (valuesOrFormulas.length > 0) {
    setValuesOrFormulas(
      sheet,
      // range の大きさは一致していなくても、valuesOrFormulas の大きさに合わせてセットされる
      `A${afterPosition + 1}:A${afterPosition + 1 + h}`,
      valuesOrFormulas
    )
  }
}

function setBorder(
  sheet: GoogleAppsScript.Spreadsheet.Sheet,
  a1Notaion: string,
  top?: boolean | null,
  left?: boolean | null,
  bottom?: boolean | null,
  right?: boolean | null,
  vertical?: boolean | null,
  horizontal?: boolean | null,
  color?: string | null,
  style?: string | null
): void {
  const range = sheet.getRange(normalizeA1Notation(a1Notaion))
  const borderStyle = (() => {
    if (style === null) {
      return null
    }
    switch (style) {
      case 'DOTTED':
        return SpreadsheetApp.BorderStyle.DOTTED
      case 'DASHED':
        return SpreadsheetApp.BorderStyle.DASHED
      case 'SOLID':
        return SpreadsheetApp.BorderStyle.SOLID
      case 'SOLID_MEDIUM':
        return SpreadsheetApp.BorderStyle.SOLID_MEDIUM
      case 'SOLID_THICK':
        return SpreadsheetApp.BorderStyle.SOLID_THICK
      default:
    }
    return null
  })()
  const border = range.setBorder(
    top === undefined ? null : top,
    left === undefined ? null : left,
    bottom === undefined ? null : bottom,
    right === undefined ? null : right,
    vertical === undefined ? null : vertical,
    horizontal === undefined ? null : horizontal,
    color === undefined ? null : color,
    borderStyle
  )
}

function setCustomBanding(
  sheet: GoogleAppsScript.Spreadsheet.Sheet,
  a1Notaion: string,
  firstRowColor?: string | null,
  secondRowColor?: string | null,
  headerRowColor?: string | null,
  footerRowColor?: string | null
): void {
  const range = sheet.getRange(normalizeA1Notation(a1Notaion))
  const banding = (() => {
    const currentBanding = range.getBandings()
    if (currentBanding.length > 0) {
      return currentBanding[0]
    }
    return range.applyRowBanding(SpreadsheetApp.BandingTheme.BLUE, false, false)
  })()
  if (firstRowColor) {
    banding.setFirstRowColor(firstRowColor)
  }
  if (secondRowColor) {
    banding.setSecondRowColor(secondRowColor)
  }
  if (headerRowColor) {
    banding.setHeaderRowColor(headerRowColor)
  }
  if (footerRowColor) {
    banding.setFooterRowColor(footerRowColor)
  }
}

type FunctionCallSetValues = {
  name: 'setValuesOrFormulas'
  args: {
    a1Notation: string
    valuesOrFormulas: string[][]
  }
}
type FunctionCallInsertRowsAfter = {
  name: 'insertRowsAfter'
  args: {
    afterPosition: number
    howMany?: number
    valuesOrFormulas?: string[][]
  }
}
type FunctionCallSetBorder = {
  name: 'setBorder'
  args: {
    a1Notation: string
    top?: boolean | null
    left?: boolean | null
    bottom?: boolean | null
    right?: boolean | null
    vertical?: boolean | null
    horizontal?: boolean | null
    color?: string | null
    style?: string | null
  }
}
type FunctionCallSetCustomBanding = {
  name: 'setCustomBanding'
  args: {
    a1Notation: string
    fisetRowColor?: string | null
    secondRowColor?: string | null
    headerRowColor?: string | null
    footerRowColor?: string | null
  }
}
type FunctionCall = {
  functionCall:
    | FunctionCallSetValues
    | FunctionCallInsertRowsAfter
    | FunctionCallSetBorder
    | FunctionCallSetCustomBanding
}
export function apply(
  sheet: GoogleAppsScript.Spreadsheet.Sheet,
  parts: ({ text: string } | FunctionCall)[]
) {
  console.log(JSON.stringify(parts, null, 2))
  for (const part of parts) {
    try {
      if (typeof part === 'object' && 'text' in part) {
        console.log(part.text)
      } else if (typeof part === 'object' && 'functionCall' in part) {
        const { functionCall } = part
        if (functionCall.name === 'setValuesOrFormulas') {
          setValuesOrFormulas(
            sheet,
            functionCall.args.a1Notation,
            functionCall.args.valuesOrFormulas
          )
        } else if (functionCall.name === 'insertRowsAfter') {
          insertRowsAfter(
            sheet,
            functionCall.args.afterPosition,
            functionCall.args.howMany,
            functionCall.args.valuesOrFormulas
          )
        } else if (functionCall.name === 'setBorder') {
          setBorder(
            sheet,
            functionCall.args.a1Notation,
            functionCall.args.top,
            functionCall.args.left,
            functionCall.args.bottom,
            functionCall.args.right,
            functionCall.args.vertical,
            functionCall.args.horizontal,
            functionCall.args.color,
            functionCall.args.style
          )
        } else if (functionCall.name === 'setCustomBanding') {
          setCustomBanding(
            sheet,
            functionCall.args.a1Notation,
            functionCall.args.fisetRowColor,
            functionCall.args.secondRowColor,
            functionCall.args.headerRowColor,
            functionCall.args.footerRowColor
          )
        }
      }
    } catch (e: any) {
      console.error(e)
      throw new Error(
        `Error in part: ${JSON.stringify(part, null, 2)}: ${e.message}`
      )
    }
  }
}
