import { apply } from './apply'

type CustomBanding = {
  a1Notation: string
  firstRowColor: string | null
  secondRowColor: string | null
  headerRowColor: string | null
  footerRowColor: string | null
}

function getThumbBase64(imageUrl: string) {
  const fileId = imageUrl.split('/')[5]
  const thumb = DriveApp.getFileById(fileId).getThumbnail()
  return Utilities.base64Encode(thumb.getBytes())
}

function getActiveSheetStat(sheetStat: {
  name: string
  activeCell: string
  activeRanges: string[]
  customBandings: CustomBanding[]
  values: string[][]
}) {
  return `## 知識

以下はアクティブシートの状態

シート名: ${sheetStat.name}

${
  //sheetStat.activeRanges.length == 0
  false ? `アクティブセル: ${sheetStat.activeCell}` : ''
}

${
  sheetStat.activeRanges.length > 0
    ? `アクティブレンジ: ${sheetStat.activeRanges.join(', ')}`
    : ''
}

${
  sheetStat.customBandings.length > 0
    ? `以下はレンジに設定されている banding の一覧です
${'```json\n' + JSON.stringify(sheetStat.customBandings) + '\n```'}
`
    : ''
}

以下はシート内の A! から収められている値と式の二次元配列です。

${'```json\n' + JSON.stringify(sheetStat.values) + '\n```'}

`
}

function buildPrompt(
  imageUrl: string,
  prompt: string,
  sheetStat: {
    name: string
    activeCell: string
    activeRanges: string[]
    customBandings: CustomBanding[]
    values: string[][]
  }
) {
  const parts =
    imageUrl === ''
      ? [
          {
            text: `Goolge SpreadSheets についての要望または質問に対応してください。`
          },
          {
            text: `## 要望または質問

${prompt}}
`
          },
          {
            text: `## 条件

 要望が複数記述されている場合は、プロンプトを部分的なプロンプトに分割してください。
 要望についてはできるだけ関数呼び出すを使ってシートを更新してください。
 シートの更新が難しい場合や質問の場合はテキストで返答してください。 
`
          },
          {
            text: getActiveSheetStat(sheetStat)
          }
        ]
      : [
          {
            inline_data: {
              mime_type: 'image/jpeg',
              data: getThumbBase64(imageUrl)
            }
          },
          {
            text: `Goolge SpreadSheets についての要望または質問に対応してください。`
          },
          {
            text: `## 要望または質問

${prompt}}
`
          },
          {
            text: getActiveSheetStat(sheetStat)
          }
        ]

  return {
    contents: {
      parts
    },
    generationConfig: {
      temperature: 0.9,
      topK: 1,
      topP: 1,
      maxOutputTokens: 2048,
      stopSequences: []
    },
    safetySettings: [
      {
        category: 'HARM_CATEGORY_HARASSMENT',
        threshold: 'BLOCK_MEDIUM_AND_ABOVE'
      },
      {
        category: 'HARM_CATEGORY_HATE_SPEECH',
        threshold: 'BLOCK_MEDIUM_AND_ABOVE'
      },
      {
        category: 'HARM_CATEGORY_SEXUALLY_EXPLICIT',
        threshold: 'BLOCK_MEDIUM_AND_ABOVE'
      },
      {
        category: 'HARM_CATEGORY_DANGEROUS_CONTENT',
        threshold: 'BLOCK_MEDIUM_AND_ABOVE'
      }
    ],
    tools:
      imageUrl === ''
        ? [
            {
              function_declarations: [
                {
                  name: 'runPartialPrompts',
                  description:
                    '要望が複数記述されている場合、プロンプトを複数の部分的なプロンプトとして実行します。',
                  parameters: {
                    type: 'object',
                    properties: {
                      partialPrompts: {
                        type: 'array',
                        items: { type: 'string' },
                        description: `要望を部分的なプロンプトに分割した配列を指定します。
例.1:
オリジナルのプロンプト
\`\`\`
## 要望または質問

表の交互の背景色をグレーにしてください。その後、ヘッダ行の下に罫線をセットしてください。
\`\`\`

分割されたプロンプト
\`\`\`json
[
  "表の交互の背景色をグレーにしてください。",
  "ヘッダ行の下に罫線をセットしてください。"
]
\`\`\`

例.2:
オリジナルのプロンプト
\`\`\`
## 要望または質問

- 表の交互の背景色をグレーにしてください。
- ヘッダ行の下に罫線をセットしてください。
\`\`\`

分割されたプロンプト
\`\`\`json
[
  "表の交互の背景色をグレーにしてください。",
  "ヘッダ行の下に罫線をセットしてください。"
]
\`\`\`
`
                      }
                    },
                    required: ['partialPrompts']
                  }
                },
                {
                  name: 'setValuesOrFormulas',
                  description:
                    'スプレッドシートの指定された範囲へ値または式をセットする。表の作成や更新に使われます。',
                  parameters: {
                    type: 'object',
                    properties: {
                      a1Notation: {
                        type: 'string',
                        description:
                          '値または式の二次元配列をセットする範囲を A1 表記で指定する'
                      },
                      valuesOrFormulas: {
                        type: 'array',
                        items: { type: 'array', items: { type: 'string' } },
                        description: `値または式がセットされている二次元配列(string[][])を指定する。
この二次元配列は Google Apps Script の setValues() または setFormulas() に渡す二次元配列と同じ形式である必要がある。
例: [["品名", "単価", "個数", "売上"], ["みかん", "100", "10" , "=B2*C2"]]`
                      }
                    },
                    required: ['a1Notation', 'valuesOrFormulas']
                  }
                },
                {
                  name: 'insertRowsAfter',
                  description:
                    'スプレッドシートの指定された行位置の後に行を挿入(追加)します。表にデータを挿入(追加)するときに使われます。',
                  parameters: {
                    type: 'object',
                    properties: {
                      afterPosition: {
                        type: 'number',
                        description:
                          'この行の後に新しい行を追加する行。１はじまりの行番号です。'
                      },
                      howMany: {
                        type: 'number',
                        description:
                          '挿入する行数。省略すると 1 行挿入されます。'
                      },
                      valuesOrFormulas: {
                        type: 'array',
                        items: { type: 'array', items: { type: 'string' } },
                        description: `挿入後に値または式をセットする場合に指定。値または式がセットされている二次元配列(string[][])を指定する。
この二次元配列は Google Apps Script の setValues() または setFormulas() に渡す二次元配列と同じ形式である必要がある。
例: [["品名", "単価", "個数", "売上"], ["みかん", "100", "10" , "=B2*C2"]]`
                      }
                    },
                    required: ['afterPosition']
                  }
                },
                {
                  name: 'setBorder',
                  description: 'レンジの枠線を指定する。',
                  parameters: {
                    type: 'object',
                    properties: {
                      a1Notation: {
                        type: 'string',
                        description: 'レンジを A1 表記で指定する'
                      },
                      top: {
                        type: 'boolean',
                        description:
                          'レンジの上辺が枠線の場合は true、なしの場合は false、変更しない場合は null を指定します。'
                      },
                      left: {
                        type: 'boolean',
                        description:
                          'レンジの左辺が枠線の場合は true、なしの場合は false、変更しない場合は null を指定します。'
                      },
                      bottom: {
                        type: 'boolean',
                        description:
                          'レンジの底辺が枠線の場合は true、なしの場合は false、変更しない場合は null を指定します。'
                      },
                      right: {
                        type: 'boolean',
                        description:
                          'レンジの右辺が枠線の場合は true、なしの場合は false、変更しない場合は null を指定します。'
                      },
                      vertical: {
                        type: 'boolean',
                        description:
                          'レンジ内側の縦枠線は true、なしの場合は false、変更しない場合は null です。'
                      },
                      horizontal: {
                        type: 'boolean',
                        description:
                          'レンジ内部の水平枠線は true、なしの場合は false、変更しない場合は null です。'
                      },
                      color: {
                        type: 'string',
                        description:
                          "CSS 表記の色（'#ffffff' や 'white' など）。null はデフォルトの色（黒）を表します。"
                      },
                      style: {
                        type: 'string',
                        description: `枠線のスタイル。デフォルトのスタイル（単色）の場合は null。
| スタイル | 説明 |
| --- | --- |
| DOTTED | 点線の枠線。|
| DASHED | 破線の枠線。|
| SOLID | 細い実線の枠線。|
| SOLID_MEDIUM | 中程度の実線の枠線。|
| SOLID_THICK | 太い実線の枠線。|
| DOUBLE | 2 本の実線の枠線。|
`
                      }
                    },
                    required: ['a1Notation']
                  }
                },
                {
                  name: 'setCustomBanding',
                  description: 'bandingを作成しレンジへセットする',
                  parameters: {
                    type: 'object',
                    properties: {
                      a1Notation: {
                        type: 'string',
                        description: 'レンジを A1 表記で指定する'
                      },
                      fisetRowColor: {
                        type: 'string',
                        description:
                          "交互に表示する 1 行目の色を CSS 表記の色（'#ffffff' や 'white' など）で指定する"
                      },
                      secondRowColor: {
                        type: 'string',
                        description:
                          "交互に表示する 2 行目の色を CSS 表記の色（'#ffffff' や 'white' など）で指定する"
                      },
                      headerRowColor: {
                        type: 'string',
                        description:
                          "ヘッダー行の色を CSS 表記の色（'#ffffff' や 'white' など）で指定する。指定しない場合は null を指定します。"
                      },
                      footerRowColor: {
                        type: 'string',
                        description:
                          "最後の行の色を CSS 表記の色（'#ffffff' や 'white' など）で指定する。指定しない場合は null を指定します。"
                      }
                    },
                    required: ['a1Notation']
                  }
                }
              ]
            }
          ]
        : []
  }
}

export function run(
  imageUrl: string,
  prompt: string,
  activeSheet: GoogleAppsScript.Spreadsheet.Sheet
) {
  const activeCell = activeSheet.getActiveCell().getA1Notation()
  const activeRanges =
    activeSheet
      .getActiveRangeList()
      ?.getRanges()
      .map((range) => {
        return range.getA1Notation()
      }) || []
  if (activeRanges.length === 1 && activeRanges[0] === activeCell) {
    activeRanges.pop()
  }
  const customBandings: CustomBanding[] = activeRanges.flatMap((range) => {
    const bandings: CustomBanding[] = activeSheet
      .getRange(range)
      .getBandings()
      .map((b) => {
        return {
          a1Notation: range,
          firstRowColor: b.getFirstRowColor(),
          secondRowColor: b.getSecondRowColor(),
          headerRowColor: b.getHeaderRowColor(),
          footerRowColor: b.getFooterRowColor()
        }
      })
    return bandings
  })
  const range = activeSheet.getDataRange()
  const sheetValues = range.getValues()
  for (let i = 0; i < sheetValues.length; i++) {
    const row = sheetValues[i]
    for (let j = 0; j < row.length; j++) {
      const formula = range.getCell(i + 1, j + 1).getFormula()
      if (formula) {
        sheetValues[i][j] = formula
      }
    }
  }
  const apiKey =
    PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY')
  const url =
    imageUrl === ''
      ? `https://generativelanguage.googleapis.com/v1beta/models/gemini-pro:generateContent?key=${apiKey}`
      : `https://generativelanguage.googleapis.com/v1beta/models/gemini-pro-vision:generateContent?key=${apiKey}`
  const payload = JSON.stringify(
    buildPrompt(imageUrl, prompt, {
      name: activeSheet.getName(),
      activeCell,
      activeRanges,
      customBandings,
      values: sheetValues
    })
  )
  console.log('--- payload')
  console.log(JSON.stringify(JSON.parse(payload), null, 2))
  const res = UrlFetchApp.fetch(url, {
    method: 'post',
    contentType: 'application/json',
    payload,
    muteHttpExceptions: true
  })
  const ret = res.getContentText()
  const json = JSON.parse(ret)
  console.log(JSON.stringify(json, null, 2))
  apply(activeSheet, json.candidates[0].content.parts)
  //return ret
  return json.candidates[0].content.parts
    .map((part: any) => {
      if (part.text) {
        return part.text
      } else if (part.functionCall) {
        return JSON.stringify(part.functionCall, null, 2)
      }
      return ''
    })
    .join('\n')
}
