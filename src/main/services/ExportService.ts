/* oxlint-disable no-case-declarations */
// ExportService

import { loggerService } from '@logger'
import {
  AlignmentType,
  BorderStyle,
  Document,
  ExternalHyperlink,
  HeadingLevel,
  Packer,
  Paragraph,
  ShadingType,
  Table,
  TableCell,
  TableRow,
  TextRun,
  VerticalAlign,
  WidthType
} from 'docx'
import { dialog } from 'electron'
import MarkdownIt from 'markdown-it'

import { fileStorage } from './FileStorage'

const logger = loggerService.withContext('ExportService')
export class ExportService {
  private md: MarkdownIt

  constructor() {
    this.md = new MarkdownIt()
  }

  private convertMarkdownToDocxElements(markdown: string) {
    const tokens = this.md.parse(markdown, {})
    const elements: any[] = []
    let listLevel = 0
    let currentTable: Table | null = null
    let currentRowCells: TableCell[] = []
    let isHeaderRow = false
    let tableColumnCount = 0
    let tableRows: TableRow[] = [] // Store rows temporarily

    const processInlineTokens = (tokens: any[], isHeaderRow: boolean): (TextRun | ExternalHyperlink)[] => {
      const runs: (TextRun | ExternalHyperlink)[] = []
      let linkText = ''
      let linkUrl = ''
      let insideLink = false
      let boldStack = 0 // 跟踪嵌套的粗体标记
      let italicStack = 0 // 跟踪嵌套的斜体标记

      for (let i = 0; i < tokens.length; i++) {
        const token = tokens[i]
        switch (token.type) {
          case 'link_open':
            insideLink = true
            linkUrl = token.attrs.find((attr: [string, string]) => attr[0] === 'href')[1]
            linkText = tokens[i + 1].content
            i += 1
            break
          case 'link_close':
            if (insideLink && linkUrl && linkText) {
              // Handle any accumulated link text with the ExternalHyperlink
              runs.push(
                new ExternalHyperlink({
                  children: [
                    new TextRun({
                      text: linkText,
                      style: 'Hyperlink',
                      color: '0000FF',
                      underline: {
                        type: 'single'
                      }
                    })
                  ],
                  link: linkUrl
                })
              )

              // Reset link variables
              linkText = ''
              linkUrl = ''
              insideLink = false
            }
            break
          case 'strong_open':
            boldStack++
            break
          case 'strong_close':
            boldStack--
            break
          case 'em_open':
            italicStack++
            break
          case 'em_close':
            italicStack--
            break
          case 'text':
            runs.push(
              new TextRun({
                text: token.content,
                bold: isHeaderRow || boldStack > 0,
                italics: italicStack > 0
              })
            )
            break
          case 'code_inline':
            runs.push(
              new TextRun({
                text: token.content,
                font: 'Consolas',
                size: 20,
                bold: isHeaderRow || boldStack > 0,
                italics: italicStack > 0
              })
            )
            break
        }
      }
      return runs
    }

    for (let i = 0; i < tokens.length; i++) {
      const token = tokens[i]
      switch (token.type) {
        case 'heading_open':
          // 获取标题级别 (h1 -> h6)
          const level = parseInt(token.tag.slice(1)) as 1 | 2 | 3 | 4 | 5 | 6
          const headingText = tokens[i + 1].content
          elements.push(
            new Paragraph({
              text: headingText,
              heading: HeadingLevel[`HEADING_${level}`],
              spacing: {
                before: 240,
                after: 120
              }
            })
          )
          i += 2 // 跳过内容标记和闭合标记
          break

        case 'paragraph_open':
          const inlineTokens = tokens[i + 1].children || []
          elements.push(
            new Paragraph({
              children: processInlineTokens(inlineTokens, false),
              spacing: {
                before: 120,
                after: 120
              }
            })
          )
          i += 2
          break

        case 'bullet_list_open':
          listLevel++
          break

        case 'bullet_list_close':
          listLevel--
          break

        case 'list_item_open':
          const itemInlineTokens = tokens[i + 2].children || []
          elements.push(
            new Paragraph({
              children: [
                new TextRun({ text: '•', bold: true }),
                new TextRun({ text: '\t' }),
                ...processInlineTokens(itemInlineTokens, false)
              ],
              indent: {
                left: listLevel * 720
              }
            })
          )
          i += 3
          break

        case 'fence': // 代码块
          const codeLines = token.content.split('\n')
          elements.push(
            new Paragraph({
              children: codeLines.map(
                (line) =>
                  new TextRun({
                    text: line + '\n',
                    font: 'Consolas',
                    size: 20,
                    break: 1
                  })
              ),
              shading: {
                type: ShadingType.SOLID,
                color: 'F5F5F5'
              },
              spacing: {
                before: 120,
                after: 120
              },
              border: {
                top: { style: BorderStyle.SINGLE, size: 1, color: 'DDDDDD' },
                bottom: { style: BorderStyle.SINGLE, size: 1, color: 'DDDDDD' },
                left: { style: BorderStyle.SINGLE, size: 1, color: 'DDDDDD' },
                right: { style: BorderStyle.SINGLE, size: 1, color: 'DDDDDD' }
              }
            })
          )
          break

        case 'hr':
          elements.push(
            new Paragraph({
              children: [new TextRun({ text: '─'.repeat(50), color: '999999' })],
              alignment: AlignmentType.CENTER
            })
          )
          break

        case 'blockquote_open':
          const quoteText = tokens[i + 2].content
          elements.push(
            new Paragraph({
              children: [
                new TextRun({
                  text: quoteText,
                  italics: true
                })
              ],
              indent: {
                left: 720
              },
              border: {
                left: {
                  style: BorderStyle.SINGLE,
                  size: 3,
                  color: 'CCCCCC'
                }
              },
              spacing: {
                before: 120,
                after: 120
              }
            })
          )
          i += 3
          break

        // 表格处理
        case 'table_open':
          tableRows = [] // Reset table rows for new table
          break

        case 'thead_open':
          isHeaderRow = true
          break

        case 'tbody_open':
          isHeaderRow = false
          break

        case 'tr_open':
          currentRowCells = []
          break

        case 'tr_close':
          const row = new TableRow({
            children: currentRowCells,
            tableHeader: isHeaderRow
          })
          tableRows.push(row)
          // 计算表格有多少列（针对第一行）
          if (tableColumnCount === 0) {
            tableColumnCount = currentRowCells.length
          }
          break

        case 'th_open':
        case 'td_open':
          const isFirstColumn = currentRowCells.length === 0 // 判断是否是第一列
          const borders = {
            top: {
              style: BorderStyle.NONE
            },
            bottom: isHeaderRow
              ? {
                  style: BorderStyle.SINGLE,
                  size: 0.5,
                  color: '000000'
                }
              : {
                  style: BorderStyle.NONE
                },
            left: {
              style: BorderStyle.NONE
            },
            right: {
              style: BorderStyle.NONE
            }
          }
          const cellContent = tokens[i + 1]
          const cellOptions = {
            children: [
              new Paragraph({
                children: cellContent.children
                  ? processInlineTokens(cellContent.children, isHeaderRow || isFirstColumn)
                  : [new TextRun({ text: cellContent.content || '', bold: isHeaderRow || isFirstColumn })],
                alignment: AlignmentType.CENTER
              })
            ],
            verticalAlign: VerticalAlign.CENTER,
            borders: borders
          }
          currentRowCells.push(new TableCell(cellOptions))
          i += 2 // 跳过内容和结束标记
          break
        case 'table_close':
          // Create table with the collected rows - avoid using protected properties
          // Create the table with all rows
          currentTable = new Table({
            width: {
              size: 100,
              type: WidthType.PERCENTAGE
            },
            rows: tableRows,
            borders: {
              top: {
                style: BorderStyle.SINGLE,
                size: 1,
                color: '000000'
              },
              bottom: {
                style: BorderStyle.SINGLE,
                size: 1,
                color: '000000'
              },
              left: {
                style: BorderStyle.NONE
              },
              right: {
                style: BorderStyle.NONE
              },
              insideHorizontal: {
                style: BorderStyle.NONE
              },
              insideVertical: {
                style: BorderStyle.NONE
              }
            }
          })
          elements.push(currentTable)
          currentTable = null
          tableColumnCount = 0
          tableRows = []
          currentRowCells = []
          isHeaderRow = false
          break
      }
    }

    return elements
  }

  public exportToWord = async (_: Electron.IpcMainInvokeEvent, markdown: string, fileName: string): Promise<void> => {
    try {
      const elements = this.convertMarkdownToDocxElements(markdown)

      const doc = new Document({
        styles: {
          paragraphStyles: [
            {
              id: 'Normal',
              name: 'Normal',
              run: {
                size: 24,
                font: 'Arial'
              }
            }
          ]
        },
        sections: [
          {
            properties: {},
            children: elements
          }
        ]
      })

      const buffer = await Packer.toBuffer(doc)

      const filePath = dialog.showSaveDialogSync({
        title: '保存文件',
        filters: [{ name: 'Word Document', extensions: ['docx'] }],
        defaultPath: fileName
      })

      if (filePath) {
        await fileStorage.writeFile(_, filePath, buffer)
        logger.debug('Document exported successfully')
      }
    } catch (error) {
      logger.error('Export to Word failed:', error as Error)
      throw error
    }
  }

  private getHtmlTemplate(title: string, content: string): string {
    return `<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>${title}</title>
  <style>
    * {
      margin: 0;
      padding: 0;
      box-sizing: border-box;
    }

    body {
      font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', 'Roboto', 'Helvetica Neue', Arial, sans-serif;
      line-height: 1.6;
      color: #333;
      background-color: #fff;
      padding: 40px 20px;
      max-width: 900px;
      margin: 0 auto;
    }

    h1, h2, h3, h4, h5, h6 {
      margin: 24px 0 16px;
      font-weight: 600;
      line-height: 1.25;
      color: #1a1a1a;
    }

    h1 {
      font-size: 2em;
      border-bottom: 2px solid #eaecef;
      padding-bottom: 0.3em;
    }

    h2 {
      font-size: 1.5em;
      border-bottom: 1px solid #eaecef;
      padding-bottom: 0.3em;
    }

    h3 { font-size: 1.25em; }
    h4 { font-size: 1em; }
    h5 { font-size: 0.875em; }
    h6 { font-size: 0.85em; color: #6a737d; }

    p {
      margin: 16px 0;
    }

    a {
      color: #0366d6;
      text-decoration: none;
    }

    a:hover {
      text-decoration: underline;
    }

    code {
      background-color: rgba(27, 31, 35, 0.05);
      border-radius: 3px;
      font-family: 'SFMono-Regular', Consolas, 'Liberation Mono', Menlo, monospace;
      font-size: 85%;
      padding: 0.2em 0.4em;
    }

    pre {
      background-color: #f6f8fa;
      border-radius: 6px;
      padding: 16px;
      overflow: auto;
      line-height: 1.45;
      margin: 16px 0;
    }

    pre code {
      background-color: transparent;
      padding: 0;
      font-size: 100%;
    }

    blockquote {
      border-left: 4px solid #dfe2e5;
      color: #6a737d;
      padding: 0 15px;
      margin: 16px 0;
    }

    blockquote > :first-child {
      margin-top: 0;
    }

    blockquote > :last-child {
      margin-bottom: 0;
    }

    ul, ol {
      padding-left: 2em;
      margin: 16px 0;
    }

    li {
      margin: 4px 0;
    }

    li > p {
      margin: 0;
    }

    table {
      border-collapse: collapse;
      width: 100%;
      margin: 16px 0;
      overflow: auto;
      display: block;
    }

    table th {
      font-weight: 600;
      background-color: #f6f8fa;
      padding: 6px 13px;
      border: 1px solid #dfe2e5;
    }

    table td {
      padding: 6px 13px;
      border: 1px solid #dfe2e5;
    }

    table tr {
      background-color: #fff;
      border-top: 1px solid #c6cbd1;
    }

    table tr:nth-child(2n) {
      background-color: #f6f8fa;
    }

    hr {
      height: 2px;
      padding: 0;
      margin: 24px 0;
      background-color: #e1e4e8;
      border: 0;
    }

    img {
      max-width: 100%;
      height: auto;
    }

    strong {
      font-weight: 600;
    }

    em {
      font-style: italic;
    }

    del {
      text-decoration: line-through;
    }

    @media print {
      body {
        padding: 0;
      }

      pre {
        page-break-inside: avoid;
      }

      table {
        page-break-inside: avoid;
      }
    }

    @media (prefers-color-scheme: dark) {
      body {
        background-color: #0d1117;
        color: #c9d1d9;
      }

      h1, h2, h3, h4, h5 {
        color: #c9d1d9;
      }

      h1, h2 {
        border-bottom-color: #21262d;
      }

      h6 {
        color: #8b949e;
      }

      a {
        color: #58a6ff;
      }

      code {
        background-color: rgba(110, 118, 129, 0.4);
      }

      pre {
        background-color: #161b22;
      }

      blockquote {
        border-left-color: #3b434b;
        color: #8b949e;
      }

      table th {
        background-color: #161b22;
        border-color: #30363d;
      }

      table td {
        border-color: #30363d;
      }

      table tr {
        background-color: #0d1117;
        border-top-color: #21262d;
      }

      table tr:nth-child(2n) {
        background-color: #161b22;
      }

      hr {
        background-color: #21262d;
      }
    }
  </style>
</head>
<body>
${content}
</body>
</html>`
  }

  public exportToHtml = async (_: Electron.IpcMainInvokeEvent, markdown: string, fileName: string): Promise<void> => {
    try {
      // Convert markdown to HTML
      const htmlContent = this.md.render(markdown)

      // Wrap in template with styling
      const fullHtml = this.getHtmlTemplate(fileName, htmlContent)

      const filePath = dialog.showSaveDialogSync({
        title: '保存文件',
        filters: [{ name: 'HTML Document', extensions: ['html'] }],
        defaultPath: fileName
      })

      if (filePath) {
        await fileStorage.writeFile(_, filePath, Buffer.from(fullHtml, 'utf-8'))
        logger.debug('HTML document exported successfully')
      }
    } catch (error) {
      logger.error('Export to HTML failed:', error as Error)
      throw error
    }
  }
}
