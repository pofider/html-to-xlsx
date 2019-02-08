const util = require('util')
const path = require('path')
const fs = require('fs')
const moment = require('moment')
const XlsxPopulate = require('xlsx-populate')
const stylesMap = require('./stylesMap')
const utils = require('./utils')
const writeFileAsync = util.promisify(fs.writeFile)

async function tableToXlsx (options, tables, id) {
  const tablesToProcess = Array.isArray(tables) ? tables : [tables]
  const workbook = await XlsxPopulate.fromBlankAsync()

  tablesToProcess.forEach((table, tableIdx) => {
    const sheetName = table.name != null ? table.name : `Sheet${tableIdx + 1}`
    let sheet

    if (tableIdx === 0) {
      // the excel is created with default sheet, we rename it
      sheet = workbook.sheet('Sheet1').name(sheetName)
    } else {
      sheet = workbook.addSheet(sheetName)
    }

    const maxWidths = []
    const currentCellOffsetsPerRow = []
    let currentRowInFile = 0

    // rows processing
    for (let rIdx = 0; rIdx < table.rows.length; rIdx++) {
      let maxHeight = 0 // default height

      if (currentCellOffsetsPerRow[currentRowInFile] === undefined) {
        currentCellOffsetsPerRow[currentRowInFile] = 0
      }

      const allCellsAreRowSpan = table.rows[rIdx].filter(c => c.rowspan > 1).length === table.rows[rIdx].length

      if (table.rows[rIdx].length === 0) {
        throw new Error('Cell not found, make sure there are td elements inside tr')
      }

      // cells processing
      for (let cIdx = 0; cIdx < table.rows[rIdx].length; cIdx++) {
        const cellInfo = table.rows[rIdx][cIdx]

        utils.assetLegalXMLChar(cellInfo.valueText)

        // when all cells are rowspan in a row then the row itself doesn't count
        let rowSpan = cellInfo.rowspan - (allCellsAreRowSpan ? 1 : 0)

        // condition for rowspan don't merge more rows than rows available in table
        if (((currentRowInFile + 1) + (rowSpan - 1)) > table.rows.length) {
          rowSpan = (table.rows.length - (currentRowInFile + 1)) + 1
        }

        rowSpan = Math.max(rowSpan, 1)

        const cellSpan = cellInfo.colspan

        // row height & col width
        if (cellInfo.height) {
          const pt = utils.sizePxToPt(cellInfo.height)

          if (pt > maxHeight) {
            maxHeight = pt / rowSpan
          }
        }

        if (cellInfo.width) {
          if (!maxWidths[cIdx]) {
            maxWidths[cIdx] = 0 // default width
          }

          const pt = cellInfo.width / 7

          if (pt > maxWidths[cIdx]) {
            maxWidths[cIdx] = pt / cellSpan
          }
        }

        const cell = sheet.cell(currentRowInFile + 1, currentCellOffsetsPerRow[currentRowInFile] + 1)

        if (cellInfo.type === 'number') {
          cell.value(parseFloat(cellInfo.valueText))
        } else if (cellInfo.type === 'bool' || cellInfo.type === 'boolean') {
          cell.value(cellInfo.valueText === 'true' || cellInfo.valueText === '1')
        } else if (cellInfo.type === 'date') {
          // dates in excel are just numbers with some format applied to make
          // the number appear as date
          // https://github.com/dtjohnson/xlsx-populate#dates
          cell.value(moment(cellInfo.valueText).toDate()).style('numberFormat', 'yyyy-mm-dd')
        } else if (cellInfo.type === 'datetime') {
          cell.value(moment(cellInfo.valueText).toDate()).style('numberFormat', 'yyyy-mm-dd h:mm:ss')
        } else if (cellInfo.type === 'formula') {
          cell.formula(cellInfo.valueText)
        } else {
          cell.value(cellInfo.valueText)
        }

        if (cellInfo.formatStr != null) {
          cell.style('numberFormat', cellInfo.formatStr)
        } else if (cellInfo.formatEnum != null && utils.numFmtMap[cellInfo.formatEnum] != null) {
          cell.style('numberFormat', utils.numFmtMap[cellInfo.formatEnum])
        }

        const styles = getXlsxStyles(cellInfo)

        if (rowSpan > 1 || cellSpan > 1) {
          const rowIncrement = Math.max(rowSpan - 1, 0)
          const cellIncrement = Math.max(cellSpan - 1, 0)

          // row number is returned as 1-based
          const startRow = cell.rowNumber()
          const endRow = startRow + rowIncrement
          // column number is returned as 1-base
          const startCell = cell.columnNumber()
          const endCell = startCell + cellIncrement

          // range takes numbers as 1-based
          const range = sheet.range(
            startRow,
            startCell,
            endRow,
            endCell
          ).merged(true)

          if (Object.keys(styles).length > 0) {
            range.style(styles)
          }
        } else {
          if (Object.keys(styles).length > 0) {
            cell.style(styles)
          }
        }

        for (let r = 0; r < rowSpan; r++) {
          if (currentCellOffsetsPerRow[currentRowInFile + r] == null) {
            currentCellOffsetsPerRow[currentRowInFile + r] = 0
          }

          currentCellOffsetsPerRow[currentRowInFile + r] += cellSpan
        }
      }

      // set row height according to the max height of cells in current row
      sheet.row(currentRowInFile + 1).height(maxHeight)

      if (!allCellsAreRowSpan) {
        currentRowInFile++
      }
    }

    // set col width according to the max width of all cells in table
    for (let i = 0; i < maxWidths.length; i++) {
      const width = maxWidths[i]

      if (width) {
        sheet.column(i + 1).width(width)
      }
    }
  })

  const outputFilePath = path.join(options.tmpDir, `${id}.xlsx`)

  const xlsxBuf = await workbook.outputAsync()

  await writeFileAsync(outputFilePath, xlsxBuf)

  return fs.createReadStream(outputFilePath)
}

function getXlsxStyles (cellInfo) {
  const styles = {}

  Object.entries(stylesMap).forEach(([styleName, getStyle]) => {
    const result = getStyle(cellInfo)

    if (result !== undefined) {
      styles[styleName] = result
    }
  })

  return styles
}

module.exports = tableToXlsx
