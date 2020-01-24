const path = require('path')
const fs = require('fs')
const moment = require('moment')
const pEachSeries = require('p-each-series')
const ExcelJS = require('exceljs')
const stylesMap = require('./stylesMap')
const utils = require('./utils')

async function tableToXlsx (options, tables, xlsxTemplateBuf, id) {
  const outputFilePath = path.join(options.tmpDir, `${id}.xlsx`)

  const workbook = new ExcelJS.stream.xlsx.WorkbookWriter({
    template: xlsxTemplateBuf,
    filename: outputFilePath,
    useStyles: true,
    useSharedStrings: false
  })

  if (xlsxTemplateBuf) {
    await workbook.waitForTemplateParse()
  }

  const tablesToProcess = tables

  await pEachSeries(tablesToProcess, async (table) => {
    const sheet = await workbook.addWorksheetAsync(table.name)

    const context = {
      currentRowInFile: 0,
      currentCellOffsetsPerRow: [],
      pendingCellOffsetsPerRow: [],
      usedCells: [],
      maxWidths: [],
      totalRows: table.rowsCount
    }

    await table.getRows((row) => {
      addRow(sheet, row, context)
    })

    sheet.commit()
  })

  await workbook.commit()

  return fs.createReadStream(outputFilePath)
}

function addRow (sheet, row, context) {
  const currentCellOffsetsPerRow = context.currentCellOffsetsPerRow
  const pendingCellOffsetsPerRow = context.pendingCellOffsetsPerRow
  const usedCells = context.usedCells
  const maxWidths = context.maxWidths
  const totalRows = context.totalRows
  let maxHeight = 0 // default height

  if (currentCellOffsetsPerRow[context.currentRowInFile] === undefined) {
    currentCellOffsetsPerRow[context.currentRowInFile] = 0
  }

  if (
    pendingCellOffsetsPerRow[context.currentRowInFile] === undefined ||
    pendingCellOffsetsPerRow[context.currentRowInFile].length === 0
  ) {
    pendingCellOffsetsPerRow[context.currentRowInFile] = [{
      pending: 0
    }]
  }

  const allCellsAreRowSpan = row.filter(c => c.rowspan > 1).length === row.length

  if (row.length === 0) {
    throw new Error('Cell not found, make sure there are td elements inside tr')
  }

  for (let cIdx = 0; cIdx < row.length; cIdx++) {
    const cellInfo = row[cIdx]

    utils.assetLegalXMLChar(cellInfo.valueText)

    // when all cells are rowspan in a row then the row itself doesn't count
    let rowSpan = cellInfo.rowspan - (allCellsAreRowSpan ? 1 : 0)

    // condition for rowspan don't merge more rows than rows available in table
    if (((context.currentRowInFile + 1) + (rowSpan - 1)) > totalRows) {
      rowSpan = (totalRows - (context.currentRowInFile + 1)) + 1
    }

    rowSpan = Math.max(rowSpan, 1)

    const cellSpan = cellInfo.colspan

    // row height
    if (cellInfo.height) {
      const pt = utils.sizePxToPt(cellInfo.height)

      if (pt > maxHeight) {
        maxHeight = pt / rowSpan
      }
    }

    // col width
    if (cellInfo.width) {
      if (!maxWidths[cIdx]) {
        maxWidths[cIdx] = 0 // default width
      }

      const pt = cellInfo.width / 7

      if (pt > maxWidths[cIdx]) {
        const width = pt / cellSpan
        maxWidths[cIdx] = width
        // we need to set column width before row commit in order
        // to make it work
        sheet.getColumn(cIdx + 1).width = width
      }
    }

    const cell = sheet.getCell(`${context.currentRowInFile + 1}`, `${currentCellOffsetsPerRow[context.currentRowInFile] + 1}`)
    let startCell = cell.col
    let endCell = cell.col

    usedCells[`${cell.row},${cell.col}`] = true

    if (cellInfo.type === 'number') {
      cell.value = parseFloat(cellInfo.valueText)
    } else if (cellInfo.type === 'bool' || cellInfo.type === 'boolean') {
      cell.value = cellInfo.valueText === 'true' || cellInfo.valueText === '1'
    } else if (cellInfo.type === 'date') {
      cell.value = moment(cellInfo.valueText).toDate()
      cell.numFmt = 'yyyy-mm-dd'
    } else if (cellInfo.type === 'datetime') {
      cell.value = moment(cellInfo.valueText).toDate()
      cell.numFmt = 'yyyy-mm-dd h:mm:ss'
    } else if (cellInfo.type === 'formula') {
      cell.value = {
        formula: cellInfo.valueText
      }
    } else {
      cell.value = cellInfo.valueText
    }

    const styles = getXlsxStyles(cellInfo)

    if (rowSpan > 1 || cellSpan > 1) {
      const rowIncrement = Math.max(rowSpan - 1, 0)
      const cellIncrement = Math.max(cellSpan - 1, 0)

      // row number is returned as 1-based
      const startRow = cell.row
      const endRow = startRow + rowIncrement
      // column number is returned as 1-base
      startCell = cell.col
      endCell = startCell + cellIncrement

      for (let r = startRow; r <= endRow; r++) {
        for (let c = startCell; c <= endCell; c++) {
          if (usedCells[`${r},${c}`] == null) {
            usedCells[`${r},${c}`] = true
          }
        }
      }

      sheet.mergeCells(startRow, startCell, endRow, endCell)

      // merged cells share the same style object so setting the style
      // in one cell will do it also for the other cells
      if (Object.keys(styles).length > 0) {
        setStyles(cell, styles)
      }
    } else {
      if (Object.keys(styles).length > 0) {
        setStyles(cell, styles)
      }
    }

    for (let r = 0; r < rowSpan; r++) {
      if (currentCellOffsetsPerRow[context.currentRowInFile + r] == null) {
        currentCellOffsetsPerRow[context.currentRowInFile + r] = 0
      }

      if (
        pendingCellOffsetsPerRow[context.currentRowInFile + r] == null ||
        pendingCellOffsetsPerRow[context.currentRowInFile + r].length === 0
      ) {
        pendingCellOffsetsPerRow[context.currentRowInFile + r] = [{
          pending: 0
        }]
      }

      // don't increase offset when previous cell was not set, instead reserve it for later.
      // this makes some rowspan/colspan layout to work properly
      if (usedCells[`${context.currentRowInFile + r + 1},${Math.max(startCell - 1, 1)}`] != null) {
        currentCellOffsetsPerRow[context.currentRowInFile + r] += cellSpan

        const currentPending = pendingCellOffsetsPerRow[context.currentRowInFile + r][0]

        if (
          currentPending &&
          currentPending.pending !== 0 &&
          Math.max(endCell + 1) >= currentPending.lastCellStart
        ) {
          currentCellOffsetsPerRow[context.currentRowInFile + r] += currentPending.pending
          pendingCellOffsetsPerRow[context.currentRowInFile + r].shift()
        }
      } else {
        const lastPending = pendingCellOffsetsPerRow[context.currentRowInFile + r][pendingCellOffsetsPerRow[context.currentRowInFile + r].length - 1]

        if (lastPending && lastPending.lastCellStart != null && lastPending.lastCellStart !== startCell) {
          pendingCellOffsetsPerRow[context.currentRowInFile + r].push({
            lastCellStart: startCell,
            pending: cellSpan
          })
        } else if (lastPending) {
          lastPending.lastCellStart = startCell
          lastPending.pending += cellSpan
        }
      }
    }
  }

  // set row height according to the max height of cells in current row
  sheet.getRow(context.currentRowInFile + 1).height = maxHeight

  if (!allCellsAreRowSpan) {
    sheet.getRow(context.currentRowInFile + 1).commit()
    context.currentRowInFile++
  }
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

function setStyles (cell, styles) {
  for (let [styleName, styleValue] of Object.entries(styles)) {
    cell[styleName] = styleValue
  }
}

module.exports = tableToXlsx
