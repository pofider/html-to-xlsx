const util = require('util')
const path = require('path')
const fs = require('fs')
const moment = require('moment')
const XlsxPopulate = require('xlsx-populate')
const excelbuilder = require('msexcel-builder-extended')
const utils = require('./utils')
const writeFileAsync = util.promisify(fs.writeFile)

async function oldTableToXlsx (options, table, id, convertOptions) {
  const workbook = excelbuilder.createWorkbook(
    options.tmpDir,
    `${id}.xlsx`,
    options
  )

  const totalCells = utils.getTotalCells(table.rows)
  const totalRows = utils.getTotalRows(table.rows)

  const sheet = workbook.createSheet(
    'Sheet1',
    totalCells,
    totalRows
  )

  const maxWidths = []
  let currRow = 0
  let rowsToMerge = 0
  let rowOffset = []

  for (let i = 0; i < table.rows.length; i++) {
    // set at the row level column offsets and column position
    let maxHeight = 0
    let maxRowSpan = 0
    const tmpOffsets = []
    let currCol = 0
    let colOffset = 0
    let tmpCol = 0
    let mergedRow = false
    const cellsCount = table.rows[i].length
    const allCellsAreRowSpan = table.rows[i].filter(c => c.rowspan > 1).length === cellsCount

    if (rowsToMerge > 0) {
      mergedRow = true
      currRow--
      rowsToMerge--
    }

    // clean out offsets that are no longer valid
    rowOffset = rowOffset.filter((offset) => {
      // eslint-disable-next-line no-unneeded-ternary
      return offset.stop <= currRow ? false : true
    })

    // Is the current column in the offset list?
    rowOffset.forEach((item) => {
      // if so add an offset to shift the column start
      if (currRow <= item.stop) {
        colOffset += item.colOffset || 1
      }
    })

    let cellInfo

    // column iterator
    for (let j = 0; j < cellsCount; j++) {
      cellInfo = table.rows[i][j]

      // On the start of each row, reset the column counters
      // Use tmpCol to manipulate the column value
      currCol = j === 0 ? 1 : currCol + 1
      tmpCol = currCol + colOffset

      // don't take into account height of merged cells (rowspan) for max height.
      // this is the behavior of browsers too when rendering tables.
      // see issue #20 for more info
      if (cellInfo.height > maxHeight && (cellInfo.rowspan == null || cellInfo.rowspan <= 1)) {
        maxHeight = cellInfo.height
      }

      // don't take into account width of merged cells for max width,
      // and don't apply width for merged cells.
      // this is the behavior of browsers too when rendering tables.
      // see issue #19, #24 for more info
      if (cellInfo.colspan == null || cellInfo.colspan <= 1) {
        if (cellInfo.width > (maxWidths[j] || 0)) {
          sheet.width(currCol, cellInfo.width / 7)
          maxWidths[j] = cellInfo.width
        }
      }

      sheet.set(
        tmpCol,
        currRow + 1,
        cellInfo.value
          ? cellInfo.value.replace(/&(?!amp;)/g, '&').replace(/&amp;(?!amp;)/g, '&')
          : cellInfo.value
      )

      sheet.align(tmpCol, currRow + 1, cellInfo.horizontalAlign)

      sheet.valign(
        tmpCol,
        currRow + 1,
        cellInfo.verticalAlign === 'middle' ? 'center' : cellInfo.verticalAlign
      )

      sheet.wrap(tmpCol, currRow + 1, cellInfo.wrapText === 'scroll')

      if (utils.isColorDefined(cellInfo.backgroundColor)) {
        sheet.fill(tmpCol, currRow + 1, {
          type: 'solid',
          fgColor: 'FF' + utils.rgbToHex(cellInfo.backgroundColor),
          bgColor: '64'
        })
      }

      sheet.font(tmpCol, currRow + 1, {
        family: '3',
        scheme: 'minor',
        sz: parseInt(cellInfo.fontSize.replace('px', '')) * 18 / 24,
        bold:
          cellInfo.fontWeight === 'bold' || parseInt(cellInfo.fontWeight, 10) >= 700,
        color: utils.isColorDefined(cellInfo.foregroundColor)
          ? 'FF' + utils.rgbToHex(cellInfo.foregroundColor)
          : undefined
      })

      sheet.border(tmpCol, currRow + 1, {
        left: utils.getBorderStyle(cellInfo.border.left),
        top: utils.getBorderStyle(cellInfo.border.top),
        right: utils.getBorderStyle(cellInfo.border.right),
        bottom: utils.getBorderStyle(cellInfo.border.bottom)
      })

      // Now that we have done all of the formatting to the cell, see if the row needs merged.
      // Note that calling merge twice on the same cell causes Excel to be unreadable.
      if (cellInfo.rowspan > 1) {
        // address colspan at the same time as rowspan
        let coloffset = cellInfo.colspan > 1 ? cellInfo.colspan - 1 : 0
        let endRow = (currRow + cellInfo.rowspan) - ((cellsCount === 1 || allCellsAreRowSpan) ? 1 : 0)

        sheet.merge(
          { col: tmpCol, row: currRow + 1 },
          { col: tmpCol + coloffset, row: endRow }
        )

        // store the rowspan for later use to shift over the column starting point
        tmpOffsets.push({
          col: tmpCol,
          stop: endRow,
          colOffset: cellInfo.colspan
        })

        currCol += cellInfo.colspan - 1

        for (let k = currRow + 1; k <= cellInfo.rowspan; k++) {
          sheet.border(k + 1, currRow + cellInfo.colspan, {
            left: utils.getBorderStyle(cellInfo.border.left),
            top: utils.getBorderStyle(cellInfo.border.top),
            right: utils.getBorderStyle(cellInfo.border.right),
            bottom: utils.getBorderStyle(cellInfo.border.bottom)
          })
        }

        if (cellInfo.rowspan > maxRowSpan) {
          maxRowSpan = cellInfo.rowspan
        }
      }

      // If we already did rowspan, we did the colspan at the same time so this only does colspan.
      // No need to store the colspan as that doesn't carry over to another row
      if (cellInfo.colspan > 1 && cellInfo.rowspan === 1) {
        let coloffset = cellInfo.colspan > 1 ? cellInfo.colspan - 1 : 0

        sheet.merge(
          { col: tmpCol, row: currRow + 1 },
          { col: tmpCol + coloffset, row: currRow + 1 }
        )

        currCol += cellInfo.colspan - 1

        for (let k = tmpCol; k <= cellInfo.colspan; k++) {
          sheet.border(k + 1, currRow + 1, {
            left: utils.getBorderStyle(cellInfo.border.left),
            top: utils.getBorderStyle(cellInfo.border.top),
            right: utils.getBorderStyle(cellInfo.border.right),
            bottom: utils.getBorderStyle(cellInfo.border.bottom)
          })
        }
      }
    }

    sheet.height(currRow + 1, maxHeight * 18 / 24)

    if (!cellInfo) {
      throw new Error(
        'Cell not found, make sure there are td elements inside tr'
      )
    }

    if (mergedRow && rowsToMerge > 0) {
      // if row was merged one we restore its index
      currRow += 1
    }

    currRow += 1

    if (maxRowSpan > 1) {
      rowsToMerge = maxRowSpan - 1

      if (cellsCount > 1 && !allCellsAreRowSpan) {
        currRow++
      }
    }

    rowOffset = rowOffset.concat(tmpOffsets)
  }

  return new Promise((resolve, reject) => {
    workbook.save((err) => {
      if (err) {
        return reject(err)
      }

      resolve(
        fs.createReadStream(path.join(options.tmpDir, `${id}.xlsx`))
      )
    })
  })
}

async function tableToXlsx (options, tables, id, convertOptions) {
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

        const styles = getXlsxStyles(cellInfo, convertOptions)

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
          sheet.range(
            startRow,
            startCell,
            endRow,
            endCell
          ).merged(true).style(styles)
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

function getXlsxStyles (cellInfo, convertOptions) {
  const styles = {}

  // horizontal align
  const hMap = {
    left: 'left',
    right: 'right',
    center: 'center',
    justify: 'justify'
  }

  if (cellInfo.horizontalAlign && hMap[cellInfo.horizontalAlign]) {
    styles.horizontalAlignment = hMap[cellInfo.horizontalAlign]
  }

  // vertical align
  const vMap = {
    top: 'top',
    bottom: 'bottom',
    middle: 'center'
  }

  if (cellInfo.verticalAlign && vMap[cellInfo.verticalAlign]) {
    styles.verticalAlignment = vMap[cellInfo.verticalAlign]
  }

  if (cellInfo.wrapText === 'scroll') {
    styles.wrapText = true
  }

  if (utils.isColorDefined(cellInfo.backgroundColor)) {
    styles.fill = {
      type: 'solid',
      color: utils.colorToArgb(cellInfo.backgroundColor)
    }
  }

  if (utils.isColorDefined(cellInfo.foregroundColor)) {
    styles.fontColor = utils.colorToArgb(cellInfo.foregroundColor)
  }

  styles.fontSize = utils.sizePxToPt(cellInfo.fontSize)
  styles.fontFamily = convertOptions.fontFamily != null ? convertOptions.fontFamily : `Calibri`
  styles.bold = cellInfo.fontWeight === 'bold' || parseInt(cellInfo.fontWeight, 10) >= 700
  styles.italic = cellInfo.fontStyle === 'italic'

  if (cellInfo.textDecoration) {
    styles.underline = cellInfo.textDecoration.line === 'underline'
    styles.strikethrough = cellInfo.textDecoration.line === 'line-through'
  }

  const left = utils.getBorder(cellInfo, 'left')

  if (left) {
    if (left.style != null) {
      styles.leftBorderStyle = left.style
    }

    styles.leftBorderColor = left.color
  }

  const right = utils.getBorder(cellInfo, 'right')

  if (right) {
    if (right.style != null) {
      styles.rightBorderStyle = right.style
    }

    styles.rightBorderColor = right.color
  }

  const top = utils.getBorder(cellInfo, 'top')

  if (top) {
    if (top.style != null) {
      styles.topBorderStyle = top.style
    }

    styles.topBorderColor = top.color
  }

  const bottom = utils.getBorder(cellInfo, 'bottom')

  if (bottom) {
    if (bottom.style != null) {
      styles.bottomBorderStyle = bottom.style
    }

    styles.bottomBorderColor = bottom.color
  }

  return styles
}

module.exports = tableToXlsx
module.exports.legacy = oldTableToXlsx
