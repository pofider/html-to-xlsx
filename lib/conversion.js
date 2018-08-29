'use strict'

const util = require('util')
const path = require('path')
const fs = require('fs')
const uuid = require('uuid/v4')
const tmpDir = require('os').tmpdir()
const excelbuilder = require('msexcel-builder-extended')

const readFileAsync = util.promisify(fs.readFile)
const writeFileAsync = util.promisify(fs.writeFile)

function componentToHex (c) {
  const hex = parseInt(c).toString(16)
  return hex.length === 1 ? '0' + hex : hex
}

function rgbToHex (c) {
  return componentToHex(c[0]) + componentToHex(c[1]) + componentToHex(c[2])
}

function isColorDefined (c) {
  return c[0] !== '0' || c[1] !== '0' || c[2] !== '0' || c[3] !== '0'
}

function getTotalCells (rows) {
  let max = 0

  rows.forEach((row, idx) => {
    let cellsCount = row.length
    let maxRowSpan = 0
    const allCellsAreRowSpan = row.filter(c => c.rowspan > 1).length === row.length

    maxRowSpan = row.reduce((acu, cell) => {
      if (cell.rowspan && cell.rowspan > acu) {
        acu = cell.rowspan
      }

      return acu
    }, 0)

    row.forEach((cell) => {
      if (cell.colspan && cell.colspan > 1) {
        cellsCount += (cell.colspan) - 1
      }
    })

    if (maxRowSpan > 1 && (row.length === 1 || allCellsAreRowSpan)) {
      const rowsToMerge = maxRowSpan - 1
      let mergedMaxCells = 0

      for (let i = 0; i < rowsToMerge; i++) {
        const nextRowIndex = idx + (i + 1)

        if (rows[nextRowIndex] == null) {
          continue
        }

        if (rows[nextRowIndex].length > mergedMaxCells) {
          mergedMaxCells = rows[nextRowIndex].length
        }
      }

      cellsCount += mergedMaxCells
    }

    if (cellsCount > max) {
      max = cellsCount
    }
  })

  return max
}

function getTotalRows (rows) {
  let rowCount = 0
  let rowsToMerge = 0

  rows.forEach((row) => {
    let maxRowSpan = 0
    const allCellsAreRowSpan = row.filter(c => c.rowspan > 1).length === row.length

    maxRowSpan = row.reduce((acu, cell) => {
      if (cell.rowspan && cell.rowspan > acu) {
        acu = cell.rowspan
      }

      return acu
    }, 0)

    if (rowsToMerge > 0) {
      rowsToMerge--
    } else {
      if (maxRowSpan > 1) {
        rowsToMerge += maxRowSpan - 1
        rowCount += rowsToMerge

        if (row.length > 1 && !allCellsAreRowSpan) {
          rowCount++
        }
      } else {
        rowCount++
      }
    }
  })

  return rowCount
}

function getBorderStyle (border) {
  if (border === 'none') return undefined

  if (border === 'solid') return 'thin'

  if (border === 'double') return 'double'

  return undefined
}

function tableToXlsx (options, table, id, cb) {
  const workbook = excelbuilder.createWorkbook(
    options.tmpDir,
    `${id}.xlsx`,
    options
  )

  const totalCells = getTotalCells(table.rows)
  const totalRows = getTotalRows(table.rows)

  const sheet1 = workbook.createSheet(
    'sheet1',
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

    let cell

    // column iterator
    for (let j = 0; j < table.rows[i].length; j++) {
      cell = table.rows[i][j]

      // On the start of each row, reset the column counters
      // Use tmpCol to manipulate the column value
      currCol = j === 0 ? 1 : currCol + 1
      tmpCol = currCol + colOffset

      // don't take into account height of merged cells (rowspan) for max height.
      // this is the behavior of browsers too when rendering tables.
      // see issue #20 for more info
      if (cell.height > maxHeight && (cell.rowspan == null || cell.rowspan <= 1)) {
        maxHeight = cell.height
      }

      // don't take into account width of merged cells for max width,
      // and don't apply width for merged cells.
      // this is the behavior of browsers too when rendering tables.
      // see issue #19, #24 for more info
      if (cell.colspan == null || cell.colspan <= 1) {
        if (cell.width > (maxWidths[j] || 0)) {
          sheet1.width(currCol, cell.width / 7)
          maxWidths[j] = cell.width
        }
      }

      sheet1.set(
        tmpCol,
        currRow + 1,
        cell.value
          ? cell.value.replace(/&(?!amp;)/g, '&').replace(/&amp;(?!amp;)/g, '&')
          : cell.value
      )

      sheet1.align(tmpCol, currRow + 1, cell.horizontalAlign)

      sheet1.valign(
        tmpCol,
        currRow + 1,
        cell.verticalAlign === 'middle' ? 'center' : cell.verticalAlign
      )

      sheet1.wrap(tmpCol, currRow + 1, cell.wrapText === 'scroll')

      if (isColorDefined(cell.backgroundColor)) {
        sheet1.fill(tmpCol, currRow + 1, {
          type: 'solid',
          fgColor: 'FF' + rgbToHex(cell.backgroundColor),
          bgColor: '64'
        })
      }

      sheet1.font(tmpCol, currRow + 1, {
        family: '3',
        scheme: 'minor',
        sz: parseInt(cell.fontSize.replace('px', '')) * 18 / 24,
        bold:
          cell.fontWeight === 'bold' || parseInt(cell.fontWeight, 10) >= 700,
        color: isColorDefined(cell.foregroundColor)
          ? 'FF' + rgbToHex(cell.foregroundColor)
          : undefined
      })

      sheet1.border(tmpCol, currRow + 1, {
        left: getBorderStyle(cell.border.left),
        top: getBorderStyle(cell.border.top),
        right: getBorderStyle(cell.border.right),
        bottom: getBorderStyle(cell.border.bottom)
      })

      // Now that we have done all of the formatting to the cell, see if the row needs merged.
      // Note that calling merge twice on the same cell causes Excel to be unreadable.
      if (cell.rowspan > 1) {
        // address colspan at the same time as rowspan
        let coloffset = cell.colspan > 1 ? cell.colspan - 1 : 0
        let endRow = (currRow + cell.rowspan) - ((cellsCount === 1 || allCellsAreRowSpan) ? 1 : 0)

        sheet1.merge(
          { col: tmpCol, row: currRow + 1 },
          { col: tmpCol + coloffset, row: endRow }
        )

        // store the rowspan for later use to shift over the column starting point
        tmpOffsets.push({
          col: tmpCol,
          stop: endRow,
          colOffset: cell.colspan
        })

        currCol += cell.colspan - 1

        for (let k = currRow + 1; k <= cell.rowspan; k++) {
          sheet1.border(k + 1, currRow + cell.colspan, {
            left: getBorderStyle(cell.border.left),
            top: getBorderStyle(cell.border.top),
            right: getBorderStyle(cell.border.right),
            bottom: getBorderStyle(cell.border.bottom)
          })
        }

        if (cell.rowspan > maxRowSpan) {
          maxRowSpan = cell.rowspan
        }
      }

      // If we already did rowspan, we did the colspan at the same time so this only does colspan.
      // No need to store the colspan as that doesn't carry over to another row
      if (cell.colspan > 1 && cell.rowspan === 1) {
        let coloffset = cell.colspan > 1 ? cell.colspan - 1 : 0

        sheet1.merge(
          { col: tmpCol, row: currRow + 1 },
          { col: tmpCol + coloffset, row: currRow + 1 }
        )

        currCol += cell.colspan - 1

        for (let k = tmpCol; k <= cell.colspan; k++) {
          sheet1.border(k + 1, currRow + 1, {
            left: getBorderStyle(cell.border.left),
            top: getBorderStyle(cell.border.top),
            right: getBorderStyle(cell.border.right),
            bottom: getBorderStyle(cell.border.bottom)
          })
        }
      }
    }

    sheet1.height(currRow + 1, maxHeight * 18 / 24)

    if (!cell) {
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

module.exports = (opt = {}) => {
  const options = { ...opt }

  options.timeout = options.timeout || 10000
  options.tmpDir = options.tmpDir || tmpDir

  if (typeof options.extract !== 'function') {
    throw new Error('`extract` option must be a function')
  }

  if (typeof options.timeout !== 'number') {
    throw new Error('`timeout` option must be a number')
  }

  const timeout = options.timeout
  const currentExtractFn = options.extract

  async function convert (html, convertOptions = {}) {
    const id = uuid()

    if (html == null) {
      throw new Error('required `html` option not specified')
    }

    const htmlPath = path.join(options.tmpDir, `${id}-html-to-xlsx.html`)
    const scriptFnPath = options.conversionScriptPath != null ? options.conversionScriptPath : path.join(__dirname, 'scripts/conversionScript.js')

    let scriptFn = await readFileAsync(scriptFnPath)

    scriptFn = scriptFn.toString()

    await writeFileAsync(htmlPath, html)

    const table = await currentExtractFn({
      ...convertOptions,
      html: htmlPath,
      scriptFn,
      timeout
    })

    const stream = await tableToXlsx(options, table, id)

    return stream
  }

  return convert
}
