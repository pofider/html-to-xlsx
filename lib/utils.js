const color = require('tinycolor2')

function sizeToPx (value) {
  if (!value) {
    return 0
  }

  if (typeof value === 'number') {
    return value
  }

  const pt = value.match(/([.\d]+)pt/i)

  if (pt && pt.length === 2) {
    return parseFloat(pt[1], 10) * 96 / 72
  }

  const em = value.match(/([.\d]+)em/i)

  if (em && em.length === 2) {
    return parseFloat(em[1], 10) * 16
  }

  const px = value.match(/([.\d]+)px/i)

  if (px && px.length === 2) {
    return parseFloat(px[1], 10)
  }

  const pe = value.match(/([.\d]+)%/i)

  if (pe && pe.length === 2) {
    return (parseFloat(pe[1], 10) / 100) * 16
  }

  return 0
}

function sizePxToPt (value) {
  const numPx = sizeToPx(value)

  if (numPx > 0) {
    return numPx * 72 / 96
  }

  return 12
}

function sizePtToCm (numPt) {
  return numPt * 0.03528
}

function componentToHex (c) {
  const hex = parseInt(c).toString(16)
  return hex.length === 1 ? '0' + hex : hex
}

function rgbToHex (c) {
  return componentToHex(c[0]) + componentToHex(c[1]) + componentToHex(c[2])
}

function isColorDefined (c) {
  const total = c.length
  let result = false

  for (let i = 0; i < total; i++) {
    result = c[i] !== '0'

    if (result) {
      break
    }
  }

  return result
}

function colorToArgb (c) {
  const input = Array.isArray(c) ? {
    r: c[0],
    g: c[1],
    b: c[2],
    a: c[3]
  } : c

  const rgba = color(input).toHex8()
  return rgba.substr(6) + rgba.substr(0, 6)
}

function getBorderStyle (border) {
  if (border === 'none') return undefined

  if (border === 'solid') return 'thin'

  if (border === 'double') return 'double'

  return undefined
}

function getBorder (cellInfo, type) {
  let color = cellInfo.border[`${type}Color`]
  let style = cellInfo.border[`${type}Style`]
  let width = cellInfo.border[`${type}Width`]

  if (!color) {
    return null
  }

  width = sizeToPx(width)

  if (width <= 0) {
    return null
  }

  color = colorToArgb(color)

  if (style === 'dashed' || style === 'dotted' || style === 'double') {
    return { style, color }
  }

  style = 'thin'

  if (width >= 3 && width < 5) {
    style = 'medium'
  }

  if (width >= 5) {
    style = 'thick'
  }

  return { style, color }
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

function assetLegalXMLChar (str) {
  const validChars = /[\u0000-\u0008\u000B-\u000C\u000E-\u001F\uD800-\uDFFF\uFFFE-\uFFFF]/

  const result = str.match(validChars)

  if (result) {
    throw new Error(`Invalid character "${result}"${result[0] != null ? `, char code: ${result[0].charCodeAt(0)}` : ''} in string "${str}" at index: ${result.index}`)
  }

  return str
}

module.exports.assetLegalXMLChar = assetLegalXMLChar
module.exports.sizePxToPt = sizePxToPt
module.exports.sizePtToCm = sizePtToCm
module.exports.componentToHex = componentToHex
module.exports.rgbToHex = rgbToHex
module.exports.isColorDefined = isColorDefined
module.exports.colorToArgb = colorToArgb
module.exports.getBorderStyle = getBorderStyle
module.exports.getBorder = getBorder
module.exports.getTotalRows = getTotalRows
module.exports.getTotalCells = getTotalCells
