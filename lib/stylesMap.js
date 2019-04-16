const utils = require('./utils')

module.exports = {
  horizontalAlignment: (cellInfo) => {
    const hMap = {
      left: 'left',
      right: 'right',
      center: 'center',
      justify: 'justify'
    }

    if (cellInfo.horizontalAlign && hMap[cellInfo.horizontalAlign]) {
      return hMap[cellInfo.horizontalAlign]
    }
  },
  verticalAlignment: (cellInfo) => {
    const vMap = {
      top: 'top',
      bottom: 'bottom',
      middle: 'center'
    }

    if (cellInfo.verticalAlign && vMap[cellInfo.verticalAlign]) {
      return vMap[cellInfo.verticalAlign]
    }
  },
  wrapText: (cellInfo) => cellInfo.wrapText === 'scroll' || cellInfo.wrapText === 'auto',
  fill: (cellInfo) => {
    if (utils.isColorDefined(cellInfo.backgroundColor)) {
      return {
        type: 'solid',
        color: utils.colorToArgb(cellInfo.backgroundColor)
      }
    }
  },
  fontColor: (cellInfo) => {
    if (utils.isColorDefined(cellInfo.foregroundColor)) {
      return utils.colorToArgb(cellInfo.foregroundColor)
    }
  },
  fontSize: (cellInfo) => utils.sizePxToPt(cellInfo.fontSize),
  fontFamily: (cellInfo) => cellInfo.fontFamily != null ? cellInfo.fontFamily : `Calibri`,
  bold: (cellInfo) => cellInfo.fontWeight === 'bold' || parseInt(cellInfo.fontWeight, 10) >= 700,
  italic: (cellInfo) => cellInfo.fontStyle === 'italic',
  underline: (cellInfo) => {
    if (cellInfo.textDecoration) {
      return cellInfo.textDecoration.line === 'underline'
    }
  },
  strikethrough: (cellInfo) => {
    if (cellInfo.textDecoration) {
      return cellInfo.textDecoration.line === 'line-through'
    }
  },
  leftBorderStyle: (cellInfo) => {
    const left = utils.getBorder(cellInfo, 'left')

    if (left && left.style != null) {
      return left.style
    }
  },
  leftBorderColor: (cellInfo) => {
    const left = utils.getBorder(cellInfo, 'left')

    if (left && left.color != null) {
      return left.color
    }
  },
  rightBorderStyle: (cellInfo) => {
    const right = utils.getBorder(cellInfo, 'right')

    if (right && right.style != null) {
      return right.style
    }
  },
  rightBorderColor: (cellInfo) => {
    const right = utils.getBorder(cellInfo, 'right')

    if (right && right.color != null) {
      return right.color
    }
  },
  topBorderStyle: (cellInfo) => {
    const top = utils.getBorder(cellInfo, 'top')

    if (top && top.style != null) {
      return top.style
    }
  },
  topBorderColor: (cellInfo) => {
    const top = utils.getBorder(cellInfo, 'top')

    if (top && top.color != null) {
      return top.color
    }
  },
  bottomBorderStyle: (cellInfo) => {
    const bottom = utils.getBorder(cellInfo, 'bottom')

    if (bottom && bottom.style != null) {
      return bottom.style
    }
  },
  bottomBorderColor: (cellInfo) => {
    const bottom = utils.getBorder(cellInfo, 'bottom')

    if (bottom && bottom.color != null) {
      return bottom.color
    }
  }
}
