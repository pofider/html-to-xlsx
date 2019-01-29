// eslint-disable-next-line no-unused-vars
function conversion () {
  var tables = document.querySelectorAll('table')
  var tablesOutput = []
  var i

  if (tables.length === 0) {
    return tablesOutput
  }

  for (i = 0; i < tables.length; i++) {
    var table = tables[i]
    var tableOut = { rows: [] }
    var nameAttr = table.getAttribute('name')
    var dataSheetName = table.dataset.sheetName
    var dataIgnoreName = table.dataset.ignoreSheetName

    if (dataIgnoreName == null) {
      if (dataSheetName != null) {
        tableOut.name = dataSheetName
      } else if (nameAttr != null) {
        tableOut.name = nameAttr
      }
    }

    for (var r = 0, n = table.rows.length; r < n; r++) {
      var row = []
      tableOut.rows.push(row)

      for (var c = 0, m = table.rows[r].cells.length; c < m; c++) {
        var cell = table.rows[r].cells[c]
        var cs = document.defaultView.getComputedStyle(cell, null)
        var type = cell.dataset.cellType != null && cell.dataset.cellType !== '' ? cell.dataset.cellType.toLowerCase() : undefined
        var formatStr = cell.dataset.cellFormatStr != null ? cell.dataset.cellFormatStr : undefined
        var formatEnum = cell.dataset.cellFormatEnum != null && !isNaN(parseInt(cell.dataset.cellFormatEnum, 10)) ? parseInt(cell.dataset.cellFormatEnum, 10) : undefined

        row.push({
          type: type,
          // returns the html inside the td element with special html characters like "&" escaped to &amp;
          value: cell.innerHTML,
          // returns just the real text inside the td element with special html characters like "&" left as it is
          valueText: cell.innerText,
          formatStr: formatStr,
          formatEnum: formatEnum,
          backgroundColor: cs.getPropertyValue('background-color').match(/\d+/g),
          foregroundColor: cs.getPropertyValue('color').match(/\d+/g),
          fontFamily: cs.getPropertyValue('font-family').replace(/["']/g, ''),
          fontSize: cs.getPropertyValue('font-size'),
          fontStyle: cs.getPropertyValue('font-style'),
          fontWeight: cs.getPropertyValue('font-weight'),
          textDecoration: {
            line: cs.getPropertyValue('text-decoration').split(' ')[0]
          },
          verticalAlign: cs.getPropertyValue('vertical-align'),
          horizontalAlign: cs.getPropertyValue('text-align'),
          wrapText: cs.getPropertyValue('overflow'),
          width: cell.clientWidth,
          height: cell.clientHeight,
          rowspan: cell.rowSpan,
          colspan: cell.colSpan,
          border: {
            top: cs.getPropertyValue('border-top-style'),
            topColor: cs.getPropertyValue('border-top-color').match(/\d+/g),
            topStyle: cs.getPropertyValue('border-top-style'),
            topWidth: cs.getPropertyValue('border-top-width'),
            right: cs.getPropertyValue('border-right-style'),
            rightColor: cs.getPropertyValue('border-right-color').match(/\d+/g),
            rightStyle: cs.getPropertyValue('border-right-style'),
            rightWidth: cs.getPropertyValue('border-right-width'),
            bottom: cs.getPropertyValue('border-bottom-style'),
            bottomColor: cs.getPropertyValue('border-bottom-color').match(/\d+/g),
            bottomStyle: cs.getPropertyValue('border-bottom-style'),
            bottomWidth: cs.getPropertyValue('border-bottom-width'),
            left: cs.getPropertyValue('border-left-style'),
            leftColor: cs.getPropertyValue('border-left-color').match(/\d+/g),
            leftStyle: cs.getPropertyValue('border-left-style'),
            leftWidth: cs.getPropertyValue('border-left-width')
          }
        })
      }
    }

    tablesOutput.push(tableOut)
  }

  if (tablesOutput.length === 1) {
    return tablesOutput[0]
  }

  return tablesOutput
}
