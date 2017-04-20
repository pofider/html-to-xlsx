var path = require("path"),
    fs = require("fs"),
    uuid = require("uuid").v1,
    tmpDir = require("os").tmpdir(),
    excelbuilder = require('msexcel-builder-extended');


function componentToHex(c) {
    var hex = parseInt(c).toString(16);
    return hex.length === 1 ? "0" + hex : hex;
}

function rgbToHex(c) {
    return componentToHex(c[0]) + componentToHex(c[1]) + componentToHex(c[2]);
}

function isColorDefined(c) {
    return c[0] !== "0" || c[1] !== "0" || c[2] !== "0" || c[3] !== "0";
}

function getMaxLength(array) {
    var max = 0;
    array.forEach(function (a) {
        if (a.length > max)
            max = a.length;
    });
    return max;
}

function getBorderStyle(border) {
    if (border === "none")
        return undefined;

    if (border === "solid")
        return "thin";

    if (border === "double")
        return "double";

    return undefined;
}


function convert(html, cb) {
    var id = uuid();

    function icb(err, table) {
        if (err) {
            return cb(err);
        }

        try {
            tableToXlsx(table, id, cb);
        }
        catch(e) {
            e.message = JSON.stringify(e.message);
            cb(e);
        }
    }

    if (options.strategy === "phantom-server")
        return require("./serverStrategy.js")(options, html, id, icb);
    if (options.strategy === "dedicated-process")
        return require("./dedicatedProcessStrategy.js")(options, html, id, icb);

    cb(new Error("Unsupported strategy " + options.strategy));
}

function tableToXlsx(table, id, cb) {

    function searchArray(val) {
        if (val.col <= tmpCol) {
            return true;
        }
        return false;
    }

    var workbook = excelbuilder.createWorkbook(options.tmpDir, id + '.xlsx', options);
    var sheet1 = workbook.createSheet('sheet1', getMaxLength(table.rows), table.rows.length);

    var maxWidths = [];
    var curr_row = 0;
    var row_offset = [];

    for (var i = 0; i < table.rows.length; i++) {
      //set at the row level column offsets and column position
      var maxHeight = 0;
      var tmp_offsets = [];
      var curr_col = 0;
      var col_offset = 0;
      var tmpCol = 0;

      //clean out offsets that are no longer valid
      row_offset = row_offset.filter(function(offset) {
        return (offset.stop <= curr_row) ? false : true;
      });

      //Is the current column in the offset list?
      row_offset.map(function(item) {
        //if so add an offset to shift the column start
        if (curr_row <= item.stop) {
          col_offset += item.col_offset || 1;
        }
      });

        //column iterator
        for (var j = 0; j < table.rows[i].length; j++) {
            var cell = table.rows[i][j];

            //On the start of each row, reset the column counters
            //Use tmpCol to manipulate the column value
            curr_col = (j===0) ? 1 : curr_col + 1;
            tmpCol = curr_col + col_offset;

            if (cell.height > maxHeight) {
                maxHeight = cell.height;
            }

            if (cell.width > (maxWidths[j] || 0)) {
                sheet1.width(curr_col, cell.width / 7);
                maxWidths[j] = cell.width;
            }

            sheet1.set(tmpCol, curr_row + 1, cell.value ? cell.value.replace(/&(?!amp;)/g, '&').replace(/&amp;(?!amp;)/g, '&') : cell.value);
            sheet1.align(tmpCol, curr_row + 1, cell.horizontalAlign);
            sheet1.valign(tmpCol, curr_row + 1, cell.verticalAlign === "middle" ? "center" : cell.verticalAlign);

            if (isColorDefined(cell.backgroundColor)) {
                sheet1.fill(tmpCol, curr_row + 1, {
                    type: 'solid',
                    fgColor: 'FF' + rgbToHex(cell.backgroundColor),
                    bgColor: '64'
                });
            }
            sheet1.font(tmpCol, curr_row + 1, {
                family: '3',
                scheme: 'minor',
                sz: parseInt(cell.fontSize.replace("px", "")) * 18 / 24,
                bold: cell.fontWeight === "bold" || parseInt(cell.fontWeight, 10) >= 700,
                color: isColorDefined(cell.foregroundColor) ? ('FF' + rgbToHex(cell.foregroundColor)) : undefined
            });

            sheet1.border(tmpCol, curr_row + 1, {
                left: getBorderStyle(cell.border.left),
                top: getBorderStyle(cell.border.top),
                right: getBorderStyle(cell.border.right),
                bottom: getBorderStyle(cell.border.bottom)
            });

            //Now that we have done all of the formatting to the cell, see if the row needs merged.
            //Note that calling merge twice on the same cell causes Excel to be unreadable.
            if (cell.rowspan > 1) {
                //address colspan at the same time as rowspan
                var coloffset = (cell.colspan > 1) ? cell.colspan - 1 : 0
                sheet1.merge({col: tmpCol, row: curr_row+1},{col: tmpCol + coloffset, row: curr_row+cell.rowspan});
                //store the rowspan for later use to shift over the column starting point
                tmp_offsets.push({col: tmpCol, stop: curr_row+cell.rowspan, col_offset: cell.colspan});
                curr_col += cell.colspan - 1

                for (var k = curr_row + 1; k <= cell.rowspan; k++) {
                    sheet1.border(k + 1, curr_row + cell.colspan, {
                        left: getBorderStyle(cell.border.left),
                        top: getBorderStyle(cell.border.top),
                        right: getBorderStyle(cell.border.right),
                        bottom: getBorderStyle(cell.border.bottom)
                    });
                }
            }

            //If we already did rowspan, we did the colspan at the same time so this only does colspan.
            //No need to store the colspan as that doesn't carry over to another row
            if (cell.colspan > 1 && cell.rowspan === 1) {
              var coloffset = (cell.colspan > 1) ? cell.colspan - 1 : 0
                sheet1.merge({col: tmpCol, row: curr_row+1},{col: tmpCol+coloffset, row: curr_row+1});
                curr_col += cell.colspan - 1
                for (var k = tmpCol; k <= cell.colspan; k++) {
                    sheet1.border(k + 1, curr_row + 1, {
                        left: getBorderStyle(cell.border.left),
                        top: getBorderStyle(cell.border.top),
                        right: getBorderStyle(cell.border.right),
                        bottom: getBorderStyle(cell.border.bottom)
                    });
                }
            }
        }

        sheet1.height(curr_row + 1, maxHeight * 18 / 24);

        if (!cell) {
               throw new Error('Cell not found, make sure there are td elements inside tr')
        }
        curr_row += cell.rowspan;
        row_offset = row_offset.concat(tmp_offsets);
    }

    workbook.save(function (err) {
        if (err) {
            return cb(err);
        }

        cb(null, fs.createReadStream(path.join(options.tmpDir, id + ".xlsx")));
    });
}

module.exports = function (opt) {
    options = opt || {};
    options.timeout = options.timeout || 10000;
    options.tmpDir = options.tmpDir || tmpDir;
    options.strategy = options.strategy || "phantom-server";

    // always set env var names for phantom-workers (don't let the user override this config)
    options.hostEnvVarName = 'PHANTOM_WORKER_HOST';
    options.portEnvVarName = 'PHANTOM_WORKER_PORT';

    convert.options = options;
    return convert;
};
