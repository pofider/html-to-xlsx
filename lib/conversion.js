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
        if (val.col === j) {
            return val;
        }
    }

    var workbook = excelbuilder.createWorkbook(options.tmpDir, id + '.xlsx', options);
    var sheet1 = workbook.createSheet('sheet1', getMaxLength(table.rows), table.rows.length);

    var maxWidths = [];
    var curr_row = 0;
    var curr_col = 0;
    var col_offset = 0;
    var row_offset = [];
    for (var i = 0; i < table.rows.length; i++) {
        var maxHeight = 0;
        for (var j = 0; j < table.rows[i].length; j++) {

            //On the start of each row, reset the column counters
            if(j===0){
                curr_col = 0;
                col_offset = 0;
            }
            //Is the current column in the offset list?
            var offset = row_offset.find(searchArray)

            if(offset) {
                //If we have should still be offsetting for a row merge...
                if (curr_row < offset.stop) {
                    curr_col += offset.col_offset || 1;
                }
                //We should not check for an offset as we have passed the row
                //the merge stopped at.  Delete that offset from the array
                else {
                    row_offset = row_offset.filter(function (val) {
                        if (val !== offset) {
                            return true;
                        }
                        else {
                            return false;
                        }
                    });
                }

            }
            var cell = table.rows[i][j];

            if (cell.height > maxHeight) {
                maxHeight = cell.height;
            }

            if (cell.width > (maxWidths[j] || 0)) {
                sheet1.width(curr_col + 1, cell.width / 7);
                maxWidths[j] = cell.width;
            }

            sheet1.set(curr_col + 1, curr_row + 1, cell.value ? cell.value.replace(/&(?!amp;)/g, '&').replace(/&amp;(?!amp;)/g, '&') : cell.value);
            sheet1.align(curr_col + 1, curr_row + 1, cell.horizontalAlign);
            sheet1.valign(curr_col + 1, curr_row + 1, cell.verticalAlign === "middle" ? "center" : cell.verticalAlign);

            if (isColorDefined(cell.backgroundColor)) {
                sheet1.fill(curr_col + 1, curr_row + 1, {
                    type: 'solid',
                    fgColor: 'FF' + rgbToHex(cell.backgroundColor),
                    bgColor: '64'
                });
            }
            sheet1.font(curr_col + 1, curr_row + 1, {
                family: '3',
                scheme: 'minor',
                sz: parseInt(cell.fontSize.replace("px", "")) * 18 / 24,
                bold: cell.fontWeight === "bold" || parseInt(cell.fontWeight, 10) >= 700,
                color: isColorDefined(cell.foregroundColor) ? ('FF' + rgbToHex(cell.foregroundColor)) : undefined
            });

            sheet1.border(curr_col + 1, curr_row + 1, {
                left: getBorderStyle(cell.border.left),
                top: getBorderStyle(cell.border.top),
                right: getBorderStyle(cell.border.right),
                bottom: getBorderStyle(cell.border.bottom)
            });

            //Now that we have done all of the formatting to the cell, see if the row needs merged.
            //Note that calling merge twice on the same cell causes Excel to be unreadable.
            if (cell.rowspan > 1) {
                sheet1.merge({col: curr_col+1, row: curr_row+1},{col: curr_col+cell.colspan, row: curr_row+cell.rowspan});
                row_offset.push({col: j, stop: curr_row+cell.rowspan, col_offset: cell.colspan});

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
            if (cell.colspan > 1 && cell.rowspan === 1) {
                sheet1.merge({col: curr_col+1, row: curr_row+1},{col: curr_col+cell.colspan, row: curr_row+1});
                col_offset += cell.colspan;

                for (var k = curr_col + 1; k <= cell.colspan; k++) {
                    sheet1.border(k + 1, curr_row + 1, {
                        left: getBorderStyle(cell.border.left),
                        top: getBorderStyle(cell.border.top),
                        right: getBorderStyle(cell.border.right),
                        bottom: getBorderStyle(cell.border.bottom)
                    });
                }
            }

            curr_col += cell.colspan
        }

        sheet1.height(curr_row + 1, maxHeight * 18 / 24);

        if (!cell) {
            throw new Error('Cell not found, make sure there are td elements inside tr')
        }

        curr_row += cell.rowspan;
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
