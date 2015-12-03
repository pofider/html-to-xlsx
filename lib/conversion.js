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
        if (err)
            return cb(err);

        tableToXlsx(table, id, cb);
    }

    if (options.strategy === "phantom-server")
        return require("./serverStrategy.js")(options, html, id, icb);
    if (options.strategy === "dedicated-process")
        return require("./dedicatedProcessStrategy.js")(options, html, id, icb);

    cb(new Error("Unsupported strategy " + options.strategy));
}

function tableToXlsx(table, id, cb) {

    var workbook = excelbuilder.createWorkbook(options.tmpDir, id + ".xlsx");
    var sheet1 = workbook.createSheet('sheet1', getMaxLength(table.rows), table.rows.length);

    var maxWidths = [];
    for (var i = 0; i < table.rows.length; i++) {
        var maxHeight = 0;
        for (var j = 0; j < table.rows[i].length; j++) {
            var cell = table.rows[i][j];

            if (cell.height > maxHeight) {
                maxHeight = cell.height;
            }

            if (cell.width > (maxWidths[j] || 0)) {
                sheet1.width(j + 1, cell.width / 7);
                maxWidths[j] = cell.width;
            }

            sheet1.set(j + 1, i + 1, cell.value);
            sheet1.align(j + 1, i + 1, cell.horizontalAlign);
            sheet1.valign(j + 1, i + 1, cell.verticalAlign === "middle" ? "center" : cell.verticalAlign);

            if (isColorDefined(cell.backgroundColor)) {
                sheet1.fill(j + 1, i + 1, {
                    type: 'solid',
                    fgColor: 'FF' + rgbToHex(cell.backgroundColor),
                    bgColor: '64'
                });
            }

            sheet1.font(j + 1, i + 1, {
                family: '3',
                scheme: 'minor',
                sz: parseInt(cell.fontSize.replace("px", "")) * 18 / 24,
                bold: cell.fontWeight === "bold" || parseInt(cell.fontWeight, 10) >= 700,
                color: isColorDefined(cell.foregroundColor) ? ('FF' + rgbToHex(cell.foregroundColor)) : undefined
            });

            sheet1.border(j + 1, i + 1, {
                left: getBorderStyle(cell.border.left),
                top: getBorderStyle(cell.border.top),
                right: getBorderStyle(cell.border.right),
                bottom: getBorderStyle(cell.border.bottom)
            });
        }

        sheet1.height(i + 1, maxHeight * 18 / 24);
    }

    try {
        workbook.save(function (err) {
            if (err)
                return cb(err);

            cb(null, fs.createReadStream(path.join(options.tmpDir, id + ".xlsx")));
        });
    }
    catch(e) {
        e.message = JSON.stringify(e.message);
        cb(e);
    }
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

