var path = require("path"),
    childProcess = require('child_process'),
    phantomjs = require('phantomjs'),
    fs = require("fs");


module.exports = function(options, html, id, cb) {
    var htmlFilePath = path.resolve(path.join(options.tmpDir, id + ".html"));

    fs.writeFile(htmlFilePath, html, function (err) {
        if (err)
            return cb(err);

        var childArgs = [
            options.standaloneScriptPath || path.join(__dirname, "scripts", "standaloneScript.js"),
            htmlFilePath,
            '--ignore-ssl-errors=yes',
            '--web-security=false',
            '--ssl-protocol=any'
        ];

        var isDone = false;

        var output = "";
        var path2PhantomJS = options.phantomJSPath || phantomjs.path;
        var child = childProcess.execFile(path2PhantomJS, childArgs, { maxBuffer: 1024 * 50000 }, function (err, stdout, stderr) {
            if (isDone)
                return;

            isDone = true;
            if (err) {
                return cb(err);
            }

            cb(null, JSON.parse(output));
        });
        child.stdout.on("data", function(data) {
            output += data;
        });

        setTimeout(function() {
            if (isDone)
                return;

            isDone = true;
            cb(new Error("Timeout when executing in phantom"));
        }, options.timeout);
    });
};
