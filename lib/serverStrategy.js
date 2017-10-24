var Phantom = require("phantom-workers"),
    fs = require("fs"),
    path = require("path");

var phantoms = {};

function ensurePhantom(phantom, cb) {
    if (phantom.started)
        return cb();

    phantom.startCb = phantom.startCb || [];
    phantom.startCb.push(cb);

    if (phantom.starting)
        return;

    phantom.starting = true;

    phantom.start(function(startErr) {
        phantom.started = true;
        phantom.starting = false;
        phantom.startCb.forEach(function(cb) { cb(startErr); })
    });
}

module.exports = function(options, html, id, cb) {
    var phantomInstanceId = options.phantomPath || "default";
    var phantom

    options.numberOfWorkers = options.numberOfWorkers || 2;
    options.pathToPhantomScript = options.pathToPhantomScript || path.join(__dirname, "scripts", "serverScript.js");

    if (!phantoms[phantomInstanceId]) {
        phantoms[phantomInstanceId] = Phantom(options);
    }

    phantom = phantoms[phantomInstanceId];

    ensurePhantom(phantom, function(err) {
        if (err)
            return cb(err);

        phantom.execute(html, function (err, res) {
            if (err)
                return cb(err);

            cb(null, res);
        });
    })
};

module.exports.kill = function() {
    Object.keys(phantoms).forEach(function(key) {
        var phantom = phantoms[key]
        if (!phantom.started)
            return;

        phantom.started = false;
        phantom.startCb = [];
        return phantom.kill();
    });
    phantoms = {}
}
