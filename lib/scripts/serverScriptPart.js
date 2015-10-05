var system = require('system'),
    webserver = require('webserver').create();

var port = require("system").env['PHANTOM_WORKER_PORT'];
var host = require("system").env['PHANTOM_WORKER_HOST'];

var service = webserver.listen(host + ':' + port, function (req, res) {
    page = require('webpage').create();

    page.onResourceRequested = function (request, networkRequest) {
        //potentially dangerous request
        if (request.url.lastIndexOf("file:///", 0) === 0) {
            networkRequest.changeUrl(request.url.replace("file:///", "http://"));
            return;
        }

        //to support cdn like format //cdn.jquery...
        if (request.url.lastIndexOf("file://", 0) === 0 && request.url.lastIndexOf("file:///", 0) !== 0) {
            networkRequest.changeUrl(request.url.replace("file://", "http://"));
        }
    };

    page.onLoadFinished = function(status) {
        var result = $conversion

        res.statusCode = 200;
        res.write(JSON.stringify(result));
        res.close();
    };

    page.setContent(JSON.parse(req.post), "http://localhost");
});
