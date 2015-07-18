var system = require('system');
var fs = require('fs');
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

page.onLoadFinished = function (status) {
    var result = $conversion
    console.log(JSON.stringify(result));
    phantom.exit();
};

var stream = fs.open(system.args[1], "r");
page.setContent(stream.read(), "http://localhost");
stream.close();

