#html-to-xlsx
[![Build Status](https://travis-ci.org/pofider/html-to-xlsx.png?branch=master)](https://travis-ci.org/pofider/html-to-xlsx)

**node.js html to xlsx transformation**

Transformation only supports html table and several basic style properties. No images or charts are currently supported. 

```js
var conversion = require("html-to-xlsx")();
conversion("<table><tr><td>cell value</td></tr></table>" }, function(err, stream){
  //readable stream to xlsx file
  stream.pipe(res);
});
```

##Supported properties

##Options
```js
var conversion = require("phantom-html-to-pdf")({
    /* number of allocated phantomjs processes */
	numberOfWorkers: 2,
	/* timeout in ms for html conversion, when the timeout is reached, the phantom process is recycled */
	timeout: 5000,
	/* directory where are stored temporary html and pdf files, use something like npm package reaper to clean this up */
	tmpDir: "os/tmpdir"
});
```

##License
See [license](https://github.com/pofider/phantom-html-to-pdf/blob/master/LICENSE)