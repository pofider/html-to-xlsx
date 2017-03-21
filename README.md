# html-to-xlsx
[![NPM Version](http://img.shields.io/npm/v/html-to-xlsx.svg?style=flat-square)](https://npmjs.com/package/html-to-xlsx)
[![License](http://img.shields.io/npm/l/html-to-xlsx.svg?style=flat-square)](http://opensource.org/licenses/MIT)
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

## Supported properties
- `background-color` - cell background color
- `color` - cell foreground color
- `border-left-style` - as well as positions will be transformed into excel cells borders
- `text-align` - text horizontal align in the excel cell
- `vertical-align` - vertical align in the excel cell
- `width` - the excel column will get the highest width, it can be little bit inaccurate because of pixel to excel points conversion
- `height` - the excel row will get the highest height
- `font-size` - font size
- `colspan` - numeric value that merges current column with columns to the right
- `rowspan` - numeric value that merges current row with rows below.  


## Options
```js
var conversion = require("html-to-xlsx")({
    /* number of allocated phantomjs processes */
	numberOfWorkers: 2,
	/* timeout in ms for html conversion, when the timeout is reached, the phantom process is recycled */
	timeout: 5000,
	/* directory where are stored temporary html and pdf files, use something like npm package reaper to clean this up */
	tmpDir: "os/tmpdir",
	/* optional port range where to start phantomjs server */
    portLeftBoundary: 1000,
    portRightBoundary: 2000,
    /* optional hostname where to start phantomjs server */
    host: '127.0.0.1'
});
```

## License
See [license](https://github.com/pofider/html-to-xlsx/blob/master/LICENSE)
