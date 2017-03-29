var should = require("should"),
    path = require("path"),
    fs = require("fs"),
    xlsx = require('xlsx'),
    tmpDir = path.join(__dirname, "temp"),
    phantomServerStrategy = require("../lib/serverStrategy.js"),
    dedicatedProcessStrategy = require("../lib/dedicatedProcessStrategy.js");

describe("html extraction", function () {
    options = {
        tmpDir: tmpDir
    };

    beforeEach(function() {
        rmDir(tmpDir);
    })

    describe("phantom-server", function() {
       common(phantomServerStrategy);
    });

    describe("dedicated-process", function() {
        common(dedicatedProcessStrategy);
    });

    describe("dedicated-process use phantomJSPath", function() {
        common(dedicatedProcessStrategy);
    });

    function common(strategy) {
        it("should build simple table", function (done) {
          if (strategy === "") {
            options.phantomJSPath =  path.join(__dirname, "../node_modules/phantomjs/bin/phantomjs");
          }
            strategy(options, "<table><tr><td>1</td></tr></table>", "", function (err, table) {
                if (err)
                    return done(err);

                table.rows.should.have.length(1);
                table.rows[0].should.have.length(1);
                table.rows[0][0].value.should.be.eql('1');
                done();
            });
        });

        it("should parse backgroud color", function (done) {
            strategy(options, "<table><tr><td style='background-color:red'>1</td></tr></table>", "", function (err, table) {
                if (err)
                    return done(err);

                table.rows[0][0].backgroundColor[0].should.be.eql('255');
                done();
            });
        });

        it("should parse foregorund color", function (done) {
            strategy(options, "<table><tr><td style='color:red'>1</td></tr></table>", "", function (err, table) {
                if (err)
                    return done(err);

                table.rows[0][0].foregroundColor[0].should.be.eql('255');
                done();
            });
        });

        it("should parse fontsize", function (done) {
            strategy(options, "<table><tr><td style='font-size:19px'>1</td></tr></table>", "", function (err, table) {
                if (err)
                    return done(err);

                table.rows[0][0].fontSize.should.be.eql('19px');
                done();
            });
        });

        it("should parse verticalAlign", function (done) {
            strategy(options, "<table><tr><td style='vertical-align:bottom'>1</td></tr></table>", "", function (err, table) {
                if (err)
                    return done(err);

                table.rows[0][0].verticalAlign.should.be.eql('bottom');
                done();
            });
        });

        it("should parse horizontal align", function (done) {
            strategy(options, "<table><tr><td style='text-align:left'>1</td></tr></table>", "", function (err, table) {
                if (err)
                    return done(err);

                table.rows[0][0].horizontalAlign.should.be.eql('left');
                done();
            });
        });

        it("should parse width", function (done) {
            strategy(options, "<table><tr><td style='width:19px'>1</td></tr></table>", "", function (err, table) {
                if (err)
                    return done(err);

                table.rows[0][0].width.should.be.eql.ok;
                done();
            });
        });

        it("should parse height", function (done) {
            strategy(options, "<table><tr><td style='height:19px'>1</td></tr></table>", "", function (err, table) {
                if (err)
                    return done(err);

                table.rows[0][0].height.should.be.ok;
                done();
            });
        });

        it("should parse border", function (done) {
            strategy(options, "<table><tr><td style='border-style:solid;'>1</td></tr></table>", "", function (err, table) {
                if (err)
                    return done(err);


                table.rows[0][0].border.left.should.be.eql('solid');
                table.rows[0][0].border.right.should.be.eql('solid');
                table.rows[0][0].border.bottom.should.be.eql('solid');
                table.rows[0][0].border.top.should.be.eql('solid');
                done();
            });
        });

        it("should parse backgroud color from styles with line endings", function (done) {
            strategy(options, "<style> td { \n background-color: red \n } </style><table><tr><td>1</td></tr></table>", "", function (err, table) {
                if (err)
                    return done(err);

                table.rows[0][0].backgroundColor[0].should.be.eql('255');
                done();
            });
        });

        it("should work for long tables", function (done) {
            this.timeout(7000);
            var rows = "";
            for (var i = 0; i < 10000; i++) {
                rows += "<tr><td>1</td></tr>"
            }
            strategy(options, "<table>" + rows + "</table>", "", function (err, table) {
                if (err)
                    return done(err);

                table.rows.should.have.length(10000);
                done();
            });
        });

        it("should parse colspan", function (done) {
            strategy(options, "<table><tr><td colspan='6'></td><td>Column 7</td></tr></table>", "", function (err, table) {
                if (err)
                    return done(err);

                table.rows[0][0].colspan.should.be.eql(6);
                table.rows[0][1].value.should.be.eql("Column 7");
                done();
            });
        });

        it("should parse rowspan", function (done) {
            strategy(options, "<table><tr><td rowspan='2'>Col 1</td><td>Col 2</td></tr></table>", "", function (err, table) {
                if (err)
                    return done(err);

                table.rows[0][0].rowspan.should.be.eql(2);
                table.rows[0][0].value.should.be.eql("Col 1");
                table.rows[0][1].value.should.be.eql("Col 2");
                done();
            });
        });        
    }
});


describe("html to xlsx conversion in phantom", function () {

    describe("phantom-server", function () {
        commonConversion("phantom-server");
    });

    describe("dedicated-process", function () {
        commonConversion("dedicated-process");
    });

    function commonConversion(strategyName) {

        var conversion;

        beforeEach(function () {
            rmDir(tmpDir);
            conversion = require("../lib/conversion.js")({
                tmpDir: tmpDir,
                numberOfWorkers: 1,
                strategy: strategyName
            });
        });

        it("should not fail", function (done) {
            conversion("<table><tr><td>hello</td></tr>", function (err, res) {
                if (err)
                    return done(err);

                res.should.have.property("readable");
                done();
            });
        });

        it("should callback error when input contains invalid characters", function (done) {
            conversion("<table><tr><td></td></tr></table>", function (err, res) {
                if (err)
                    return done();

                done(new Error('Should have failed'));
            });
        });

        it("should be able to parse xlsx", function (done) {
            conversion("<table><tr><td>hello</td></tr>", function (err, res) {
                if (err)
                    return done(err);

                var bufs = [];
                res.on('data', function(d){ bufs.push(d); });
                res.on('end', function() {
                    var buf = Buffer.concat(bufs);
                    xlsx.read(buf).Strings[0].t.should.be.eql('hello');
                    done();
                })
            });
        });

        it("should translate ampersands", function (done) {
            conversion("<table><tr><td>& &</td></tr>", function (err, res) {
                if (err)
                    return done(err);

                var bufs = [];
                res.on('data', function(d){ bufs.push(d); });
                res.on('end', function() {
                    var buf = Buffer.concat(bufs);
                    xlsx.read(buf).Strings[0].t.should.be.eql('& &');
                    done();
                })
            });
        });

        it("should callback error when row doesn't contain cells", function (done) {         
            conversion("<table><tr>Hello</tr></table>", function (err, res) {
                if (err)
                    return done();
                
                done(new Error('It should have callback error'));
            });
        });
    }

    rmDir = function (dirPath) {
        if (!fs.existsSync(dirPath))
            fs.mkdirSync(dirPath);

        try {
            var files = fs.readdirSync(dirPath);
        }
        catch (e) {
            return;
        }
        if (files.length > 0)
            for (var i = 0; i < files.length; i++) {
                var filePath = dirPath + '/' + files[i];
                try {
                    if (fs.statSync(filePath).isFile()) {
                        fs.unlinkSync(filePath);
                    }
                }
                catch(e) { }
            }
    };
});
