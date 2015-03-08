var should = require("should"),
    path = require("path"),
    fs = require("fs"),
    tmpDir = path.join(__dirname, "temp"),
    phantom = require("phantom-workers")({
        tmpDir: tmpDir,
        numberOfWorkers: 1,
        pathToPhantomScript: path.join(__dirname, "../", "lib", "phantomScript.js")
    }),
    conversion = require("../lib/conversion.js")({
        tmpDir: tmpDir,
        numberOfWorkers: 1
    });

describe("html extraction", function () {
    beforeEach(function (done) {
        phantom.start(function (err) {
            if (err)
                return done;

            done();
        });
    });

    afterEach(function () {
        phantom.kill();
    });

    it("should build simple table", function (done) {
        phantom.execute("<table><tr><td>1</td></tr></table>", function (err, table) {
            if (err)
                return cb(err);

            table.rows.should.have.length(1);
            table.rows[0].should.have.length(1);
            table.rows[0][0].value.should.be.eql('1');
            done();
        });
    });

    it("should parse backgroud color", function (done) {
        phantom.execute("<table><tr><td style='background-color:red'>1</td></tr></table>", function (err, table) {
            if (err)
                return cb(err);

            table.rows[0][0].backgroundColor[0].should.be.eql('255');
            done();
        });
    });

    it("should parse foregorund color", function (done) {
        phantom.execute("<table><tr><td style='color:red'>1</td></tr></table>", function (err, table) {
            if (err)
                return cb(err);

            table.rows[0][0].foregroundColor[0].should.be.eql('255');
            done();
        });
    });

    it("should parse fontsize", function (done) {
        phantom.execute("<table><tr><td style='font-size:19px'>1</td></tr></table>", function (err, table) {
            if (err)
                return cb(err);

            table.rows[0][0].fontSize.should.be.eql('19px');
            done();
        });
    });

    it("should parse verticalAlign", function (done) {
        phantom.execute("<table><tr><td style='vertical-align:bottom'>1</td></tr></table>", function (err, table) {
            if (err)
                return cb(err);

            table.rows[0][0].verticalAlign.should.be.eql('bottom');
            done();
        });
    });

    it("should parse horizontal align", function (done) {
        phantom.execute("<table><tr><td style='text-align:left'>1</td></tr></table>", function (err, table) {
            if (err)
                return cb(err);

            table.rows[0][0].horizontalAlign.should.be.eql('left');
            done();
        });
    });

    it("should parse width", function (done) {
        phantom.execute("<table><tr><td style='width:19px'>1</td></tr></table>", function (err, table) {
            if (err)
                return cb(err);

            table.rows[0][0].width.should.be.eql.ok;
            done();
        });
    });

    it("should parse height", function (done) {
        phantom.execute("<table><tr><td style='height:19px'>1</td></tr></table>", function (err, table) {
            if (err)
                return cb(err);

            table.rows[0][0].height.should.be.ok;
            done();
        });
    });

    it("should parse border", function (done) {
        phantom.execute("<table><tr><td style='border-style:solid;'>1</td></tr></table>", function (err, table) {
            if (err)
                return cb(err);


            table.rows[0][0].border.left.should.be.eql('solid');
            table.rows[0][0].border.right.should.be.eql('solid');
            table.rows[0][0].border.bottom.should.be.eql('solid');
            table.rows[0][0].border.top.should.be.eql('solid');
            done();
        });
    });
});

describe("phantom html to pdf", function () {

    beforeEach(function () {
        rmDir(tmpDir);
    });

    it("should not fail", function (done) {
        conversion("<table><tr><td>hello</td></tr>", function (err, res) {
            if (err)
                return done(err);

            res.should.have.property("readable");
            done();
        });
    });

    rmDir = function (dirPath) {
        try {
            var files = fs.readdirSync(dirPath);
        }
        catch (e) {
            return;
        }
        if (files.length > 0)
            for (var i = 0; i < files.length; i++) {
                var filePath = dirPath + '/' + files[i];
                if (fs.statSync(filePath).isFile())
                    fs.unlinkSync(filePath);
            }
    };
});