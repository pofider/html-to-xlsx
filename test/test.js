require('should')
const util = require('util')
const path = require('path')
const fs = require('fs')
const uuid = require('uuid/v4')
const xlsx = require('xlsx')
const chromePageEval = require('chrome-page-eval')
const phantomPageEval = require('phantom-page-eval')
const puppeteer = require('puppeteer')
const phantomPath = require('phantomjs').path
const tmpDir = path.join(__dirname, 'temp')

const writeFileAsync = util.promisify(fs.writeFile)

const extractTableScriptFn = fs.readFileSync(
  path.join(__dirname, '../lib/scripts/conversionScript.js')
).toString()

const chromeEval = chromePageEval({
  puppeteer
})

const phantomEval = phantomPageEval({
  phantomPath,
  tmpDir,
  clean: false
})

async function createHtmlFile (html) {
  const outputPath = path.join(tmpDir, `${uuid()}.html`)

  await writeFileAsync(outputPath, html)

  return outputPath
}

function rmDir (dirPath) {
  if (!fs.existsSync(dirPath)) {
    fs.mkdirSync(dirPath)
  }

  let files

  try {
    files = fs.readdirSync(dirPath)
  } catch (e) {
    return
  }

  if (files.length > 0) {
    for (let i = 0; i < files.length; i++) {
      let filePath = `${dirPath}/${files[i]}`

      try {
        if (fs.statSync(filePath).isFile()) {
          fs.unlinkSync(filePath)
        }
      } catch (e) { }
    }
  }
}

describe('html extraction', () => {
  beforeEach(() => {
    rmDir(tmpDir)
  })

  describe('chrome-strategy', () => {
    common(chromeEval)
  })

  describe('phantom-strategy', () => {
    common(phantomEval)
  })

  function common (pageEval) {
    it('should build simple table', async () => {
      const table = await pageEval({
        html: await createHtmlFile(`<table><tr><td>1</td></tr></table>`),
        scriptFn: extractTableScriptFn
      })

      table.rows.should.have.length(1)
      table.rows[0].should.have.length(1)
      table.rows[0][0].value.should.be.eql('1')
    })

    it('should parse backgroud color', async () => {
      const table = await pageEval({
        html: await createHtmlFile(`<table><tr><td style='background-color:red'>1</td></tr></table>`),
        scriptFn: extractTableScriptFn
      })

      table.rows[0][0].backgroundColor[0].should.be.eql('255')
    })

    it('should parse foregorund color', async () => {
      const table = await pageEval({
        html: await createHtmlFile(`<table><tr><td style='color:red'>1</td></tr></table>`),
        scriptFn: extractTableScriptFn
      })

      table.rows[0][0].foregroundColor[0].should.be.eql('255')
    })

    it('should parse fontsize', async () => {
      const table = await pageEval({
        html: await createHtmlFile(`<table><tr><td style='font-size:19px'>1</td></tr></table>`),
        scriptFn: extractTableScriptFn
      })

      table.rows[0][0].fontSize.should.be.eql('19px')
    })

    it('should parse verticalAlign', async () => {
      const table = await pageEval({
        html: await createHtmlFile(`<table><tr><td style='vertical-align:bottom'>1</td></tr></table>`),
        scriptFn: extractTableScriptFn
      })

      table.rows[0][0].verticalAlign.should.be.eql('bottom')
    })

    it('should parse horizontal align', async () => {
      const table = await pageEval({
        html: await createHtmlFile(`<table><tr><td style='text-align:left'>1</td></tr></table>`),
        scriptFn: extractTableScriptFn
      })

      table.rows[0][0].horizontalAlign.should.be.eql('left')
    })

    it('should parse width', async () => {
      const table = await pageEval({
        html: await createHtmlFile(`<table><tr><td style='width:19px'>1</td></tr></table>`),
        scriptFn: extractTableScriptFn
      })

      table.rows[0][0].width.should.be.ok()
    })

    it('should parse height', async () => {
      const table = await pageEval({
        html: await createHtmlFile(`<table><tr><td style='height:19px'>1</td></tr></table>`),
        scriptFn: extractTableScriptFn
      })

      table.rows[0][0].height.should.be.ok()
    })

    it('should parse border', async () => {
      const table = await pageEval({
        html: await createHtmlFile(`<table><tr><td style='border-style:solid;'>1</td></tr></table>`),
        scriptFn: extractTableScriptFn
      })

      table.rows[0][0].border.left.should.be.eql('solid')
      table.rows[0][0].border.right.should.be.eql('solid')
      table.rows[0][0].border.bottom.should.be.eql('solid')
      table.rows[0][0].border.top.should.be.eql('solid')
    })

    it('should parse overflow', async () => {
      const table = await pageEval({
        html: await createHtmlFile(`<table><tr><td style='overflow:scroll;'>1234567789012345678912457890</td></tr></table>`),
        scriptFn: extractTableScriptFn
      })

      table.rows[0][0].wrapText.should.be.eql('scroll')
    })

    it('should parse backgroud color from styles with line endings', async () => {
      const table = await pageEval({
        html: await createHtmlFile(`<style> td { \n background-color: red \n } </style><table><tr><td>1</td></tr></table>`),
        scriptFn: extractTableScriptFn
      })

      table.rows[0][0].backgroundColor[0].should.be.eql('255')
    })

    it('should work for long tables', async function () {
      this.timeout(7000)

      let rows = ''

      for (let i = 0; i < 10000; i++) {
        rows += '<tr><td>1</td></tr>'
      }

      const table = await pageEval({
        html: await createHtmlFile(`<table>${rows}</table>`),
        scriptFn: extractTableScriptFn
      })

      table.rows.should.have.length(10000)
    })

    it('should parse colspan', async () => {
      const table = await pageEval({
        html: await createHtmlFile(`<table><tr><td colspan='6'></td><td>Column 7</td></tr></table>`),
        scriptFn: extractTableScriptFn
      })

      table.rows[0][0].colspan.should.be.eql(6)
      table.rows[0][1].value.should.be.eql('Column 7')
    })

    it('should parse rowspan', async () => {
      const table = await pageEval({
        html: await createHtmlFile(`<table><tr><td rowspan='2'>Col 1</td><td>Col 2</td></tr></table>`),
        scriptFn: extractTableScriptFn
      })

      table.rows[0][0].rowspan.should.be.eql(2)
      table.rows[0][0].value.should.be.eql('Col 1')
      table.rows[0][1].value.should.be.eql('Col 2')
    })

    it('should parse complex rowspan', async () => {
      const table = await pageEval({
        html: await createHtmlFile(`
          <table>
            <tr>
              <td rowspan='3'>Row 1 Col 1</td><td>Row 1 Col 2</td>
              <td>Row 1 Col 3</td><td>Row 1 Col 4</td>
            </tr>
            <tr>
              <td rowspan='2'>Row 2 Col 1</td>
              <td rowspan='2'>Row 2 Col 2</td><td>Row 2 Col 3</td>
            </tr>
            <tr>
              <td>Row 3 Col 3</td>
            </tr>
          </table>
        `),
        scriptFn: extractTableScriptFn
      })

      table.rows[0][0].rowspan.should.be.eql(3)
      table.rows[0][0].value.should.be.eql('Row 1 Col 1')
      table.rows[1][1].value.should.be.eql('Row 2 Col 2')
    })
  }
})

describe('html to xlsx conversion with strategy', () => {
  describe('chrome-strategy', () => {
    commonConversion(chromeEval)
  })
  describe('phantom-strategy', () => {
    commonConversion(phantomEval)
  })

  function commonConversion (pageEval) {
    let conversion

    beforeEach(function () {
      rmDir(tmpDir)

      conversion = require('../')({
        tmpDir: tmpDir,
        extract: pageEval
      })
    })

    it('should not fail', async () => {
      const stream = await conversion('<table><tr><td>hello</td></tr>')

      stream.should.have.property('readable')
    })

    // enable this test when we have a fix for #21
    it.skip('should not fail when last cell of a row has rowspan', async () => {
      const stream = await conversion('<table><tr><td>hello</td></tr>')

      stream.should.have.property('readable')
    })

    it('should callback error when input contains invalid characters', async () => {
      return (
        conversion('<table><tr><td></td></tr></table>')
      ).should.be.rejected()
    })

    it('should be able to parse xlsx', async () => {
      const stream = await conversion('<table><tr><td>hello</td></tr>')

      const bufs = []

      stream.on('data', (d) => { bufs.push(d) })
      stream.on('end', () => {
        const buf = Buffer.concat(bufs)
        xlsx.read(buf).Strings[0].t.should.be.eql('hello')
      })
    })

    it('should translate ampersands', async () => {
      const stream = await conversion('<table><tr><td>& &</td></tr>')

      const bufs = []

      stream.on('data', (d) => { bufs.push(d) })
      stream.on('end', () => {
        const buf = Buffer.concat(bufs)
        xlsx.read(buf).Strings[0].t.should.be.eql('& &')
      })
    })

    it('should callback error when row doesn\'t contain cells', async () => {
      return (
        conversion('<table><tr>Hello</tr></table>')
      ).should.be.rejected()
    })
  }
})
