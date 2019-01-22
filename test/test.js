const should = require('should')
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
    it('should parse simple table', async () => {
      const table = await pageEval({
        html: await createHtmlFile(`
          <table>
            <tr>
              <td>1</td>
            </tr>
          </table>
        `),
        scriptFn: extractTableScriptFn
      })

      table.rows.should.have.length(1)
      table.rows[0].should.have.length(1)
      table.rows[0][0].value.should.be.eql('1')
    })

    it('should parse value', async () => {
      const table = await pageEval({
        html: await createHtmlFile(`
          <table>
            <tr>
              <td>node.js & javascript</td>
            </tr>
          </table>
        `),
        scriptFn: extractTableScriptFn
      })

      table.rows.should.have.length(1)
      table.rows[0].should.have.length(1)
      table.rows[0][0].value.should.be.eql('node.js &amp; javascript')
    })

    it('should parse un-escaped value', async () => {
      const table = await pageEval({
        html: await createHtmlFile(`
          <table>
            <tr>
              <td>node.js & javascript</td>
            </tr>
          </table>
        `),
        scriptFn: extractTableScriptFn
      })

      table.rows.should.have.length(1)
      table.rows[0].should.have.length(1)
      table.rows[0][0].valueText.should.be.eql('node.js & javascript')
    })

    it('should parse cell data type', async () => {
      const table = await pageEval({
        html: await createHtmlFile(`
          <table>
            <tr>
              <td data-cell-type="number">10</td>
              <td data-cell-type="boolean">1</td>
              <td data-cell-type="date">2019-01-22</td>
              <td data-cell-type="datetime">2019-01-22T17:31:36.242Z</td>
              <td data-cell-type="formula">=SUM(A1, B1)</td>
            </tr>
          </table>
        `),
        scriptFn: extractTableScriptFn
      })

      table.rows.should.have.length(1)
      table.rows[0].should.have.length(5)
      table.rows[0][0].type = 'number'
      table.rows[0][1].type = 'boolean'
      table.rows[0][2].type = 'date'
      table.rows[0][3].type = 'datetime'
      table.rows[0][4].type = 'formula'
    })

    it('should parse format str and enum', async () => {
      const table = await pageEval({
        html: await createHtmlFile(`
          <table>
            <tr>
              <td data-cell-type="number" data-cell-format-str="0.00">10</td>
              <td data-cell-type="number" data-cell-format-enum="3">100000</td>
            </tr>
          </table>
        `),
        scriptFn: extractTableScriptFn
      })

      table.rows.should.have.length(1)
      table.rows[0][0].formatStr = '0.00'
      table.rows[0][1].formatEnum = 3
    })

    it('should parse background color', async () => {
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
        html: await createHtmlFile(`
          <table>
            <tr>
              <td style='border-style:solid;'>1</td>
            </tr>
          </table>
        `),
        scriptFn: extractTableScriptFn
      })

      table.rows[0][0].border.left.should.be.eql('solid')
      table.rows[0][0].border.right.should.be.eql('solid')
      table.rows[0][0].border.bottom.should.be.eql('solid')
      table.rows[0][0].border.top.should.be.eql('solid')
    })

    it('should parse complex border', async () => {
      const table = await pageEval({
        html: await createHtmlFile(`
          <table>
            <tr>
              <td style='border-left: 1px solid red;'>1</td>
            </tr>
          </table>
        `),
        scriptFn: extractTableScriptFn
      })

      table.rows[0][0].border.leftColor.should.be.eql(['255', '0', '0'])
      table.rows[0][0].border.leftWidth.should.be.eql('1px')
      table.rows[0][0].border.leftStyle.should.be.eql('solid')
    })

    it('should parse overflow', async () => {
      const table = await pageEval({
        html: await createHtmlFile(`<table><tr><td style='overflow:scroll;'>1234567789012345678912457890</td></tr></table>`),
        scriptFn: extractTableScriptFn
      })

      table.rows[0][0].wrapText.should.be.eql('scroll')
    })

    it('should parse textDecoration', async () => {
      const table = await pageEval({
        html: await createHtmlFile(`
          <table>
            <tr>
              <td style='text-decoration: underline'>
                1234
              </td>
            </tr>
          </table>
        `),
        scriptFn: extractTableScriptFn
      })

      table.rows[0][0].textDecoration.line.should.be.eql('underline')
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
              <td rowspan='3'>Row 1 Col 1</td>
              <td>Row 1 Col 2</td>
              <td>Row 1 Col 3</td>
              <td>Row 1 Col 4</td>
            </tr>
            <tr>
              <td rowspan='2'>Row 2 Col 1</td>
              <td rowspan='2'>Row 2 Col 2</td>
              <td>Row 2 Col 3</td>
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

    it('should parse th elements', async () => {
      const table = await pageEval({
        html: await createHtmlFile(`
          <table>
            <tr>
              <th>col1</th>
              <th>col2</th>
            </tr>
            <tr>
              <td>1</td>
              <td>2</td>
            </tr>
          </table>
        `),
        scriptFn: extractTableScriptFn
      })

      table.rows.should.have.length(2)
      table.rows[0][0].value.should.be.eql('col1')
      table.rows[0][1].value.should.be.eql('col2')
      table.rows[1][0].value.should.be.eql('1')
      table.rows[1][1].value.should.be.eql('2')
    })
  }
})

describe('html to xlsx conversion with strategy', () => {
  describe('legacy', () => {
    describe('chrome-strategy', () => {
      commonConversion(chromeEval, true)
    })
    describe('phantom-strategy', () => {
      commonConversion(phantomEval, true)
    })
  })

  describe('standard', () => {
    describe('chrome-strategy', () => {
      commonConversion(chromeEval)
    })
    describe('phantom-strategy', () => {
      commonConversion(phantomEval)
    })
  })

  function commonConversion (pageEval, legacy = false) {
    let conversion

    beforeEach(function () {
      rmDir(tmpDir)

      const conversionFn = require('../')({
        tmpDir: tmpDir,
        extract: pageEval
      })

      conversion = (html, convertOptions) => {
        const opts = {
          ...convertOptions,
          legacy
        }

        return conversionFn(html, opts)
      }
    })

    it('should not fail', async () => {
      const stream = await conversion(`
        <table>
          <tr>
            <td>hello</td>
          </tr>
        </table
      `)

      stream.should.have.property('readable')
    })

    it('default sheet name should be Sheet1', async () => {
      const stream = await conversion(`
        <table>
          <tr>
            <td>hello</td>
          </tr>
        </table
      `)

      const parsedXlsx = await new Promise((resolve, reject) => {
        const bufs = []

        stream.on('error', reject)
        stream.on('data', (d) => { bufs.push(d) })

        stream.on('end', () => {
          const buf = Buffer.concat(bufs)
          resolve(xlsx.read(buf))
        })
      })

      should(parsedXlsx.SheetNames[0]).be.eql('Sheet1')
    })

    if (!legacy) {
      it('shoule be able to set custom sheet name', async () => {
        const parseXlsx = (xlsxStream) => {
          return new Promise((resolve, reject) => {
            const bufs = []

            xlsxStream.on('error', reject)
            xlsxStream.on('data', (d) => { bufs.push(d) })

            xlsxStream.on('end', () => {
              const buf = Buffer.concat(bufs)
              resolve(xlsx.read(buf))
            })
          })
        }

        let stream = await conversion(`
          <table name="custom">
            <tr>
              <td>1</td>
            </tr>
          </table>
        `)

        let parsedXlsx = await parseXlsx(stream)

        should(parsedXlsx.SheetNames[0]).be.eql('custom')

        stream = await conversion(`
          <table data-sheet-name="custom2">
            <tr>
              <td>1</td>
            </tr>
          </table>
        `)

        parsedXlsx = await parseXlsx(stream)

        should(parsedXlsx.SheetNames[0]).be.eql('custom2')
      })

      it('should be able to set cell with datatypes', async () => {
        const stream = await conversion(`
          <table>
            <tr>
              <td data-cell-type="number">10</td>
              <td data-cell-type="number">10</td>
              <td data-cell-type="boolean">1</td>
              <td data-cell-type="date">2019-01-22</td>
              <td data-cell-type="datetime">2019-01-22T17:31:36.242Z</td>
              <td data-cell-type="formula">=SUM(A1, B1)</td>
            </tr>
          </table>
        `)

        const parsedXlsx = await new Promise((resolve, reject) => {
          const bufs = []

          stream.on('error', reject)
          stream.on('data', (d) => { bufs.push(d) })

          stream.on('end', () => {
            const buf = Buffer.concat(bufs)
            resolve(xlsx.read(buf))
          })
        })

        should(parsedXlsx.Sheets[parsedXlsx.SheetNames[0]]['A1'].v).be.eql(10)
        should(parsedXlsx.Sheets[parsedXlsx.SheetNames[0]]['B1'].v).be.eql(10)
        should(parsedXlsx.Sheets[parsedXlsx.SheetNames[0]]['C1'].v).be.eql(true)
        should(parsedXlsx.Sheets[parsedXlsx.SheetNames[0]]['D1'].w).be.eql('1/22/19')
        should(parsedXlsx.Sheets[parsedXlsx.SheetNames[0]]['E1'].w).be.eql('1/22/19 17:31')
      })

      it('should be able to set cell format', async () => {
        const stream = await conversion(`
          <table>
            <tr>
              <td data-cell-type="number" data-cell-format-str="0.00">10</td>
              <td data-cell-type="number" data-cell-format-enum="3">100000</td>
            </tr>
          </table>
        `)

        const parsedXlsx = await new Promise((resolve, reject) => {
          const bufs = []

          stream.on('error', reject)
          stream.on('data', (d) => { bufs.push(d) })

          stream.on('end', () => {
            const buf = Buffer.concat(bufs)
            resolve(xlsx.read(buf))
          })
        })

        should(parsedXlsx.Sheets[parsedXlsx.SheetNames[0]]['A1'].v).be.eql(10)
        should(parsedXlsx.Sheets[parsedXlsx.SheetNames[0]]['A1'].w).be.eql('10.00')
        should(parsedXlsx.Sheets[parsedXlsx.SheetNames[0]]['B1'].v).be.eql(100000)
        should(parsedXlsx.Sheets[parsedXlsx.SheetNames[0]]['B1'].w).be.eql('100,000')
      })
    }

    it('should work with th elements', async () => {
      const stream = await conversion(`
        <table>
          <tr>
            <th>col1</th>
            <th>col2</th>
          </tr>
          <tr>
            <td>1</td>
            <td>2</td>
          </tr>
        </table>
      `)

      const parsedXlsx = await new Promise((resolve, reject) => {
        const bufs = []

        stream.on('error', reject)
        stream.on('data', (d) => { bufs.push(d) })

        stream.on('end', () => {
          const buf = Buffer.concat(bufs)
          resolve(xlsx.read(buf))
        })
      })

      should(parsedXlsx.Sheets[parsedXlsx.SheetNames[0]]['A1'].v).be.eql('col1')
      should(parsedXlsx.Sheets[parsedXlsx.SheetNames[0]]['B1'].v).be.eql('col2')
      should(parsedXlsx.Sheets[parsedXlsx.SheetNames[0]]['A2'].v).be.eql('1')
      should(parsedXlsx.Sheets[parsedXlsx.SheetNames[0]]['B2'].v).be.eql('2')
    })

    it('should not fail when last cell of a row has rowspan', async () => {
      const stream = await conversion(`
        <table>
          <tr>
            <td rowspan="2">Cell RowSpan</td>
          </tr>
          <tr>
            <td>Foo</td>
          </tr>
        </table>
      `)

      stream.should.have.property('readable')
    })

    it('should work when using special rowspan layout #1 (row with just one cell)', async () => {
      const stream = await conversion(`
        <table>
          <tr>
              <td rowspan="3">ROWSPAN 3</td>
          </tr>
          <tr>
              <td>Ipsum</td>
              <td>Data</td>
          </tr>
          <tr>
              <td>Hello</td>
              <td>World</td>
          </tr>
        </table>
      `)

      const parsedXlsx = await new Promise((resolve, reject) => {
        const bufs = []

        stream.on('error', reject)
        stream.on('data', (d) => { bufs.push(d) })

        stream.on('end', () => {
          const buf = Buffer.concat(bufs)
          resolve(xlsx.read(buf))
        })
      })

      should(parsedXlsx.Sheets[parsedXlsx.SheetNames[0]]['A1'].v).be.eql('ROWSPAN 3')
      should(parsedXlsx.Sheets[parsedXlsx.SheetNames[0]]['A2']).be.undefined()
      should(parsedXlsx.Sheets[parsedXlsx.SheetNames[0]]['B1'].v).be.eql('Ipsum')
      should(parsedXlsx.Sheets[parsedXlsx.SheetNames[0]]['C1'].v).be.eql('Data')
      should(parsedXlsx.Sheets[parsedXlsx.SheetNames[0]]['B2'].v).be.eql('Hello')
      should(parsedXlsx.Sheets[parsedXlsx.SheetNames[0]]['C2'].v).be.eql('World')

      should(parsedXlsx.Sheets[parsedXlsx.SheetNames[0]]['!merges'][0].s.r).be.eql(0)
      should(parsedXlsx.Sheets[parsedXlsx.SheetNames[0]]['!merges'][0].s.c).be.eql(0)
      should(parsedXlsx.Sheets[parsedXlsx.SheetNames[0]]['!merges'][0].e.r).be.eql(1)
      should(parsedXlsx.Sheets[parsedXlsx.SheetNames[0]]['!merges'][0].e.c).be.eql(0)
    })

    it('should work when using special rowspan layout #2 (row with just one cell, more rows)', async () => {
      const stream = await conversion(`
        <table>
          <tr>
              <td rowspan="3">ROWSPAN 3</td>
          </tr>
          <tr>
              <td>Ipsum</td>
              <td>Data</td>
          </tr>
          <tr>
              <td>Hello</td>
              <td>World</td>
          </tr>
          <tr>
              <td>Something</td>
              <td>Else</td>
          </tr>
        </table>
      `)

      const parsedXlsx = await new Promise((resolve, reject) => {
        const bufs = []

        stream.on('error', reject)
        stream.on('data', (d) => { bufs.push(d) })

        stream.on('end', () => {
          const buf = Buffer.concat(bufs)
          resolve(xlsx.read(buf))
        })
      })

      should(parsedXlsx.Sheets[parsedXlsx.SheetNames[0]]['A1'].v).be.eql('ROWSPAN 3')
      should(parsedXlsx.Sheets[parsedXlsx.SheetNames[0]]['A2']).be.undefined()
      should(parsedXlsx.Sheets[parsedXlsx.SheetNames[0]]['B1'].v).be.eql('Ipsum')
      should(parsedXlsx.Sheets[parsedXlsx.SheetNames[0]]['C1'].v).be.eql('Data')
      should(parsedXlsx.Sheets[parsedXlsx.SheetNames[0]]['B2'].v).be.eql('Hello')
      should(parsedXlsx.Sheets[parsedXlsx.SheetNames[0]]['C2'].v).be.eql('World')
      should(parsedXlsx.Sheets[parsedXlsx.SheetNames[0]]['A3'].v).be.eql('Something')
      should(parsedXlsx.Sheets[parsedXlsx.SheetNames[0]]['B3'].v).be.eql('Else')

      should(parsedXlsx.Sheets[parsedXlsx.SheetNames[0]]['!merges'][0].s.r).be.eql(0)
      should(parsedXlsx.Sheets[parsedXlsx.SheetNames[0]]['!merges'][0].s.c).be.eql(0)
      should(parsedXlsx.Sheets[parsedXlsx.SheetNames[0]]['!merges'][0].e.r).be.eql(1)
      should(parsedXlsx.Sheets[parsedXlsx.SheetNames[0]]['!merges'][0].e.c).be.eql(0)
    })

    it('should work when using special rowspan layout #3 (row with normal cells and first cell with rowspan)', async () => {
      const stream = await conversion(`
        <table>
          <tr>
              <td rowspan="3">ROWSPAN 3</td>
              <td>Header 2</td>
              <td>Header 3</td>
          </tr>
          <tr>
              <td>Ipsum</td>
              <td>Data</td>
          </tr>
          <tr>
              <td>Hello</td>
              <td>World</td>
          </tr>
        </table>
      `)

      const parsedXlsx = await new Promise((resolve, reject) => {
        const bufs = []

        stream.on('error', reject)
        stream.on('data', (d) => { bufs.push(d) })

        stream.on('end', () => {
          const buf = Buffer.concat(bufs)
          resolve(xlsx.read(buf))
        })
      })

      should(parsedXlsx.Sheets[parsedXlsx.SheetNames[0]]['A1'].v).be.eql('ROWSPAN 3')
      should(parsedXlsx.Sheets[parsedXlsx.SheetNames[0]]['A2']).be.undefined()
      should(parsedXlsx.Sheets[parsedXlsx.SheetNames[0]]['A3']).be.undefined()
      should(parsedXlsx.Sheets[parsedXlsx.SheetNames[0]]['B1'].v).be.eql('Header 2')
      should(parsedXlsx.Sheets[parsedXlsx.SheetNames[0]]['C1'].v).be.eql('Header 3')
      should(parsedXlsx.Sheets[parsedXlsx.SheetNames[0]]['B2'].v).be.eql('Ipsum')
      should(parsedXlsx.Sheets[parsedXlsx.SheetNames[0]]['C2'].v).be.eql('Data')
      should(parsedXlsx.Sheets[parsedXlsx.SheetNames[0]]['B3'].v).be.eql('Hello')
      should(parsedXlsx.Sheets[parsedXlsx.SheetNames[0]]['C3'].v).be.eql('World')

      should(parsedXlsx.Sheets[parsedXlsx.SheetNames[0]]['!merges'][0].s.r).be.eql(0)
      should(parsedXlsx.Sheets[parsedXlsx.SheetNames[0]]['!merges'][0].s.c).be.eql(0)
      should(parsedXlsx.Sheets[parsedXlsx.SheetNames[0]]['!merges'][0].e.r).be.eql(2)
      should(parsedXlsx.Sheets[parsedXlsx.SheetNames[0]]['!merges'][0].e.c).be.eql(0)
    })

    it('should work when using special rowspan layout #4 (row with all cells using rowspan)', async () => {
      const stream = await conversion(`
        <table>
          <tr>
              <td rowspan="3">NRO1</td>
              <td rowspan="3">NRO2</td>
              <td rowspan="3">NRO3</td>
              <td rowspan="3">NRO4</td>
          </tr>
          <tr>
              <td>Doc1.</td>
          </tr>
          <tr>
              <td>Doc2.</td>
          </tr>
        </table>
      `)

      const parsedXlsx = await new Promise((resolve, reject) => {
        const bufs = []

        stream.on('error', reject)
        stream.on('data', (d) => { bufs.push(d) })

        stream.on('end', () => {
          const buf = Buffer.concat(bufs)
          resolve(xlsx.read(buf))
        })
      })

      should(parsedXlsx.Sheets[parsedXlsx.SheetNames[0]]['A1'].v).be.eql('NRO1')
      should(parsedXlsx.Sheets[parsedXlsx.SheetNames[0]]['A2']).be.undefined()
      should(parsedXlsx.Sheets[parsedXlsx.SheetNames[0]]['B1'].v).be.eql('NRO2')
      should(parsedXlsx.Sheets[parsedXlsx.SheetNames[0]]['B2']).be.undefined()
      should(parsedXlsx.Sheets[parsedXlsx.SheetNames[0]]['C1'].v).be.eql('NRO3')
      should(parsedXlsx.Sheets[parsedXlsx.SheetNames[0]]['C2']).be.undefined()
      should(parsedXlsx.Sheets[parsedXlsx.SheetNames[0]]['D1'].v).be.eql('NRO4')
      should(parsedXlsx.Sheets[parsedXlsx.SheetNames[0]]['D2']).be.undefined()
      should(parsedXlsx.Sheets[parsedXlsx.SheetNames[0]]['E1'].v).be.eql('Doc1.')
      should(parsedXlsx.Sheets[parsedXlsx.SheetNames[0]]['E2'].v).be.eql('Doc2.')

      should(parsedXlsx.Sheets[parsedXlsx.SheetNames[0]]['!merges'][0].s.r).be.eql(0)
      should(parsedXlsx.Sheets[parsedXlsx.SheetNames[0]]['!merges'][0].s.c).be.eql(0)
      should(parsedXlsx.Sheets[parsedXlsx.SheetNames[0]]['!merges'][0].e.r).be.eql(1)
      should(parsedXlsx.Sheets[parsedXlsx.SheetNames[0]]['!merges'][0].e.c).be.eql(0)

      should(parsedXlsx.Sheets[parsedXlsx.SheetNames[0]]['!merges'][1].s.r).be.eql(0)
      should(parsedXlsx.Sheets[parsedXlsx.SheetNames[0]]['!merges'][1].s.c).be.eql(1)
      should(parsedXlsx.Sheets[parsedXlsx.SheetNames[0]]['!merges'][1].e.r).be.eql(1)
      should(parsedXlsx.Sheets[parsedXlsx.SheetNames[0]]['!merges'][1].e.c).be.eql(1)

      should(parsedXlsx.Sheets[parsedXlsx.SheetNames[0]]['!merges'][2].s.r).be.eql(0)
      should(parsedXlsx.Sheets[parsedXlsx.SheetNames[0]]['!merges'][2].s.c).be.eql(2)
      should(parsedXlsx.Sheets[parsedXlsx.SheetNames[0]]['!merges'][2].e.r).be.eql(1)
      should(parsedXlsx.Sheets[parsedXlsx.SheetNames[0]]['!merges'][2].e.c).be.eql(2)

      should(parsedXlsx.Sheets[parsedXlsx.SheetNames[0]]['!merges'][3].s.r).be.eql(0)
      should(parsedXlsx.Sheets[parsedXlsx.SheetNames[0]]['!merges'][3].s.c).be.eql(3)
      should(parsedXlsx.Sheets[parsedXlsx.SheetNames[0]]['!merges'][3].e.r).be.eql(1)
      should(parsedXlsx.Sheets[parsedXlsx.SheetNames[0]]['!merges'][3].e.c).be.eql(3)
    })

    it('should work when using special rowspan layout #5 (row with lot of cells using rowspan but last one without it)', async () => {
      const stream = await conversion(`
        <table>
          <tr>
              <td rowspan="3">NRO1</td>
              <td rowspan="3">NRO2</td>
              <td rowspan="3">NRO3</td>
              <td>NRO4</td>
          </tr>
          <tr>
              <td>Doc1.</td>
          </tr>
          <tr>
              <td>Doc2.</td>
          </tr>
        </table>
      `)

      const parsedXlsx = await new Promise((resolve, reject) => {
        const bufs = []

        stream.on('error', reject)
        stream.on('data', (d) => { bufs.push(d) })

        stream.on('end', () => {
          const buf = Buffer.concat(bufs)
          resolve(xlsx.read(buf))
        })
      })

      should(parsedXlsx.Sheets[parsedXlsx.SheetNames[0]]['A1'].v).be.eql('NRO1')
      should(parsedXlsx.Sheets[parsedXlsx.SheetNames[0]]['A2']).be.undefined()
      should(parsedXlsx.Sheets[parsedXlsx.SheetNames[0]]['A3']).be.undefined()
      should(parsedXlsx.Sheets[parsedXlsx.SheetNames[0]]['B1'].v).be.eql('NRO2')
      should(parsedXlsx.Sheets[parsedXlsx.SheetNames[0]]['B2']).be.undefined()
      should(parsedXlsx.Sheets[parsedXlsx.SheetNames[0]]['B3']).be.undefined()
      should(parsedXlsx.Sheets[parsedXlsx.SheetNames[0]]['C1'].v).be.eql('NRO3')
      should(parsedXlsx.Sheets[parsedXlsx.SheetNames[0]]['C2']).be.undefined()
      should(parsedXlsx.Sheets[parsedXlsx.SheetNames[0]]['C3']).be.undefined()
      should(parsedXlsx.Sheets[parsedXlsx.SheetNames[0]]['D1'].v).be.eql('NRO4')
      should(parsedXlsx.Sheets[parsedXlsx.SheetNames[0]]['D2'].v).be.eql('Doc1.')
      should(parsedXlsx.Sheets[parsedXlsx.SheetNames[0]]['D3'].v).be.eql('Doc2.')

      should(parsedXlsx.Sheets[parsedXlsx.SheetNames[0]]['!merges'][0].s.r).be.eql(0)
      should(parsedXlsx.Sheets[parsedXlsx.SheetNames[0]]['!merges'][0].s.c).be.eql(0)
      should(parsedXlsx.Sheets[parsedXlsx.SheetNames[0]]['!merges'][0].e.r).be.eql(2)
      should(parsedXlsx.Sheets[parsedXlsx.SheetNames[0]]['!merges'][0].e.c).be.eql(0)

      should(parsedXlsx.Sheets[parsedXlsx.SheetNames[0]]['!merges'][1].s.r).be.eql(0)
      should(parsedXlsx.Sheets[parsedXlsx.SheetNames[0]]['!merges'][1].s.c).be.eql(1)
      should(parsedXlsx.Sheets[parsedXlsx.SheetNames[0]]['!merges'][1].e.r).be.eql(2)
      should(parsedXlsx.Sheets[parsedXlsx.SheetNames[0]]['!merges'][1].e.c).be.eql(1)

      should(parsedXlsx.Sheets[parsedXlsx.SheetNames[0]]['!merges'][2].s.r).be.eql(0)
      should(parsedXlsx.Sheets[parsedXlsx.SheetNames[0]]['!merges'][2].s.c).be.eql(2)
      should(parsedXlsx.Sheets[parsedXlsx.SheetNames[0]]['!merges'][2].e.r).be.eql(2)
      should(parsedXlsx.Sheets[parsedXlsx.SheetNames[0]]['!merges'][2].e.c).be.eql(2)
    })

    it('should work when using special rowspan layout #6 (row with lot of cells using rowspan but last one using colspan)', async () => {
      const stream = await conversion(`
        <table>
          <tr>
              <td rowspan="3">NRO1</td>
              <td rowspan="3">Text1</td>
              <td rowspan="3">Text2</td>
              <td colspan="3">Receip</td>
          </tr>
          <tr>
              <td>Doc.</td>
              <td colspan="2">Information</td>
          </tr>
          <tr>
              <td>Text3</td>
          </tr>
        </table>
      `)

      const parsedXlsx = await new Promise((resolve, reject) => {
        const bufs = []

        stream.on('error', reject)
        stream.on('data', (d) => { bufs.push(d) })

        stream.on('end', () => {
          const buf = Buffer.concat(bufs)
          resolve(xlsx.read(buf))
        })
      })

      should(parsedXlsx.Sheets[parsedXlsx.SheetNames[0]]['A1'].v).be.eql('NRO1')
      should(parsedXlsx.Sheets[parsedXlsx.SheetNames[0]]['A2']).be.undefined()
      should(parsedXlsx.Sheets[parsedXlsx.SheetNames[0]]['A3']).be.undefined()
      should(parsedXlsx.Sheets[parsedXlsx.SheetNames[0]]['B1'].v).be.eql('Text1')
      should(parsedXlsx.Sheets[parsedXlsx.SheetNames[0]]['B2']).be.undefined()
      should(parsedXlsx.Sheets[parsedXlsx.SheetNames[0]]['B3']).be.undefined()
      should(parsedXlsx.Sheets[parsedXlsx.SheetNames[0]]['C1'].v).be.eql('Text2')
      should(parsedXlsx.Sheets[parsedXlsx.SheetNames[0]]['C2']).be.undefined()
      should(parsedXlsx.Sheets[parsedXlsx.SheetNames[0]]['C3']).be.undefined()
      should(parsedXlsx.Sheets[parsedXlsx.SheetNames[0]]['D1'].v).be.eql('Receip')
      should(parsedXlsx.Sheets[parsedXlsx.SheetNames[0]]['E1']).be.undefined()
      should(parsedXlsx.Sheets[parsedXlsx.SheetNames[0]]['F1']).be.undefined()
      should(parsedXlsx.Sheets[parsedXlsx.SheetNames[0]]['D2'].v).be.eql('Doc.')
      should(parsedXlsx.Sheets[parsedXlsx.SheetNames[0]]['E2'].v).be.eql('Information')
      should(parsedXlsx.Sheets[parsedXlsx.SheetNames[0]]['F2']).be.undefined()
      should(parsedXlsx.Sheets[parsedXlsx.SheetNames[0]]['D3'].v).be.eql('Text3')

      should(parsedXlsx.Sheets[parsedXlsx.SheetNames[0]]['!merges'][0].s.r).be.eql(0)
      should(parsedXlsx.Sheets[parsedXlsx.SheetNames[0]]['!merges'][0].s.c).be.eql(0)
      should(parsedXlsx.Sheets[parsedXlsx.SheetNames[0]]['!merges'][0].e.r).be.eql(2)
      should(parsedXlsx.Sheets[parsedXlsx.SheetNames[0]]['!merges'][0].e.c).be.eql(0)

      should(parsedXlsx.Sheets[parsedXlsx.SheetNames[0]]['!merges'][1].s.r).be.eql(0)
      should(parsedXlsx.Sheets[parsedXlsx.SheetNames[0]]['!merges'][1].s.c).be.eql(1)
      should(parsedXlsx.Sheets[parsedXlsx.SheetNames[0]]['!merges'][1].e.r).be.eql(2)
      should(parsedXlsx.Sheets[parsedXlsx.SheetNames[0]]['!merges'][1].e.c).be.eql(1)

      should(parsedXlsx.Sheets[parsedXlsx.SheetNames[0]]['!merges'][2].s.r).be.eql(0)
      should(parsedXlsx.Sheets[parsedXlsx.SheetNames[0]]['!merges'][2].s.c).be.eql(2)
      should(parsedXlsx.Sheets[parsedXlsx.SheetNames[0]]['!merges'][2].e.r).be.eql(2)
      should(parsedXlsx.Sheets[parsedXlsx.SheetNames[0]]['!merges'][2].e.c).be.eql(2)

      should(parsedXlsx.Sheets[parsedXlsx.SheetNames[0]]['!merges'][3].s.r).be.eql(0)
      should(parsedXlsx.Sheets[parsedXlsx.SheetNames[0]]['!merges'][3].s.c).be.eql(3)
      should(parsedXlsx.Sheets[parsedXlsx.SheetNames[0]]['!merges'][3].e.r).be.eql(0)
      should(parsedXlsx.Sheets[parsedXlsx.SheetNames[0]]['!merges'][3].e.c).be.eql(5)

      should(parsedXlsx.Sheets[parsedXlsx.SheetNames[0]]['!merges'][4].s.r).be.eql(1)
      should(parsedXlsx.Sheets[parsedXlsx.SheetNames[0]]['!merges'][4].s.c).be.eql(4)
      should(parsedXlsx.Sheets[parsedXlsx.SheetNames[0]]['!merges'][4].e.r).be.eql(1)
      should(parsedXlsx.Sheets[parsedXlsx.SheetNames[0]]['!merges'][4].e.c).be.eql(5)
    })

    it('should work when using special rowspan layout #7 (row with only one cell that uses rowspan greater than available rows)', async () => {
      const stream = await conversion(`
        <table>
          <tr>
              <td rowspan="3">ROWSPAN 3</td>
          </tr>
        </table>
      `)

      const parsedXlsx = await new Promise((resolve, reject) => {
        const bufs = []

        stream.on('error', reject)
        stream.on('data', (d) => { bufs.push(d) })

        stream.on('end', () => {
          const buf = Buffer.concat(bufs)
          resolve(xlsx.read(buf))
        })
      })

      should(parsedXlsx.Sheets[parsedXlsx.SheetNames[0]]['A1'].v).be.eql('ROWSPAN 3')
      should(parsedXlsx.Sheets[parsedXlsx.SheetNames[0]]['A2']).be.undefined()

      if (!legacy) {
        should(parsedXlsx.Sheets[parsedXlsx.SheetNames[0]]['!merges']).be.undefined()
      }
    })

    it('should work when using special rowspan layout #8 (complex calendar like layout)', async () => {
      const stream = await conversion(`
        <table border="1" style="border-collapse:collapse;">
          <tr>
            <td rowspan="2" colspan="2">corner</td>
            <td colspan="5">2015</td>
            <td colspan="5">2016</td>
            <td colspan="5">Summary</td>
          </tr>
          <tr>
            <td>Amount 1</td>
            <td>Amount 2</td>
            <td>Amount 3</td>
            <td>Amount 4</td>
            <td>Amount 5</td>
            <td>Amount 1</td>
            <td>Amount 2</td>
            <td>Amount 3</td>
            <td>Amount 4</td>
            <td>Amount 5</td>
            <td>Total Amount 1</td>
            <td>Total Amount 2</td>
            <td>Total Amount 3</td>
            <td>Total Amount 4</td>
            <td>Total Amount 5</td>
          </tr>
          <tr>
            <td rowspan="2" >Buffer</td>
            <td>Jane Doe</td>
            <td>10</td>
            <td>15</td>
            <td>20</td>
            <td>25</td>
            <td>30</td>
            <td>2</td>
            <td>4</td>
            <td>6</td>
            <td>8</td>
            <td>10</td>
            <td>12</td>
            <td>19</td>
            <td>26</td>
            <td>32</td>
            <td>40</td>
          </tr>
          <tr>
            <td>Thomas Smith</td>
            <td>0</td>
            <td>25</td>
            <td>20</td>
            <td>15</td>
            <td>10</td>
            <td>5</td>
            <td>3</td>
            <td>6</td>
            <td>9</td>
            <td>12</td>
            <td>15</td>
            <td>5</td>
            <td>28</td>
            <td>26</td>
            <td>22</td>
          </tr>
        </table>
      `)

      const parsedXlsx = await new Promise((resolve, reject) => {
        const bufs = []

        stream.on('error', reject)
        stream.on('data', (d) => { bufs.push(d) })

        stream.on('end', () => {
          const buf = Buffer.concat(bufs)
          resolve(xlsx.read(buf))
        })
      })

      should(parsedXlsx.Sheets[parsedXlsx.SheetNames[0]]['A1'].v).be.eql('corner')
      should(parsedXlsx.Sheets[parsedXlsx.SheetNames[0]]['B1']).be.undefined()
      should(parsedXlsx.Sheets[parsedXlsx.SheetNames[0]]['A2']).be.undefined()
      should(parsedXlsx.Sheets[parsedXlsx.SheetNames[0]]['B2']).be.undefined()
      should(parsedXlsx.Sheets[parsedXlsx.SheetNames[0]]['C1'].v).be.eql('2015')
      should(parsedXlsx.Sheets[parsedXlsx.SheetNames[0]]['D1']).be.undefined()
      should(parsedXlsx.Sheets[parsedXlsx.SheetNames[0]]['E1']).be.undefined()
      should(parsedXlsx.Sheets[parsedXlsx.SheetNames[0]]['F1']).be.undefined()
      should(parsedXlsx.Sheets[parsedXlsx.SheetNames[0]]['G1']).be.undefined()
      should(parsedXlsx.Sheets[parsedXlsx.SheetNames[0]]['H1'].v).be.eql('2016')
      should(parsedXlsx.Sheets[parsedXlsx.SheetNames[0]]['I1']).be.undefined()
      should(parsedXlsx.Sheets[parsedXlsx.SheetNames[0]]['J1']).be.undefined()
      should(parsedXlsx.Sheets[parsedXlsx.SheetNames[0]]['K1']).be.undefined()
      should(parsedXlsx.Sheets[parsedXlsx.SheetNames[0]]['L1']).be.undefined()
      should(parsedXlsx.Sheets[parsedXlsx.SheetNames[0]]['M1'].v).be.eql('Summary')
      should(parsedXlsx.Sheets[parsedXlsx.SheetNames[0]]['N1']).be.undefined()
      should(parsedXlsx.Sheets[parsedXlsx.SheetNames[0]]['O1']).be.undefined()
      should(parsedXlsx.Sheets[parsedXlsx.SheetNames[0]]['P1']).be.undefined()
      should(parsedXlsx.Sheets[parsedXlsx.SheetNames[0]]['Q1']).be.undefined()
      should(parsedXlsx.Sheets[parsedXlsx.SheetNames[0]]['C2'].v).be.eql('Amount 1')
      should(parsedXlsx.Sheets[parsedXlsx.SheetNames[0]]['D2'].v).be.eql('Amount 2')
      should(parsedXlsx.Sheets[parsedXlsx.SheetNames[0]]['E2'].v).be.eql('Amount 3')
      should(parsedXlsx.Sheets[parsedXlsx.SheetNames[0]]['F2'].v).be.eql('Amount 4')
      should(parsedXlsx.Sheets[parsedXlsx.SheetNames[0]]['G2'].v).be.eql('Amount 5')
      should(parsedXlsx.Sheets[parsedXlsx.SheetNames[0]]['H2'].v).be.eql('Amount 1')
      should(parsedXlsx.Sheets[parsedXlsx.SheetNames[0]]['I2'].v).be.eql('Amount 2')
      should(parsedXlsx.Sheets[parsedXlsx.SheetNames[0]]['J2'].v).be.eql('Amount 3')
      should(parsedXlsx.Sheets[parsedXlsx.SheetNames[0]]['K2'].v).be.eql('Amount 4')
      should(parsedXlsx.Sheets[parsedXlsx.SheetNames[0]]['L2'].v).be.eql('Amount 5')
      should(parsedXlsx.Sheets[parsedXlsx.SheetNames[0]]['M2'].v).be.eql('Total Amount 1')
      should(parsedXlsx.Sheets[parsedXlsx.SheetNames[0]]['N2'].v).be.eql('Total Amount 2')
      should(parsedXlsx.Sheets[parsedXlsx.SheetNames[0]]['O2'].v).be.eql('Total Amount 3')
      should(parsedXlsx.Sheets[parsedXlsx.SheetNames[0]]['P2'].v).be.eql('Total Amount 4')
      should(parsedXlsx.Sheets[parsedXlsx.SheetNames[0]]['Q2'].v).be.eql('Total Amount 5')
      should(parsedXlsx.Sheets[parsedXlsx.SheetNames[0]]['A3'].v).be.eql('Buffer')
      should(parsedXlsx.Sheets[parsedXlsx.SheetNames[0]]['A4']).be.undefined()
      should(parsedXlsx.Sheets[parsedXlsx.SheetNames[0]]['B3'].v).be.eql('Jane Doe')
      should(parsedXlsx.Sheets[parsedXlsx.SheetNames[0]]['C3'].v).be.eql('10')
      should(parsedXlsx.Sheets[parsedXlsx.SheetNames[0]]['D3'].v).be.eql('15')
      should(parsedXlsx.Sheets[parsedXlsx.SheetNames[0]]['E3'].v).be.eql('20')
      should(parsedXlsx.Sheets[parsedXlsx.SheetNames[0]]['F3'].v).be.eql('25')
      should(parsedXlsx.Sheets[parsedXlsx.SheetNames[0]]['G3'].v).be.eql('30')
      should(parsedXlsx.Sheets[parsedXlsx.SheetNames[0]]['H3'].v).be.eql('2')
      should(parsedXlsx.Sheets[parsedXlsx.SheetNames[0]]['I3'].v).be.eql('4')
      should(parsedXlsx.Sheets[parsedXlsx.SheetNames[0]]['J3'].v).be.eql('6')
      should(parsedXlsx.Sheets[parsedXlsx.SheetNames[0]]['K3'].v).be.eql('8')
      should(parsedXlsx.Sheets[parsedXlsx.SheetNames[0]]['L3'].v).be.eql('10')
      should(parsedXlsx.Sheets[parsedXlsx.SheetNames[0]]['M3'].v).be.eql('12')
      should(parsedXlsx.Sheets[parsedXlsx.SheetNames[0]]['N3'].v).be.eql('19')
      should(parsedXlsx.Sheets[parsedXlsx.SheetNames[0]]['O3'].v).be.eql('26')
      should(parsedXlsx.Sheets[parsedXlsx.SheetNames[0]]['P3'].v).be.eql('32')
      should(parsedXlsx.Sheets[parsedXlsx.SheetNames[0]]['Q3'].v).be.eql('40')
      should(parsedXlsx.Sheets[parsedXlsx.SheetNames[0]]['B4'].v).be.eql('Thomas Smith')
      should(parsedXlsx.Sheets[parsedXlsx.SheetNames[0]]['C4'].v).be.eql('0')
      should(parsedXlsx.Sheets[parsedXlsx.SheetNames[0]]['D4'].v).be.eql('25')
      should(parsedXlsx.Sheets[parsedXlsx.SheetNames[0]]['E4'].v).be.eql('20')
      should(parsedXlsx.Sheets[parsedXlsx.SheetNames[0]]['F4'].v).be.eql('15')
      should(parsedXlsx.Sheets[parsedXlsx.SheetNames[0]]['G4'].v).be.eql('10')
      should(parsedXlsx.Sheets[parsedXlsx.SheetNames[0]]['H4'].v).be.eql('5')
      should(parsedXlsx.Sheets[parsedXlsx.SheetNames[0]]['I4'].v).be.eql('3')
      should(parsedXlsx.Sheets[parsedXlsx.SheetNames[0]]['J4'].v).be.eql('6')
      should(parsedXlsx.Sheets[parsedXlsx.SheetNames[0]]['K4'].v).be.eql('9')
      should(parsedXlsx.Sheets[parsedXlsx.SheetNames[0]]['L4'].v).be.eql('12')
      should(parsedXlsx.Sheets[parsedXlsx.SheetNames[0]]['M4'].v).be.eql('15')
      should(parsedXlsx.Sheets[parsedXlsx.SheetNames[0]]['N4'].v).be.eql('5')
      should(parsedXlsx.Sheets[parsedXlsx.SheetNames[0]]['O4'].v).be.eql('28')
      should(parsedXlsx.Sheets[parsedXlsx.SheetNames[0]]['P4'].v).be.eql('26')
      should(parsedXlsx.Sheets[parsedXlsx.SheetNames[0]]['Q4'].v).be.eql('22')

      should(parsedXlsx.Sheets[parsedXlsx.SheetNames[0]]['!merges'][0].s.r).be.eql(0)
      should(parsedXlsx.Sheets[parsedXlsx.SheetNames[0]]['!merges'][0].s.c).be.eql(0)
      should(parsedXlsx.Sheets[parsedXlsx.SheetNames[0]]['!merges'][0].e.r).be.eql(1)
      should(parsedXlsx.Sheets[parsedXlsx.SheetNames[0]]['!merges'][0].e.c).be.eql(1)

      should(parsedXlsx.Sheets[parsedXlsx.SheetNames[0]]['!merges'][1].s.r).be.eql(0)
      should(parsedXlsx.Sheets[parsedXlsx.SheetNames[0]]['!merges'][1].s.c).be.eql(2)
      should(parsedXlsx.Sheets[parsedXlsx.SheetNames[0]]['!merges'][1].e.r).be.eql(0)
      should(parsedXlsx.Sheets[parsedXlsx.SheetNames[0]]['!merges'][1].e.c).be.eql(6)

      should(parsedXlsx.Sheets[parsedXlsx.SheetNames[0]]['!merges'][2].s.r).be.eql(0)
      should(parsedXlsx.Sheets[parsedXlsx.SheetNames[0]]['!merges'][2].s.c).be.eql(7)
      should(parsedXlsx.Sheets[parsedXlsx.SheetNames[0]]['!merges'][2].e.r).be.eql(0)
      should(parsedXlsx.Sheets[parsedXlsx.SheetNames[0]]['!merges'][2].e.c).be.eql(11)

      should(parsedXlsx.Sheets[parsedXlsx.SheetNames[0]]['!merges'][3].s.r).be.eql(0)
      should(parsedXlsx.Sheets[parsedXlsx.SheetNames[0]]['!merges'][3].s.c).be.eql(12)
      should(parsedXlsx.Sheets[parsedXlsx.SheetNames[0]]['!merges'][3].e.r).be.eql(0)
      should(parsedXlsx.Sheets[parsedXlsx.SheetNames[0]]['!merges'][3].e.c).be.eql(16)

      should(parsedXlsx.Sheets[parsedXlsx.SheetNames[0]]['!merges'][4].s.r).be.eql(2)
      should(parsedXlsx.Sheets[parsedXlsx.SheetNames[0]]['!merges'][4].s.c).be.eql(0)
      should(parsedXlsx.Sheets[parsedXlsx.SheetNames[0]]['!merges'][4].e.r).be.eql(3)
      should(parsedXlsx.Sheets[parsedXlsx.SheetNames[0]]['!merges'][4].e.c).be.eql(0)
    })

    if (!legacy) {
      it('should work when using special rowspan layout #9 (using rowspan in different rows)', async () => {
        const stream = await conversion(`
          <table>
            <tr>
              <td rowspan='3'>Row 1 Col 1</td>
              <td>Row 1 Col 2</td>
              <td>Row 1 Col 3</td>
              <td>Row 1 Col 4</td>
            </tr>
            <tr>
              <td rowspan='2'>Row 2 Col 1</td>
              <td rowspan='2'>Row 2 Col 2</td>
              <td>Row 2 Col 3</td>
            </tr>
            <tr>
              <td>Row 3 Col 3</td>
            </tr>
          </table>
        `)

        const parsedXlsx = await new Promise((resolve, reject) => {
          const bufs = []

          stream.on('error', reject)
          stream.on('data', (d) => { bufs.push(d) })

          stream.on('end', () => {
            const buf = Buffer.concat(bufs)
            resolve(xlsx.read(buf))
          })
        })

        should(parsedXlsx.Sheets[parsedXlsx.SheetNames[0]]['A1'].v).be.eql('Row 1 Col 1')
        should(parsedXlsx.Sheets[parsedXlsx.SheetNames[0]]['A2']).be.undefined()
        should(parsedXlsx.Sheets[parsedXlsx.SheetNames[0]]['A3']).be.undefined()
        should(parsedXlsx.Sheets[parsedXlsx.SheetNames[0]]['B1'].v).be.eql('Row 1 Col 2')
        should(parsedXlsx.Sheets[parsedXlsx.SheetNames[0]]['B2'].v).be.eql('Row 2 Col 1')
        should(parsedXlsx.Sheets[parsedXlsx.SheetNames[0]]['B3']).be.undefined()
        should(parsedXlsx.Sheets[parsedXlsx.SheetNames[0]]['C1'].v).be.eql('Row 1 Col 3')
        should(parsedXlsx.Sheets[parsedXlsx.SheetNames[0]]['C2'].v).be.eql('Row 2 Col 2')
        should(parsedXlsx.Sheets[parsedXlsx.SheetNames[0]]['C3']).be.undefined()
        should(parsedXlsx.Sheets[parsedXlsx.SheetNames[0]]['D1'].v).be.eql('Row 1 Col 4')
        should(parsedXlsx.Sheets[parsedXlsx.SheetNames[0]]['D2'].v).be.eql('Row 2 Col 3')
        should(parsedXlsx.Sheets[parsedXlsx.SheetNames[0]]['D3'].v).be.eql('Row 3 Col 3')

        should(parsedXlsx.Sheets[parsedXlsx.SheetNames[0]]['!merges'][0].s.r).be.eql(0)
        should(parsedXlsx.Sheets[parsedXlsx.SheetNames[0]]['!merges'][0].s.c).be.eql(0)
        should(parsedXlsx.Sheets[parsedXlsx.SheetNames[0]]['!merges'][0].e.r).be.eql(2)
        should(parsedXlsx.Sheets[parsedXlsx.SheetNames[0]]['!merges'][0].e.c).be.eql(0)

        should(parsedXlsx.Sheets[parsedXlsx.SheetNames[0]]['!merges'][1].s.r).be.eql(1)
        should(parsedXlsx.Sheets[parsedXlsx.SheetNames[0]]['!merges'][1].s.c).be.eql(1)
        should(parsedXlsx.Sheets[parsedXlsx.SheetNames[0]]['!merges'][1].e.r).be.eql(2)
        should(parsedXlsx.Sheets[parsedXlsx.SheetNames[0]]['!merges'][1].e.c).be.eql(1)

        should(parsedXlsx.Sheets[parsedXlsx.SheetNames[0]]['!merges'][1].s.r).be.eql(1)
        should(parsedXlsx.Sheets[parsedXlsx.SheetNames[0]]['!merges'][1].s.c).be.eql(1)
        should(parsedXlsx.Sheets[parsedXlsx.SheetNames[0]]['!merges'][1].e.r).be.eql(2)
        should(parsedXlsx.Sheets[parsedXlsx.SheetNames[0]]['!merges'][1].e.c).be.eql(1)
      })

      it('should be able to set fontFamily', async () => {
        const stream = await conversion(`
          <table>
            <tr>
              <td>Hello</td>
            </tr>
          </table>
        `, {
          fontFamily: 'Verdana'
        })

        const parsedXlsx = await new Promise((resolve, reject) => {
          const bufs = []

          stream.on('error', reject)
          stream.on('data', (d) => { bufs.push(d) })

          stream.on('end', () => {
            const buf = Buffer.concat(bufs)
            resolve(xlsx.read(buf))
          })
        })

        should(parsedXlsx.Styles.Fonts[0].name).be.eql('Verdana')
      })

      it('should wait for JS trigger', async () => {
        const stream = await conversion(`
          <table id="main">
          </table>
          <script>
            setTimeout(function () {
              var table = document.getElementById('main')
              var row = document.createElement('tr')
              var cell = document.createElement('td')

              cell.innerHTML = 'Hello'
              row.appendChild(cell)
              table.appendChild(row)

              window.CHROME_PAGE_EVAL_READY = true
              window.PHANTOM_PAGE_EVAL_READY = true
            }, 500)
          </script>
        `, {
          waitForJS: true
        })

        const parsedXlsx = await new Promise((resolve, reject) => {
          const bufs = []

          stream.on('error', reject)
          stream.on('data', (d) => { bufs.push(d) })

          stream.on('end', () => {
            const buf = Buffer.concat(bufs)
            resolve(xlsx.read(buf))
          })
        })

        should(parsedXlsx.Sheets[parsedXlsx.SheetNames[0]]['A1'].v).be.eql('Hello')
      })

      it('should wait for JS trigger (custom var name)', async () => {
        const stream = await conversion(`
          <table id="main">
          </table>
          <script>
            setTimeout(function () {
              var table = document.getElementById('main')
              var row = document.createElement('tr')
              var cell = document.createElement('td')

              cell.innerHTML = 'Hello'
              row.appendChild(cell)
              table.appendChild(row)

              window.READY_TO_START = true
            }, 500)
          </script>
        `, {
          waitForJS: true,
          waitForJSVarName: 'READY_TO_START'
        })

        const parsedXlsx = await new Promise((resolve, reject) => {
          const bufs = []

          stream.on('error', reject)
          stream.on('data', (d) => { bufs.push(d) })

          stream.on('end', () => {
            const buf = Buffer.concat(bufs)
            resolve(xlsx.read(buf))
          })
        })

        should(parsedXlsx.Sheets[parsedXlsx.SheetNames[0]]['A1'].v).be.eql('Hello')
      })

      it('should generate multiple sheets when there are multiple tables in html', async () => {
        const stream = await conversion(`
          <table name="Data1">
            <tr>
              <td>1</td>
              <td>2</td>
              <td>3</td>
              <td>4</td>
            </tr>
          </table>
          <table name="Data2">
            <tr>
              <td>1</td>
              <td>2</td>
              <td>3</td>
              <td>4</td>
            </tr>
          </table>
        `)

        const parsedXlsx = await new Promise((resolve, reject) => {
          const bufs = []

          stream.on('error', reject)
          stream.on('data', (d) => { bufs.push(d) })

          stream.on('end', () => {
            const buf = Buffer.concat(bufs)
            resolve(xlsx.read(buf))
          })
        })

        should(parsedXlsx.Sheets[parsedXlsx.SheetNames[0]]['A1'].v).be.eql('1')
        should(parsedXlsx.Sheets[parsedXlsx.SheetNames[0]]['B1'].v).be.eql('2')
        should(parsedXlsx.Sheets[parsedXlsx.SheetNames[0]]['C1'].v).be.eql('3')
        should(parsedXlsx.Sheets[parsedXlsx.SheetNames[0]]['D1'].v).be.eql('4')

        should(parsedXlsx.Sheets[parsedXlsx.SheetNames[1]]['A1'].v).be.eql('1')
        should(parsedXlsx.Sheets[parsedXlsx.SheetNames[1]]['B1'].v).be.eql('2')
        should(parsedXlsx.Sheets[parsedXlsx.SheetNames[1]]['C1'].v).be.eql('3')
        should(parsedXlsx.Sheets[parsedXlsx.SheetNames[1]]['D1'].v).be.eql('4')
      })
    }

    it('should callback error when input contains invalid characters', async () => {
      return (
        conversion(`
          <table>
            <tr>
              <td></td>
            </tr>
          </table>
        `)
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
      const stream = await conversion(`
        <table>
          <tr>
            <td>& &</td>
          </tr>
        </table>
      `)

      const bufs = []

      await new Promise((resolve, reject) => {
        stream.on('data', (d) => { bufs.push(d) })
        stream.on('error', reject)
        stream.on('end', () => {
          try {
            const buf = Buffer.concat(bufs)
            xlsx.read(buf).Strings[0].t.should.be.eql('& &')
            resolve()
          } catch (e) {
            reject(e)
          }
        })
      })
    })

    it('should callback error when row doesn\'t contain cells', async () => {
      return (
        conversion('<table><tr>Hello</tr></table>')
      ).should.be.rejected()
    })
  }
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
