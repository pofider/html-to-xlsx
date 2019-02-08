'use strict'

const util = require('util')
const path = require('path')
const fs = require('fs')
const uuid = require('uuid/v4')
const tmpDir = require('os').tmpdir()
const XlsxPopulate = require('xlsx-populate')
const stylesMap = require('./stylesMap')
const tableToXlsx = require('./tableToXlsx')
const readFileAsync = util.promisify(fs.readFile)
const writeFileAsync = util.promisify(fs.writeFile)

module.exports = (opt = {}) => {
  const options = { ...opt }

  options.timeout = options.timeout || 10000
  options.tmpDir = options.tmpDir || tmpDir

  if (typeof options.extract !== 'function') {
    throw new Error('`extract` option must be a function')
  }

  if (typeof options.timeout !== 'number') {
    throw new Error('`timeout` option must be a number')
  }

  const timeout = options.timeout
  const currentExtractFn = options.extract

  async function convert (html, convertOptions = {}) {
    const id = uuid()

    if (html == null) {
      throw new Error('required `html` option not specified')
    }

    const htmlPath = path.join(options.tmpDir, `${id}-html-to-xlsx.html`)
    const scriptFnPath = options.conversionScriptPath != null ? options.conversionScriptPath : path.join(__dirname, 'scripts/conversionScript.js')

    let scriptFn = await readFileAsync(scriptFnPath)

    scriptFn = scriptFn.toString()

    await writeFileAsync(htmlPath, html)

    const tables = await currentExtractFn({
      ...convertOptions,
      html: htmlPath,
      scriptFn,
      timeout
    })

    let stream

    if (!tables || (Array.isArray(tables) && tables.length === 0)) {
      throw new Error('No table element(s) found in html')
    }

    stream = await tableToXlsx(options, tables, id)

    return stream
  }

  return convert
}

module.exports.getXlsxStyleNames = () => {
  return Object.keys(stylesMap)
}

module.exports.XlsxPopulate = XlsxPopulate
