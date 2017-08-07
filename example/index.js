'use strict'
const ExcelParser = require('../lib').xlsxParser
const fs = require('fs')
const co = require('co')

const filePath = './test/assets/excel_template.xlsx'
let data = fs.readFileSync(filePath)

const excel = new ExcelParser(data, ['a', 'b'], 3)

// use generator
function * parser () {
  let row = yield excel.parse()
  console.log(row)
  let ret = excel.toArray()
  console.log(ret)
}
co(parser());

// use async
(async function () {
  let row = await excel.parse()
  console.log(row)
  let ret = excel.toArray()
  console.log(ret)
})()
