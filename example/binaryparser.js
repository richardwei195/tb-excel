'use strict'
const ExcelParser = require('../lib').xlsxParser
const co = require('co')
const request = require('request')

// use generator
function * parser () {
  let res = yield request.get('xxxx.xlsx', { encoding: null })
  let data = res.body
  let excel = new ExcelParser(data, ['a', 'b'], 3)
  let row = yield excel.parse()
  console.log(row)
  let ret = excel.toArray()
  console.log(ret)
}
co(parser());

// use async
(async function () {
  let res = await request.get('xxxx.xlsx', { encoding: null })
  let data = res.body
  let excel = new ExcelParser(data, ['a', 'b'], 3)
  let row = await excel.parse()
  console.log(row)
  let ret = excel.toArray()
  console.log(ret)
})()
