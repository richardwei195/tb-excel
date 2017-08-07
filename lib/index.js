'use strict'

const xlsxParser = require('./xlsx/binary-parser')
const exceljs = require('exceljs')

module.exports.xlsxParser = xlsxParser
module.exports = exceljs
