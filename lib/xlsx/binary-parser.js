'use strict'

const Excel = require('exceljs')
const _ = require('lodash')

function cellToString (cell) {
  if (_.isDate(cell)) {
    return cell.toISOString()
  } else if (_.isBoolean(cell)) {
    return `${cell}`
  } else if (_.isObject(cell) && _.isArray(cell.richText)) {
    return cell.richText.map((ele) => {
      return cellToString(ele.text)
    }).join('')
  } else if (_.isObject(cell) && _.isObject(cell.text)) {
    return cellToString(cell.text)
  } else if (_.isObject(cell) && _.isString(cell.text)) {
    return `${cell.text}`
  } else if (_.isString(cell)) {
    return cell
  } else if (_.isNumber(cell)) {
    return `${cell}`
  } else if (_.isUndefined(cell) || _.isNull(cell)) {
    return ''
  } else {
    throw new Error(`unknown type ${cell.constructor.name}`)
  }
}

class ExcelParser {
  constructor (data, rule, limit) {
    this.rule = rule
    this.data = data
    this.limit = limit
    this.headerline = 2
    this.rows = []
    this.hasParsed = false
    this.hastoArray = false
    this.jsonArray = []
  }

  setHeaderline (headerline) {
    this.headerline = headerline
  }

  get isExccedLimit () {
    return this.rows.length > this.limit
  }

  * parse () {
    if (this.hasParsed) return this.rows
    let workbook = new Excel.Workbook()
    let worksheet = (yield workbook.xlsx.load(this.data)).getWorksheet(1)
    this.rows = worksheet.getSheetValues()

    // the first row include(A,B,C,D...), its useless
    this.rows = _.drop(this.rows, this.headerline + 1)
    this.hasParsed = true
    this.rows = this.rows.map((row) => {
      return row.map((cell) => {
        return cellToString(cell)
      })
    })
    return this.rows
  }

  toArray () {
    if (!this.hasParsed) throw new Error('need parse before')
    if (this.hastoArray) return this.jsonArray
    for (let index = 0; index < this.rows.length; index++) {
      // 根据excel模板 从第三行开始拿数据
      let row = this.rows[index]
      // Remove the serial number of the first column of each row
      row = _.drop(row, 1)
      let _data = {}
      for (let index = 0; index < this.rule.length; index++) {
        const rule = this.rule[index]
        _data[rule] = row[index] || ''
      }
      if (_.compact(_.values(_data)).length !== 0) this.jsonArray.push(_data)
    }
    this.hastoArray = true
    return this.jsonArray
  }
}

module.exports = ExcelParser
