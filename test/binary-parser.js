'use strict'
/* eslint-env mocha */

const expect = require('should')
const fs = require('fs')
const config = require('./config')
const rewire = require('rewire')
const ExcelParser = rewire('../lib/xlsx/binary-parser')

const filePath = './test/assets/excel_template.xlsx'

describe('ExcelParser', function () {
  describe('limit', function () {
    it('should success when isntExccedLimit', function * () {
      let data = fs.readFileSync(filePath)
      let parser = new ExcelParser(data, config.IMPORT_EXCEL.TASK.RULE, config.IMPORT_EXCEL.TASK.LIMIT)
      yield parser.parse()
      parser.isExccedLimit.should.eql(false)
    })

    it('should success when isExccedLimit', function * () {
      let data = fs.readFileSync(filePath)
      let parser = new ExcelParser(data, config.IMPORT_EXCEL.TASK.RULE, config.IMPORT_EXCEL.TASK.LIMIT)
      parser.limit = 0
      yield parser.parse()
      parser.isExccedLimit.should.eql(true)
    })
  })

  describe('should parse failed when wrong constructor', function () {
    it('lack params', function * () {
      fs.readFileSync('./test/assets/test.csv')
      let parser = new ExcelParser(config.IMPORT_EXCEL.TASK.RULE, config.IMPORT_EXCEL.TASK.LIMIT)

      try {
        yield parser.parse()
      } catch (error) {
        return
      }
      throw new Error('not Thrown')
    })

    it('wrong params order', function * () {
      let data = fs.readFileSync('./test/assets/test.csv')
      let parser = new ExcelParser(config.IMPORT_EXCEL.TASK.RULE, config.IMPORT_EXCEL.TASK.LIMIT, data)

      try {
        yield parser.parse()
      } catch (error) {
        return
      }
      throw new Error('not Thrown')
    })
  })

  describe('parse', function () {
    it('should failed when wrong file', function * () {
      let data = fs.readFileSync('./test/assets/test.csv')
      let parser = new ExcelParser(data, config.IMPORT_EXCEL.TASK.RULE, config.IMPORT_EXCEL.TASK.LIMIT)
      try {
        yield parser.parse()
      } catch (error) {
        return
      }
      throw new Error('not Thrown')
    })

    it('should success when right file', function * () {
      let data = fs.readFileSync(filePath)
      let parser = new ExcelParser(data, config.IMPORT_EXCEL.TASK.RULE, config.IMPORT_EXCEL.TASK.LIMIT)
      parser.setHeaderline(0)
      expect(yield parser.parse()).be.Array()
    })

    it('should parse to String if cell is type of String', function * () {
      let data = fs.readFileSync(filePath)
      let parser = new ExcelParser(data, config.IMPORT_EXCEL.TASK.RULE, config.IMPORT_EXCEL.TASK.LIMIT)
      parser.setHeaderline(0)
      let rows = yield parser.parse()

      expect(rows).be.Array()
      expect(rows[0][1]).equal('正常')
      expect(rows[0][2]).equal('1')
      expect(rows[0][3]).equal('[123]')
      expect(rows[0][4]).be.String('"123"../$%^')
    })

    it('should parse to String if cell is type of Hyperlink or richText', function * () {
      let data = fs.readFileSync(filePath)
      let parser = new ExcelParser(data, config.IMPORT_EXCEL.TASK.RULE, config.IMPORT_EXCEL.TASK.LIMIT)
      parser.setHeaderline(0)
      let rows = yield parser.parse()

      expect(rows).be.Array()
      expect(rows[1][2]).equal('http://www.baidu.com/')
      expect(rows[1][3]).equal('工作表 1\'!A1')
    })

    it('should parse to String if value is type of number ', function * () {
      let data = fs.readFileSync(filePath)
      let parser = new ExcelParser(data, config.IMPORT_EXCEL.TASK.RULE, config.IMPORT_EXCEL.TASK.LIMIT)
      parser.setHeaderline(0)
      let rows = yield parser.parse()

      expect(rows).be.Array()
      expect(rows[2][2]).equal('数值123')
    })

    it('should parse to String if value is type of money ', function * () {
      let data = fs.readFileSync(filePath)
      let parser = new ExcelParser(data, config.IMPORT_EXCEL.TASK.RULE, config.IMPORT_EXCEL.TASK.LIMIT)
      parser.setHeaderline(0)
      let rows = yield parser.parse()

      expect(rows).be.Array()
      expect(rows[2][3]).equal('123')
    })

    it('should parse to String if value is type of accounting ', function * () {
      let data = fs.readFileSync(filePath)
      let parser = new ExcelParser(data, config.IMPORT_EXCEL.TASK.RULE, config.IMPORT_EXCEL.TASK.LIMIT)
      parser.setHeaderline(0)
      let rows = yield parser.parse()

      expect(rows).be.Array()
      expect(rows[2][4]).equal('123')
    })

    it('should parse to String if value is type of date ', function * () {
      let data = fs.readFileSync(filePath)
      let parser = new ExcelParser(data, config.IMPORT_EXCEL.TASK.RULE, config.IMPORT_EXCEL.TASK.LIMIT)
      parser.setHeaderline(0)
      let rows = yield parser.parse()

      expect(rows).be.Array()
      expect(rows[2][5]).equal('1904-01-13T00:00:00.000Z')
    })

    it('should parse to String if value is type of time ', function * () {
      let data = fs.readFileSync(filePath)
      let parser = new ExcelParser(data, config.IMPORT_EXCEL.TASK.RULE, config.IMPORT_EXCEL.TASK.LIMIT)
      parser.setHeaderline(0)
      let rows = yield parser.parse()

      expect(rows).be.Array()
      expect(rows[2][6]).equal('1904-01-13T00:00:00.000Z')
    })

    it('should parse to String if value is type of percent ', function * () {
      let data = fs.readFileSync(filePath)
      let parser = new ExcelParser(data, config.IMPORT_EXCEL.TASK.RULE, config.IMPORT_EXCEL.TASK.LIMIT)
      parser.setHeaderline(0)
      let rows = yield parser.parse()

      expect(rows).be.Array()
      expect(rows[2][7]).equal('0.11')
    })

    it('should parse to String if value is type of fraction ', function * () {
      let data = fs.readFileSync(filePath)
      let parser = new ExcelParser(data, config.IMPORT_EXCEL.TASK.RULE, config.IMPORT_EXCEL.TASK.LIMIT)
      parser.setHeaderline(0)
      let rows = yield parser.parse()

      expect(rows).be.Array()
      expect(rows[2][8]).equal('0.5')
    })

    it('should parse to String if value is type of Scientific count ', function * () {
      let data = fs.readFileSync(filePath)
      let parser = new ExcelParser(data, config.IMPORT_EXCEL.TASK.RULE, config.IMPORT_EXCEL.TASK.LIMIT)
      parser.setHeaderline(0)
      let rows = yield parser.parse()

      expect(rows).be.Array()
      expect(rows[2][9]).equal('323')
    })

    it('should parse to String if value is type of special ', function * () {
      let data = fs.readFileSync(filePath)
      let parser = new ExcelParser(data, config.IMPORT_EXCEL.TASK.RULE, config.IMPORT_EXCEL.TASK.LIMIT)
      parser.setHeaderline(0)
      let rows = yield parser.parse()

      expect(rows).be.Array()
      expect(rows[2][10]).equal('1')
    })

    it('should parse to String if value is type of custom ', function * () {
      let data = fs.readFileSync(filePath)
      let parser = new ExcelParser(data, config.IMPORT_EXCEL.TASK.RULE, config.IMPORT_EXCEL.TASK.LIMIT)
      parser.setHeaderline(0)
      let rows = yield parser.parse()

      expect(rows).be.Array()
      expect(rows[2][11]).equal('1904-01-13T00:00:00.000Z')
    })

    it('should parse to String if value is type of boolean in Number', function * () {
      let data = fs.readFileSync(filePath)
      let parser = new ExcelParser(data, config.IMPORT_EXCEL.TASK.RULE, config.IMPORT_EXCEL.TASK.LIMIT)
      parser.setHeaderline(0)
      let rows = yield parser.parse()

      expect(rows).be.Array()
      expect(rows[2][12]).equal('true')
    })

    it('should parse to String if value is type of star in Number', function * () {
      let data = fs.readFileSync(filePath)
      let parser = new ExcelParser(data, config.IMPORT_EXCEL.TASK.RULE, config.IMPORT_EXCEL.TASK.LIMIT)
      parser.setHeaderline(0)
      let rows = yield parser.parse()

      expect(rows).be.Array()
      expect(rows[2][13]).equal('4')
    })

    it('should parse to String if value is type of slide in Number', function * () {
      let data = fs.readFileSync(filePath)
      let parser = new ExcelParser(data, config.IMPORT_EXCEL.TASK.RULE, config.IMPORT_EXCEL.TASK.LIMIT)
      parser.setHeaderline(0)
      let rows = yield parser.parse()

      expect(rows).be.Array()
      expect(rows[2][14]).equal('78')
    })

    it('should parse to String if value is type of progress in Number', function * () {
      let data = fs.readFileSync(filePath)
      let parser = new ExcelParser(data, config.IMPORT_EXCEL.TASK.RULE, config.IMPORT_EXCEL.TASK.LIMIT)
      parser.setHeaderline(0)
      let rows = yield parser.parse()

      expect(rows).be.Array()
      expect(rows[2][15]).equal('5')
    })

    it('should parse to String if value is type of menue in Number', function * () {
      let data = fs.readFileSync(filePath)
      let parser = new ExcelParser(data, config.IMPORT_EXCEL.TASK.RULE, config.IMPORT_EXCEL.TASK.LIMIT)
      parser.setHeaderline(0)
      let rows = yield parser.parse()

      expect(rows).be.Array()
      expect(rows[2][16]).equal('项目 2')
    })

    it('should parse to String if cell has bg color and other type', function * () {
      let data = fs.readFileSync(filePath)
      let parser = new ExcelParser(data, config.IMPORT_EXCEL.TASK.RULE, config.IMPORT_EXCEL.TASK.LIMIT)
      parser.setHeaderline(0)
      let rows = yield parser.parse()

      expect(rows).be.Array()
      expect(rows[1][1]).equal('超链接或富文本')
    })

    // 两行合并单元格成为一行，中间若有一列未合并(note)，默认解析成为两行，并共用合并单元格后的值
    it('should parse to two rows if unmerge', function * () {
      let data = fs.readFileSync(filePath)
      let parser = new ExcelParser(data, config.IMPORT_EXCEL.TASK.RULE, config.IMPORT_EXCEL.TASK.LIMIT)
      parser.setHeaderline(0)
      let rows = yield parser.parse()

      expect(rows).be.Array()
      expect(rows.length).equal(5)
      expect(rows[3][1]).equal(rows[4][1])
      expect(rows[3][3]).equal('2')
      expect(rows[4][3]).equal('3')
    })
  })

  // 测试 cellToString函数
  describe('func: cellToString', function () {
    it('should success if valid type', function () {
      let cellToString = ExcelParser.__get__('cellToString')
      let result = cellToString(123)
      expect(result).equal('123')

      result = cellToString(true)
      expect(result).equal('true')

      let date = new Date()
      result = cellToString(date)
      expect(result).equal(date.toISOString())

      result = cellToString({ richText: [{ font: [Object], text: '工作表 1\'!A1' }] })
      expect(result).equal('工作表 1\'!A1')

      result = cellToString({ text: { richText: [{ font: [Object], text: '123' }] } })
      expect(result).equal('123')
    })

    it('should throw err if invalid type', function () {
      let cellToString = ExcelParser.__get__('cellToString')

      try {
        cellToString(new Error('err'))
      } catch (error) {
        return
      }
      throw new Error('not Thrown')
    })
  })

  describe('toArray', function () {
    it('should success', function * () {
      let data = fs.readFileSync(filePath)
      let parser = new ExcelParser(data, config.IMPORT_EXCEL.TASK.RULE, config.IMPORT_EXCEL.TASK.LIMIT)
      parser.setHeaderline(0)
      yield parser.parse()
      let Arrays = yield parser.toArray()

      expect(Arrays[0].content).equal('正常')
    })

    it('should failed when no parsed', function * () {
      let data = fs.readFileSync(filePath)
      let parser = new ExcelParser(data, config.IMPORT_EXCEL.TASK.RULE, config.IMPORT_EXCEL.TASK.LIMIT)
      try {
        yield parser.toArray()
      } catch (error) {
        return
      }
      throw new Error('not Thrown')
    })
  })
})
