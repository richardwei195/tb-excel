## tb-excel

![Build Status](https://travis-ci.org/richardwei195/tb-excel.svg?branch=master)
[![NPM version](https://img.shields.io/npm/v/tb-excel.svg)](https://www.npmjs.com/package/tb-excel])

> When you get excel binary data from http or a file, you can use this to parse it and get  object you want

## Installation

```shell
npm install tb-excel
```
## Support 
| Source |   rule | limit | toObject |
| ----- | ------- | -------- | --------- |
| binary   | support | support  | support   |
| utf8+BOM | support | support  | support   |
| utf16    | support | support  | support   |


## Notice

- Use `exceljs` native method:

  ```javascript
    const Excel = require('tb-excel').exceljs
  ```
- Use `tb-excel` extends methods. such as, use xlsx parse methods:

  ```javascript
    const Excel = require('tb-excel').xlsxParser
  ```

## Example

- Read from url: we want get excel content not only from local files but also from others

  For Example: 

  First: Use request to get data from server

    ```javascript
      let res = request.get(downloadUrl, { encoding: null }, function(error, response, body) {
          // ……
        })
      }
    ```

  Second: We can parse some binary data which get from request: 

    ```javascript
      'use strict'
      const ExcelParser = require('tb-excel').xlsxParser

      const co = require('co')

      let data = 'you get from request'
      let rule = ['a', 'b']
      const excel = new ExcelParser(data, rule, 3)

      //  you can use gengerator or async, depends yourself

      function * parser () {
        let row = yield excel.parse()
        console.log(row) // we can get each rows content from excel
      }
      co(parser())
    ```
  Last: Format data according to you rule

    ```javascript
    
     obj = row.toArray()
     console.log(obj)
     // [{
	  a: '',
          b: ''
	}]

    ```

## Doc

### Create constructor

```javascript
  const ExcelParser = require('tb-excel').xlsxParser

  const excel = new ExcelParser(data, ['a', 'b'], 3)
```

### Get Arrays from excel

```javascript
  const excel = new ExcelParser(data)
  let ret = yield excel.parse() // ['a', 18, 'man']
```

### Get Array Obj you want

```javascript
  const rule = [
    'name',
    'age',
    'sex'
  ]

  const excel = new ExcelParser(data, rule)
  let ret = yield excel.parse()
  excel.toArray() // // { name: 'a', age: 18, sex: 'man' }
```

### Skip rows

```javascript
  const rule = [
    'name',
    'age',
    'sex'
  ]

  const skip = 1

  const excel = new ExcelParser(data, rule)

  excel.setHeaderline(skip)
  yield excel.parse()
  excel.toArray() // { age: 18, sex: 'man' }
```

### Check if exceed limit

```javascript
  const rule = [
    'name',
    'age',
    'sex'
  ]

  const excel = new ExcelParser(data, rule, 1)
  yield excel.parse()
  excel.toArray() // { name: 'a', age: 18, sex: 'man' }

  excel.isExccedLimit // true
```
