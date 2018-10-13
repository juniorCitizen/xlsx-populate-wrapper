# xlsx-populate Wrapper

[xlsx-populate](https://www.npmjs.com/package/xlsx-populate) wrapper library for basic I/O operation on an Excel data file with worksheets that contained a row of headers

1. read any Excel worksheet from spcified Excel file, and get its data in json format
2. dump modified json back into any worksheet and commit the changes to the physical Excel file
3. columns are not added or deleted (`it's not a good idea to update json data from one sheet to another unless the column headings are the same on both worksheets`)
4. only works with existing worksheets

## Installation

```bash
npm install --save xlsx-populate-wrapper
```

## Usage

```javascript
// 1. require the wrapper class
const xPopWrapper = require("xlsx-populate-wrapper")
const path = require("path")

// 2. create an instance by passing in path to an excel file
const filePath = path.resolve(someFilePath)
const workbook = new xPopWrapper(filePath)

workbook.init()
  .then(wb => {
    // work with returned instance
    // to get a list of worksheet names
    console.log(wb.getSheetNames())
    // => ['sheet 1', 'sheet 2', 'sheet 3']

    // or use the initialized variable
    // to get a list of row headings
    console.log(workbook.getHeadings('sheet 1'))
    // => ['title 1', 'title 2', 'title 3']

    // get data from a worksheet
    const jsonData = workbook.getData('sheet 1')
    console.log(jsonData)
    /*
      [
        {
          'title 1': 'a',
          'title 2': 'b',
          'title 3': 'c',
        },
        {
          'title 1': 'd',
          'title 2': 'e',
          'title 3': 'f',
        }
      ]
    */

    workbook.update('sheet 1', someJsonArray) // updating a worksheet

    return workbook.commit() // commit changes to file
  })
  .catch(error => {
    throw error
  })
```
