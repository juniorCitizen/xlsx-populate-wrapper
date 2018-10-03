# xlsx-populate Wrapper

[xlsx-populate](https://www.npmjs.com/package/xlsx-populate) wrapper library for working with excel files

## Install

```bash
npm install --save xlsx-populate-wrapper
```

## Usage

```javascript
// 1. require the wrapper class
const ExcelWrapper = require("xlsx-populate-wrapper")
const path = require("path")

// 2. create an instance by passing in path to an excel file
const filePath = path.resolve(someFilePath)
const workbook = new ExcelWrapper(filePath)

workbook.initialize()
  .then(wb => {
    // work with returned instance and getting data from a specific worksheet
    const ws = wb.worksheet('example')
    console.log(ws.data())
    // or use the initialized variable and get an array of data on every existing worksheet
    console.log(workbook.data())
    // updating a worksheet
    const workingData = ws.data() // data from 'example' worksheet
    console.log(workingData.headings) // row 1 values
    console.log(workingData.aoaData) // data in array of arrays without row 1
    console.log(workingData.jsonData) // data in array of objects, with headings as prop keys
    // do some work with the jsonData
    const copy = JSON.parse(JSON.stringify(workingData.jsonData[0]))
    workingData.jsonData = workingData.jsonData.push(copy) // add one obj to the end of the array
    return workbook.update('example', workingData) // update the physical file
  })
  .then(wb=>{
    const jsonData = wb.worksheet('example').data().jsonData
    console.log(jsonData) // should find the data with the additional record
    return Promise.resolve()
  })
  .catch(error => {
    throw error
  })
```
