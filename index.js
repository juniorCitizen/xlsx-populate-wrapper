const XlsxPopulate = require("xlsx-populate")

class Xlsx {
  constructor() {
    this.filePath = null
    this.workbook = null
    this.worksheets = null
  }

  initialize(filePath) {
    this.filePath = filePath
    XlsxPopulate.fromFileAsync(filePath)
      .then(workbook => {
        this.workbook = workbook
        this.worksheets = workbook.sheets()
      })
      .catch(error => Promise.reject(error))
  }

  get workbook() {
    return this.workbook
  }

  get worksheets() {
    return this.worksheets
  }
}

module.exports = Xlsx
