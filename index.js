const XlsxPopulate = require('xlsx-populate')

/** represents a worksheet */

/**
 * custom worksheet class object, represents an Excel worksheet
 *
 * @typedef {Class} Worksheet
 */
class Worksheet {
  /**
   * initialize a worksheet
   *
   * @param {Object} worksheet - Xlsx-populate.sheet([name of sheet]) object
   */
  constructor(worksheet) {
    this._worksheet = worksheet
    this._name = this._worksheet.name()
    this._data = {}
    this._data.aoa = this._worksheet.usedRange().value()
    const mappingFn = columnHeader => columnHeader.toString()
    this._data.columnHeaders = this._data.aoa.shift().map(mappingFn)
    this._data.json = this._data.aoa.map(record => {
      return this._data.columnHeaders.reduce((jsonData, header, index) => {
        jsonData[header] = record[index]
        return jsonData
      }, {})
    })
  }

  /**
   * get name of the worksheet
   *
   * @returns {string} name of worksheet
   */
  name() {
    return this._name
  }

  /**
   * custom 'WorksheetData' object definition
   *
   * @typedef {Object} WorksheetData
   * @property {string[]} columnHeaders - values of cell 1 on each column
   * @property {(string|number)[][]} aoa - array of arrays representing data in rows
   * @property {Object[]} json - json objects with columnHeader as key to each property
   */

  /**
   * get worksheet data in different formats
   *
   * @returns {WorksheetData} worksheet data object containing column headers, data in aoa and json formats
   */
  data() {
    return {
      columnHeaders: this._data.columnHeaders,
      aoa: this._data.aoa,
      json: this._data.json,
    }
  }
}

/**
 * custom workbook class object, representing an Excel workbook
 *
 * @typedef {Class} Workbook
 */
class Workbook {
  /**
   * creating a Workbook object
   */
  constructor() {
    this._filePath = null
    this._workbook = null
    this._worksheets = null
  }

  /**
   * loads the excel file indicated by the filePath
   *
   * @param {string} filePath - absolute path to excel file
   * @returns {Workbook} instance
   */
  async initialize(filePath) {
    try {
      this._filePath = filePath
      this._workbook = await XlsxPopulate.fromFileAsync(this._filePath)
      this._worksheets = extractWorksheets(this._workbook)
      return this
    } catch (error) {
      throw error
    }
  }

  /**
   * get names of existing worksheets in the workbook
   *
   * @returns {string[]} - names of existing worksheets
   */
  worksheetNames() {
    return Object.keys(this._worksheets)
  }

  /**
   * get dataset of the complete workbook in json format
   *
   * @returns {Object} datasets for the complete workbook in json format
   */
  data() {
    let dataset = {}
    for (let key in this._worksheets) {
      dataset[key] = this._worksheets[key].data()
    }
    return dataset
  }

  /**
   * add new worksheet or update an existing worksheet
   * and commit addition/updates to the physical excel file
   *
   * @param {string} worksheetName - name of existing or new worksheet
   * @param {WorksheetData} worksheetData - custom worksheet data object
   */
  async commit(worksheetName, worksheetData) {
    // check if the sheet existed by comparing worksheet names
    const predicate = existingWsName => existingWsName === worksheetName
    const existAtIndex = Object.keys(this._worksheets).findIndex(predicate)
    if (existAtIndex === -1) {
      // sheet does not exist
      addWorksheet(this._workbook, worksheetName, worksheetData) // add a new worksheet to workbook
    } else {
      // found to be an existing worksheet, proceed to update
      updateWorksheet(this._workbook, worksheetName, worksheetData) // update an existing worksheet
    }
    this._worksheets = extractWorksheets(this._workbook) // update in-memory data
    try {
      // commit workbook to physical excel file on disk
      return await this._workbook.toFileAsync(this._filePath)
    } catch (error) {
      throw error
    }
  }
}

module.exports = Workbook

/**
 * prepare the aoa data from custom 'WorksheetData' to be committed to file
 *
 * @param {WorksheetData} worksheetData - custom 'WorksheetData' object
 * @returns {(string|number)[][]} array of arrays representing data in rows (including column header titles)
 */
function prepAoaData(worksheetData) {
  return worksheetData.columnHeaders.length > 0
    ? [worksheetData.columnHeaders, ...worksheetData.aoa]
    : [...worksheetData.aoa]
}

/**
 * update contents of a worksheet
 *
 * @param {Object} workbook - Xlsx-populate workbook object
 * @param {string} worksheetName - name of the worksheet to update
 * @param {WorksheetData} worksheetData - custom 'WorksheetData' object
 */
function updateWorksheet(workbook, worksheetName, worksheetData) {
  const sheetContents = prepAoaData(worksheetData)
  const worksheet = workbook.sheet(worksheetName)
  worksheet.usedRange().clear()
  worksheet.cell('A1').value(sheetContents)
}

/**
 * add a new worksheet to Xlsx-populate 'workbook' object
 *
 * @param {Object} workbook - Xlsx-populate workbook object
 * @param {string} worksheetName - name of the new worksheet
 * @param {WorksheetData} worksheetData - custom 'WorksheetData' object
 */
function addWorksheet(workbook, worksheetName, worksheetData) {
  // prepare the aoa data to be committed to file
  const hasColumnHeaders = worksheetData.columnHeaders.length > 0
  const sheetContents = hasColumnHeaders
    ? [worksheetData.columnHeaders, ...worksheetData.aoa]
    : [...worksheetData.aoa]
  workbook
    .addSheet(worksheetName)
    .cell('A1')
    .value(sheetContents)
}

/**
 * convert a xlsx-populate workbook object to an object of custom 'Worksheet' objects indexed by worksheet name
 *
 * @param {Object} workbook - Xlsx-populate workbook object
 * @returns {Object} object of custom 'Worksheet' objectobject of custom 'Worksheet' object indexed by worksheet name indexed by worksheet name
 */
function extractWorksheets(workbook) {
  return workbook.sheets().reduce((worksheets, worksheet) => {
    worksheets[worksheet.name()] = new Worksheet(worksheet)
    return worksheets
  }, {})
}
