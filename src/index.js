const XlsxPopulate = require('xlsx-populate')

/**
 * custom data object, representing an excel worksheet data in array of arrays or array of objects
 * @typedef {Object} worksheetData
 * @property {string[]} headings - first value of every column
 * @property {Array.<Array.<string|number>>} aoaData - array of array data
 * @property {Object[]} jsonData - array of data object
 * @property {string|number} jsonData.propA - value of propA
 * @property {string|number} jsonData.propB - value of propB
 * @property {string|number} jsonData.propN - value of propN, etc...
 */

/**
 * represents an excel workbook
 *
 * @typedef {Class} Workbook
 * @example
 * // 1. require the wrapper class
 * const ExcelWrapper = require("xlsx-populate-wrapper")
 * const path = require("path")
 *
 * // 2. create an instance by passing in path to an excel file
 * const filePath = path.resolve(someFilePath)
 * new ExcelWrapper(filePath)
 *   .then(workbook => {
 *     const worksheet = workbook.worksheet("example")
 *     console.log(worksheet.data())
 *   })
 *   .catch(error => {
 *     throw error
 *   })
 */
class Workbook {
  /**
   * initialize a Workbook object
   *
   * @param {string} filePath - absolute path to an excel file
   * @returns {Workbook} Workbook instance
   */
  constructor(filePath) {
    this.filePath = filePath
    return XlsxPopulate.fromFileAsync(this.filePath)
      .then(workbook => {
        this.workbook = workbook
        this.worksheets = this.workbook
          .sheets()
          .reduce((worksheets, worksheet) => {
            worksheets.push(new Worksheet(worksheet))
            return worksheets
          }, [])
        return this
      })
      .catch(error => Promise.reject(error))
  }

  /**
   * get dataset of all existing worksheets
   *
   * @returns {worksheetData[]} array of worksheetData
   */
  data() {
    return this.worksheets.map(worksheet => worksheet.data())
  }

  /**
   * update and persist data of a particular worksheet
   *
   * @param {string} worksheetName - name of target worksheet to update
   * @param {Object[]|worksheetData[]} dataset - can be an array of objects or an worksheetData object.  Only 'jsonData' property is required to use the later
   */
  update(worksheetName, dataset) {
    const { headings, jsonData } = dataset
    dataset =
      jsonData && headings
        ? this.convertJson(jsonData, headings)
        : this.convertJson(dataset)
    const worksheet = this.worksheet(worksheetName)
    if (worksheet) {
      this.worksheet(worksheetName).update(dataset)
      return this.workbook
        .toFileAsync(this.filePath)
        .then(() => Promise.resolve(this))
        .catch(error => Promise.reject(error))
    } else {
      const error = new Error(`'${worksheetName}' does not exist`)
      return Promise.reject(error)
    }
  }

  /**
   * get the worksheet names of existing worksheets
   *
   * @returns {string[]} - names of existing worksheets
   */
  worksheetNames() {
    return this.worksheets.map(worksheet => worksheet.name())
  }

  /**
   * get a specific worksheet by worksheet name
   *
   * @param {string} worksheetName - name of target worksheet
   * @returns {Worksheet} Worksheet instance
   */
  worksheet(worksheetName) {
    const findIndexFn = wsName => wsName === worksheetName
    const wsIndex = this.worksheetNames().findIndex(findIndexFn)
    return wsIndex === -1 ? null : this.worksheets[wsIndex]
  }

  /**
   * convert an array of objects to a worksheetData object
   *
   * @param {Object[]} jsonData - array of json objects
   * @param {string[]} headings - (optional) column headings, props of first json object is used when headings are omitted
   * @returns {worksheetData} converted data
   */
  static convertJson(jsonData, headings = null) {
    if (!headings) headings = Object.keys(jsonData[0])
    jsonData = jsonData.map(jsonRecord => sanitize(jsonRecord, headings))
    const aoaData = jsonData.map(jsonRecord => {
      return headings.reduce((arrayRecord, heading) => {
        arrayRecord.push(jsonRecord[heading])
        return arrayRecord
      }, [])
    })
    return { headings, aoaData, jsonData }
  }
}

/**
 * represents an Excel worksheet
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
    this.worksheet = worksheet
  }

  /**
   * get the worksheet name
   *
   * @returns {string} name of worksheet
   */
  name() {
    return this.worksheet.name()
  }

  /**
   * get data from the worksheet
   *
   * @returns {worksheetData[]} data from worksheet
   */
  data() {
    const aoaData = this.worksheet.usedRange().value()
    const mapFn = header => header.toString()
    const headings = aoaData.shift().map(mapFn)
    const jsonData = aoaData.map(record => {
      const reduceFn = (jsonData, heading, index) => {
        return jsonData[heading] === record[index]
      }
      return headings.reduce(reduceFn, {})
    })
    return { headings, aoaData, jsonData }
  }

  /**
   * get headings of all columns
   *
   * @returns {string[]} column headings
   */
  headings() {
    const mapFn = header => header.toString()
    return this.worksheet
      .usedRange()
      .value()
      .shift()
      .map(mapFn)
  }

  /**
   * update worksheet data (changes are not persisted with this method)
   *
   * @param {worksheetData} dataset - data to update the worksheet with
   * @param {string[]} dataset.headings - heading of columns
   * @param {Array.<Array.<string|number>>} dataset.aoaData - data in array of arrays format
   */
  update({ headings, aoaData }) {
    this.worksheet.usedRange().clear()
    this.worksheet.cell('A1').value(headings)
    this.worksheet.cell('B1').value(aoaData)
  }
}

module.exports = Workbook

/**
 * compare the object's existing properties against a list of strings
 *
 * @param {Object} object - json object to be sanitized
 * @param {string[]} properties - property keys to keep
 */
function sanitize(object, properties) {
  let data = JSON.parse(JSON.stringify(object))
  return properties.reduce((newObject, property) => {
    const isValid = data.hasOwnProperty(property)
    newObject[property] = isValid ? data[property] : null
    return newObject
  }, {})
}

module.exports = Worksheet
