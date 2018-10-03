import * as XlsxPopulate from 'xlsx-populate'

import {
  IJsonRecord,
  IWorksheet,
  IWorksheetConstructor,
  IWorksheetData,
  Worksheet,
} from './worksheet'

export interface IWorkbook {
  convertJson(jsonData: IJsonRecord[], headings: string[]): IWorksheetData
  data(): IWorksheetData[]
  initialize(): Promise<this>
  instantiateWorkbook(
    constructor: IWorkbookConstructor,
    filePath: string
  ): IWorkbook
  update(worksheetName: string, worksheetData: IWorksheetData): Promise<this>
  worksheet(worksheetName: string): IWorksheet
  worksheetNames(): string[]
}

export interface IWorkbookConstructor {
  new(filePath: string): IWorkbook
}

export class Workbook implements IWorkbook {
  private filePath: string
  private workbook: any
  private worksheets: IWorksheet[]

  constructor(filePath: string) {
    this.filePath = filePath
    this.workbook = null
    this.worksheets = []
  }

  public convertJson(
    jsonData: IJsonRecord[],
    headings: string[] = []
  ): IWorksheetData {
    if (headings.length === 0) {
      headings = Object.keys(jsonData[0])
    }
    jsonData = jsonData.map(sanitizeJsonRecord(headings))
    const aoaData = jsonData.map(
      (jsonRecord: IJsonRecord): any[][] => {
        const reduceFn = (arrayRecord: any[], heading: string): any[] => {
          arrayRecord.push(jsonRecord[heading])
          return arrayRecord
        }
        return headings.reduce(reduceFn, [])
      }
    )
    return { headings, aoaData, jsonData }
  }

  public data(): IWorksheetData[] {
    if (!this.workbook) {
      throw new Error('workbook is not ready')
    }
    return this.worksheets.map(this.getWsData)
  }

  public instantiateWorkbook(
    constructor: IWorkbookConstructor,
    filePath: string
  ): IWorkbook {
    return new constructor(filePath)
  }

  public async initialize(): Promise<this> {
    try {
      this.workbook = await XlsxPopulate.fromFileAsync(this.filePath)
      this.worksheets = this.workbook
        .sheets()
        .reduce((worksheets: any[], worksheet: any): any[] => {
          worksheets.push(this.instantiateWorksheet(Worksheet, worksheet))
          return worksheets
        }, [])
      return this
    } catch (error) {
      this.filePath = ''
      this.workbook = null
      this.worksheets = []
      throw error
    }
  }

  public async update(wsName: string, wsData: IWorksheetData): Promise<this> {
    if (!this.workbook) {
      throw new Error('workbook is not ready')
    }
    const { headings, jsonData } = wsData
    wsData = this.convertJson(jsonData, headings)
    this.worksheet(wsName).update(wsData)
    try {
      await this.workbook.toFileAsync(this.filePath)
      return this
    } catch (error) {
      this.worksheets = this.workbook
        .sheets()
        .reduce((worksheets: any[], worksheet: any): any[] => {
          worksheets.push(this.instantiateWorksheet(Worksheet, worksheet))
          return worksheets
        }, [])
      throw error
    }
  }

  public worksheet(wsName: string): IWorksheet {
    if (!this.workbook) {
      throw new Error('workbook is not ready')
    }
    const searchFn = this.worksheetNames()
    const wsIndex = searchFn.findIndex(this.findWorksheet(wsName))
    if (wsIndex !== -1) {
      return this.worksheets[wsIndex]
    } else {
      throw new Error(`worksheet ${wsName}} not found`)
    }
  }

  public worksheetNames(): string[] {
    if (!this.workbook) {
      throw new Error('workbook is not ready')
    }
    return this.worksheets.map(worksheet => worksheet.name())
  }

  private getWsData(ws: IWorksheet): IWorksheetData {
    return ws.data()
  }

  private instantiateWorksheet(
    constructor: IWorksheetConstructor,
    worksheet: any
  ): IWorksheet {
    return new constructor(worksheet)
  }

  private findWorksheet(worksheetName: string) {
    return (wsName: string): boolean => wsName === worksheetName
  }
}

function sanitizeJsonRecord(properties: string[]) {
  return (jsonRecord: IJsonRecord): IJsonRecord => {
    const originalRecord = JSON.parse(JSON.stringify(jsonRecord))
    return properties.reduce(sanitize(originalRecord), {})
  }
}

function sanitize(originalRecord: IJsonRecord) {
  return (sanitized: IJsonRecord, property: string): IJsonRecord => {
    const isValid: boolean = originalRecord.hasOwnProperty(property)
    sanitized[property] = isValid ? originalRecord[property] : null
    return sanitized
  }
}
