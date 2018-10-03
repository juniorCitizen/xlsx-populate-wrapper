import * as XlsxPopulate from 'xlsx-populate'

import {
  IJsonRecord,
  IWorksheet,
  IWorksheetConstructor,
  IWorksheetData,
  Worksheet,
} from './worksheet'

export interface IWorkbook
{
  convertJson ( jsonData: IJsonRecord[], headings: string[] ): IWorksheetData
  data (): IWorksheetData[]
  init(): Promise<this>
  update ( worksheetName: string, worksheetData: IWorksheetData ): Promise<void>
  worksheet ( worksheetName: string ): IWorksheet
  worksheetNames():string[]
}

// export interface IWorkbookConstructor {
//   new (filePath: string): IWorkbook
// }

export class Workbook implements IWorkbook {
  private filePath:string
  private workbook: any
  private worksheets: IWorksheet[]

  constructor(filePath:string) {
    this.filePath = filePath
    this.workbook = null
    this.worksheets = []
  }

  public convertJson (jsonData:IJsonRecord[], headings:string[]=[]): IWorksheetData {
    if (headings.length === 0) {
      headings = Object.keys(jsonData[0])
    }
    jsonData = jsonData.map(sanitizeJsonRecord(headings))
    const aoaData = jsonData.map( (jsonRecord:IJsonRecord):any[][] => {
      const reduceFn = (arrayRecord: any[], heading: string): any[] => {
        arrayRecord.push( jsonRecord[heading] )
        return arrayRecord
      }
      return headings.reduce(reduceFn, [] )
    })
    return { headings, aoaData, jsonData }
  }

  private getWsData(ws:IWorksheet):IWorksheetData{
    return ws.data()
  }

  public data (): IWorksheetData[]  {
    if ( !this.workbook ) { throw new Error( 'workbook is not ready' ) }
    return this.worksheets.map(this.getWsData)
  }

  public init (): Promise<this>
  {
    return XlsxPopulate.fromFileAsync( this.filePath )
      .then((workbook:any):Promise<this> =>
      {
        this.workbook = workbook
        this.worksheets = this.workbook
          .sheets()
          .reduce((worksheets:any[],worksheet:any):any[] => {
            worksheets.push( new Worksheet( worksheet ) )
            return worksheets
          }, [] )
        return Promise.resolve(this)
      } )
      .catch((error:any):Promise<any>=> Promise.reject( error ) )
  }

  public update (worksheetName: string, worksheetData: IWorksheetData ): Promise<void>  {
    if (!this.workbook) { throw new Error('workbook is not ready') }
    const {headings, jsonData} = worksheetData
    worksheetData = this.convertJson(jsonData, headings)
    this.worksheet(worksheetName).update(worksheetData)
    return this.workbook
      .toFileAsync( this.filePath )
      .then(()=> Promise.resolve( this ) )
      .catch((error:any):Promise<any>=>Promise.reject(error))
   }

  private findWorksheet(worksheetName:string) {
    return (wsName:string):boolean => wsName === worksheetName
  }

  public worksheet(wsName: string): IWorksheet {
    if (!this.workbook) { throw new Error('workbook is not ready') }
    const searchFn = this.worksheetNames()
    const wsIndex = searchFn.findIndex(this.findWorksheet(wsName))
    if(wsIndex!==-1) {
      return this.worksheets[wsIndex]
    } else {
      throw new Error(`worksheet ${wsName}} not found`)
    }
  }

  public worksheetNames():string[] {
    if (!this.workbook) { throw new Error('workbook is not ready') }
    return this.worksheets.map(worksheet => worksheet.name())
  }
}

function sanitizeJsonRecord(properties: string[]) {
  return (jsonRecord: IJsonRecord): IJsonRecord => {
    const originalRecord = JSON.parse(JSON.stringify(jsonRecord))
    return properties.reduce(reduceFn(originalRecord), {})
  }
}

function reduceFn(originalRecord:IJsonRecord) {
  return (sanitized: IJsonRecord, property: string): IJsonRecord => {
    const isValid: boolean = originalRecord.hasOwnProperty(property)
    sanitized[property] = isValid ? originalRecord[property] : null
    return sanitized
  }
}
