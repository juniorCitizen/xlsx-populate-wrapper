export {unionArrays} from './worksheet'

export function add(numA: number, numB: number): number {
  return numA + numB
}

// import * as XlsxPopulate from 'xlsx-populate'

// import {IJsonRecord, IWorksheet, IWorksheetData, Worksheet} from './worksheet'

// export interface IWorkbook {
//   addSheet(worksheetName: string, worksheetData: IWorksheetData): Promise<this>
//   checkWsData(jsonData: IJsonRecord[], headings: string[]): IWorksheetData
//   data(): IWorksheetData[]
//   initialize(): Promise<this>
//   update(worksheetName: string, worksheetData: IWorksheetData): Promise<this>
//   worksheet(worksheetName: string): IWorksheet
//   worksheetNames(): string[]
// }

// export class Workbook implements IWorkbook {
//   private filePath: string
//   private workbook: any
//   private worksheets: IWorksheet[]

//   constructor(filePath: string) {
//     this.filePath = filePath
//     this.workbook = null
//     this.worksheets = []
//   }

//   public checkWsData(
//     jsonData: IJsonRecord[],
//     headings: string[] = []
//   ): IWorksheetData {
//     if (headings.length === 0) {
//       headings = Object.keys(jsonData[0])
//     }
//     jsonData = jsonData.map(sanitizeJsonRecord(headings))
//     const aoaData = jsonData.map(
//       (jsonRecord: IJsonRecord): any[][] => {
//         const reduceFn = (arrayRecord: any[], heading: string): any[] => {
//           arrayRecord.push(jsonRecord[heading])
//           return arrayRecord
//         }
//         return headings.reduce(reduceFn, [])
//       }
//     )
//     return {headings, aoaData, jsonData}
//   }

//   public async addSheet(
//     worksheetName: string,
//     worksheetData: IWorksheetData
//   ): Promise<this> {
//     const newWorksheet = this.workbook.addSheet(worksheetName)
//     this.worksheets.push(new Worksheet(newWorksheet))
//     const {headings, jsonData} = worksheetData
//     worksheetData = this.checkWsData(jsonData, headings)
//     this.worksheet(worksheetName).update(worksheetData)
//     try {
//       await this.workbook.toFileAsync(this.filePath)
//       return this
//     } catch (error) {
//       this.worksheets = this.workbook
//         .sheets()
//         .reduce((worksheets: any[], worksheet: any): any[] => {
//           worksheets.push(new Worksheet(worksheet))
//           return worksheets
//         }, [])
//       throw error
//     }
//   }

//   public data(): IWorksheetData[] {
//     if (!this.workbook) {
//       throw new Error('workbook is not ready')
//     }
//     return this.worksheets.map(this.getWorksheetData)
//   }

//   public async initialize(): Promise<this> {
//     try {
//       this.workbook = await XlsxPopulate.fromFileAsync(this.filePath)
//       this.worksheets = this.workbook
//         .sheets()
//         .reduce((worksheets: any[], worksheet: any): any[] => {
//           worksheets.push(new Worksheet(worksheet))
//           return worksheets
//         }, [])
//       return this
//     } catch (error) {
//       this.filePath = ''
//       this.workbook = null
//       this.worksheets = []
//       throw error
//     }
//   }

//   public async update(wsName: string, wsData: IWorksheetData): Promise<this> {
//     if (!this.workbook) {
//       throw new Error('workbook is not ready')
//     }
//     const {headings, jsonData} = wsData
//     wsData = this.checkWsData(jsonData, headings)
//     this.worksheet(wsName).update(wsData)
//     try {
//       await this.workbook.toFileAsync(this.filePath)
//       return this
//     } catch (error) {
//       this.worksheets = this.workbook
//         .sheets()
//         .reduce((worksheets: any[], worksheet: any): any[] => {
//           worksheets.push(new Worksheet(worksheet))
//           return worksheets
//         }, [])
//       throw error
//     }
//   }

//   public worksheet(wsName: string): IWorksheet {
//     if (!this.workbook) {
//       throw new Error('workbook is not ready')
//     }
//     const searchFn = this.worksheetNames()
//     const wsIndex = searchFn.findIndex(this.findWorksheet(wsName))
//     if (wsIndex !== -1) {
//       return this.worksheets[wsIndex]
//     } else {
//       throw new Error(`worksheet ${wsName}} not found`)
//     }
//   }

//   public worksheetNames(): string[] {
//     if (!this.workbook) {
//       throw new Error('workbook is not ready')
//     }
//     return this.worksheets.map(worksheet => worksheet.name())
//   }

//   private findWorksheet(worksheetName: string) {
//     return (wsName: string): boolean => wsName === worksheetName
//   }

//   private getWorksheetData(ws: IWorksheet): IWorksheetData {
//     return ws.data()
//   }
// }

// function sanitizeJsonRecord(properties: string[]) {
//   return (jsonRecord: IJsonRecord): IJsonRecord => {
//     const originalRecord = JSON.parse(JSON.stringify(jsonRecord))
//     return properties.reduce(sanitize(originalRecord), {})
//   }
// }

// function sanitize(originalRecord: IJsonRecord) {
//   return (sanitized: IJsonRecord, property: string): IJsonRecord => {
//     const isValid: boolean = originalRecord.hasOwnProperty(property)
//     sanitized[property] = isValid ? originalRecord[property] : null
//     return sanitized
//   }
// }
