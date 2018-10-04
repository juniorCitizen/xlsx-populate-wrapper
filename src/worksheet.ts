export {unionArrays} from './wsData'

// export interface IJsonRecord {
//   [index: string]: any
// }

// export interface IWorksheetData {
//   aoaData: any[]
//   headings: string[]
//   jsonData: IJsonRecord[]
// }

// export interface IWorksheet {
//   data(): IWorksheetData
//   headings(): string[]
//   name(): string
//   update(worksheetData: IWorksheetData): void
// }

// export class Worksheet implements IWorksheet {
//   private worksheet: any

//   constructor(worksheet: any) {
//     this.worksheet = worksheet
//   }

//   public data(): IWorksheetData {
//     if (!this.worksheet) {
//       throw new Error('worksheet is not ready')
//     }
//     const usedRange: any = this.worksheet.usedRange()
//     if (usedRange) {
//       const aoaData: any[] = usedRange.value()
//       const headings: string[] = aoaData.shift().map(mapToString)
//       const jsonData: any[] = aoaData.map(mapAoaToJson(headings))
//       return {headings, aoaData, jsonData}
//     } else { return {headings: [], aoaData: [], jsonData: []} }
//   }

//   public headings(): string[] {
//     if (!this.worksheet) {
//       throw new Error('worksheet is not ready')
//     }
//     return this.worksheet
//       .usedRange()
//       .value()
//       .shift()
//       .map(mapToString)
//   }

//   public name(): string {
//     if (!this.worksheet) {
//       throw new Error('worksheet is not ready')
//     }
//     return this.worksheet.name()
//   }

//   public update(worksheetData: IWorksheetData): void {
//     this.worksheet.usedRange().clear()
//     this.worksheet.cell('A1').value([worksheetData.headings])
//     this.worksheet.cell('A2').value(worksheetData.aoaData)
//   }
// }

// function mapToString(value: any): string {
//   return value.toString()
// }

// function mapAoaToJson(headings: string[]) {
//   return (aoaRecord: any[]): IJsonRecord => {
//     return headings.reduce(reduceByHeadings(aoaRecord), {})
//   }
// }

// function reduceByHeadings(aoaRecord: any[]) {
//   return (
//     jsonRecord: IJsonRecord,
//     heading: string,
//     index: number
//   ): IJsonRecord => {
//     jsonRecord[heading] = aoaRecord[index]
//     return jsonRecord
//   }
// }
