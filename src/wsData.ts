import {union} from 'lodash'

export function unionArrays(arrays: string[][]): string[] {
  return union(...arrays)
}

// export interface IJsonRecord {
//   [index: string]: any
// }

// export interface IWsDataStructure {
//   headings: string[]
//   aoaData: any[][]
//   jsonData: IJsonRecord[]
// }

// export interface IWsDataClass {
//   getData(): IWsDataStructure
//   getHeadings(): string[]
//   getAoaData(): any[][]
//   getJsonData(): IJsonRecord[]
// }

// export default class WsData implements IWsDataClass {
//   private headings: string[]
//   private aoaData: any[][]
//   private jsonData: IJsonRecord[]
//   constructor(jsonData: any[]) {
//     this.jsonData = jsonData
//     this.headings = extractHeadings(jsonData)
//     this.aoaData = mapJsonToAoa(this.jsonData, this.headings)
//   }

//   public getData(): IWsDataStructure {
//     return {
//       headings: this.headings,
//       aoaData: this.aoaData,
//       jsonData: this.jsonData,
//     }
//   }

//   public getAoaData(): any[][] {
//     return this.aoaData
//   }

//   public getHeadings(): string[] {
//     return this.headings
//   }

//   public getJsonData(): IJsonRecord[] {
//     return this.jsonData
//   }
// }

// export function mapJsonToAoa(
//   jsonData: IJsonRecord[],
//   headings: string[]
// ): any[][] {
//   if (jsonData.length === 0 || headings.length === 0) {
//     return []
//   }
//   return []
// }

// export function extractHeadings(jsonData: IJsonRecord[]): string[] {
//   const keyMapper = (o: IJsonRecord): string[] => Object.keys(o)
//   const stringSorter = (a: string, b: string): number =>
//     a > b ? 1 : a < b ? -1 : 0
//   return union(...jsonData.map(keyMapper)).sort(stringSorter)
// }
