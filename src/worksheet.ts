export interface IJsonRecord {
  [index:string]:any
}

export interface IWorksheetData {
  aoaData: any[]
  headings: string[]
  jsonData: IJsonRecord[]
}

export interface IWorksheet {
  data (): IWorksheetData
  headings (): string[]
  name():string
  update(worksheetData:IWorksheetData):void
}

export interface IWorksheetConstructor {
  new (worksheet:any):IWorksheet
}

export class Worksheet implements IWorksheet {
  private worksheet: any

  constructor(worksheet:any) {
    this.worksheet = worksheet
  }

  public name():string {
    return this.worksheet.name()
  }

  public data():IWorksheetData {
    const aoaData = this.worksheet.usedRange().value()
    const headings = aoaData.shift().map(mapToString)
    const jsonData = aoaData.map(generateMapCallback(headings))
    return {headings, aoaData, jsonData}
  }

  public headings():string[] {
    return this.worksheet
      .usedRange()
      .value()
      .shift()
      .map(mapToString)
  }

  public update(worksheetData:IWorksheetData):void {
    this.worksheet.usedRange().clear()
    this.worksheet.cell('A1').value([worksheetData.headings])
    this.worksheet.cell('A2').value(worksheetData.aoaData)
  }
}

function mapToString(value:any):string {
  return value.toString()
}

function generateMapCallback(headings:string[]) {
  return (aoaRecord:any[]):IJsonRecord => {
    return headings.reduce(generateReduceCallback(aoaRecord), {})
  }
}

function generateReduceCallback(aoaRecord:any[]) {
  return (jsonRecord:IJsonRecord, heading:string, index:number):IJsonRecord => {
    jsonRecord[heading] = aoaRecord[index]
    return jsonRecord
  }
}
