import * as XPop from 'xlsx-populate'

interface IWorkbook {
  commit: () => Promise<this>
  getData: (wsName: string) => any[]
  getHeadings: (wsName: string) => any[]
  getSheetNames: () => string[]
  init: () => Promise<this>
  update: (wsName: string, jsonData: any[]) => void
}

class Workbook implements IWorkbook {
  private filePath: string = ''
  private workbook: any = null
  constructor(filePath: string = '') {
    this.filePath = filePath
  }

  public async init(): Promise<this> {
    try {
      this.workbook = await XPop.fromFileAsync(this.filePath)
      return this
    } catch (error) {
      this.filePath = ''
      this.workbook = null
      throw error
    }
  }

  public getSheetNames(): string[] {
    return this.workbook.sheets().map(
      (xPopSheet: any): string => {
        return xPopSheet.name()
      }
    )
  }

  /**
   * get data from a particular worksheet
   *
   * @param wsName - name of the worksheet to get data from
   * @return array of objects representing a record
   */
  public getData(wsName: string): any[] {
    const headings: string[] = this.getHeadings(wsName)
    const aoaData: any[][] = this.getAoaData(wsName)
    return aoaData.map(
      (aoaRecord: any[]): any => {
        return headings.reduce(
          (jsonRecord: any, heading: string, keyIndex: number): any => {
            jsonRecord[heading] = aoaRecord[keyIndex]
            return jsonRecord
          },
          {}
        )
      }
    )
  }

  public async commit(): Promise<this> {
    try {
      await this.workbook.toFileAsync(this.filePath)
      return this
    } catch (error) {
      throw error
    }
  }

  public update(wsName: string, jsonData: any[]): void {
    const headings: string[] = this.getHeadings(wsName)
    const aoaData: any[][] = jsonData.map(
      (jsonRecord: any): any[] => {
        return headings.reduce((aoaRecord: any[], heading: string): any[] => {
          aoaRecord.push(jsonRecord[heading] || undefined)
          return aoaRecord
        }, [])
      }
    )
    const worksheet = this.workbook.sheet(wsName)
    const dataRange: any = worksheet.usedRange()
    dataRange.clear()
    worksheet.cell('A1').value([headings])
    worksheet.cell('A2').value([...aoaData])
  }

  public getHeadings(wsName: string): any[] {
    return this.workbook
      .sheet(wsName)
      .usedRange()
      .value()
      .shift()
  }

  private getAoaData(wsName: string): any[][] {
    const aoaData: any[][] = this.workbook
      .sheet(wsName)
      .usedRange()
      .value()
    aoaData.shift()
    return aoaData
  }
}

export = Workbook
