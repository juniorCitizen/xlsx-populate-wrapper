import { IJsonRecord, IWorksheet, IWorksheetData } from './worksheet';
export interface IWorkbook {
    convertJson(jsonData: IJsonRecord[], headings: string[]): IWorksheetData;
    data(): IWorksheetData[];
    initialize(): Promise<this>;
    instantiateWorkbook(constructor: IWorkbookConstructor, filePath: string): IWorkbook;
    update(worksheetName: string, worksheetData: IWorksheetData): Promise<this>;
    worksheet(worksheetName: string): IWorksheet;
    worksheetNames(): string[];
}
export interface IWorkbookConstructor {
    new (filePath: string): IWorkbook;
}
export declare class Workbook implements IWorkbook {
    private filePath;
    private workbook;
    private worksheets;
    constructor(filePath: string);
    convertJson(jsonData: IJsonRecord[], headings?: string[]): IWorksheetData;
    data(): IWorksheetData[];
    instantiateWorkbook(constructor: IWorkbookConstructor, filePath: string): IWorkbook;
    initialize(): Promise<this>;
    update(wsName: string, wsData: IWorksheetData): Promise<this>;
    worksheet(wsName: string): IWorksheet;
    worksheetNames(): string[];
    private getWsData;
    private instantiateWorksheet;
    private findWorksheet;
}
//# sourceMappingURL=workbook.d.ts.map