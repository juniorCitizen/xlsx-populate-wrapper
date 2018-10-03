export interface IJsonRecord {
    [index: string]: any;
}
export interface IWorksheetData {
    aoaData: any[];
    headings: string[];
    jsonData: IJsonRecord[];
}
export interface IWorksheet {
    data(): IWorksheetData;
    headings(): string[];
    name(): string;
    update(worksheetData: IWorksheetData): void;
}
export interface IWorksheetConstructor {
    new (worksheet: any): IWorksheet;
}
export declare class Worksheet implements IWorksheet {
    private worksheet;
    constructor(worksheet: any);
    name(): string;
    data(): IWorksheetData;
    headings(): string[];
    update(worksheetData: IWorksheetData): void;
}
//# sourceMappingURL=worksheet.d.ts.map