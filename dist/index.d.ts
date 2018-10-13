interface IWorkbook {
    commit: () => Promise<this>;
    getData: (wsName: string) => any[];
    getHeadings: (wsName: string) => any[];
    getSheetNames: () => string[];
    init: () => Promise<this>;
    update: (wsName: string, jsonData: any[]) => void;
}
declare class Workbook implements IWorkbook {
    private filePath;
    private workbook;
    constructor(filePath?: string);
    init(): Promise<this>;
    getSheetNames(): string[];
    getData(wsName: string): any[];
    commit(): Promise<this>;
    update(wsName: string, jsonData: any[]): void;
    getHeadings(wsName: string): any[];
    private getAoaData;
}
export = Workbook;
//# sourceMappingURL=index.d.ts.map