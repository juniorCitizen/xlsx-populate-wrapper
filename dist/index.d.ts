declare class Workbook {
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