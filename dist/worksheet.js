"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
var Worksheet = (function () {
    function Worksheet(worksheet) {
        this.worksheet = worksheet;
    }
    Worksheet.prototype.name = function () {
        return this.worksheet.name();
    };
    Worksheet.prototype.data = function () {
        var aoaData = this.worksheet.usedRange().value();
        var headings = aoaData.shift().map(mapToString);
        var jsonData = aoaData.map(generateMapCallback(headings));
        return { headings: headings, aoaData: aoaData, jsonData: jsonData };
    };
    Worksheet.prototype.headings = function () {
        return this.worksheet
            .usedRange()
            .value()
            .shift()
            .map(mapToString);
    };
    Worksheet.prototype.update = function (worksheetData) {
        this.worksheet.usedRange().clear();
        this.worksheet.cell('A1').value([worksheetData.headings]);
        this.worksheet.cell('A2').value(worksheetData.aoaData);
    };
    return Worksheet;
}());
exports.Worksheet = Worksheet;
function mapToString(value) {
    return value.toString();
}
function generateMapCallback(headings) {
    return function (aoaRecord) {
        return headings.reduce(generateReduceCallback(aoaRecord), {});
    };
}
function generateReduceCallback(aoaRecord) {
    return function (jsonRecord, heading, index) {
        jsonRecord[heading] = aoaRecord[index];
        return jsonRecord;
    };
}
//# sourceMappingURL=worksheet.js.map