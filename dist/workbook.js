"use strict";
var __awaiter = (this && this.__awaiter) || function (thisArg, _arguments, P, generator) {
    return new (P || (P = Promise))(function (resolve, reject) {
        function fulfilled(value) { try { step(generator.next(value)); } catch (e) { reject(e); } }
        function rejected(value) { try { step(generator["throw"](value)); } catch (e) { reject(e); } }
        function step(result) { result.done ? resolve(result.value) : new P(function (resolve) { resolve(result.value); }).then(fulfilled, rejected); }
        step((generator = generator.apply(thisArg, _arguments || [])).next());
    });
};
var __generator = (this && this.__generator) || function (thisArg, body) {
    var _ = { label: 0, sent: function() { if (t[0] & 1) throw t[1]; return t[1]; }, trys: [], ops: [] }, f, y, t, g;
    return g = { next: verb(0), "throw": verb(1), "return": verb(2) }, typeof Symbol === "function" && (g[Symbol.iterator] = function() { return this; }), g;
    function verb(n) { return function (v) { return step([n, v]); }; }
    function step(op) {
        if (f) throw new TypeError("Generator is already executing.");
        while (_) try {
            if (f = 1, y && (t = op[0] & 2 ? y["return"] : op[0] ? y["throw"] || ((t = y["return"]) && t.call(y), 0) : y.next) && !(t = t.call(y, op[1])).done) return t;
            if (y = 0, t) op = [op[0] & 2, t.value];
            switch (op[0]) {
                case 0: case 1: t = op; break;
                case 4: _.label++; return { value: op[1], done: false };
                case 5: _.label++; y = op[1]; op = [0]; continue;
                case 7: op = _.ops.pop(); _.trys.pop(); continue;
                default:
                    if (!(t = _.trys, t = t.length > 0 && t[t.length - 1]) && (op[0] === 6 || op[0] === 2)) { _ = 0; continue; }
                    if (op[0] === 3 && (!t || (op[1] > t[0] && op[1] < t[3]))) { _.label = op[1]; break; }
                    if (op[0] === 6 && _.label < t[1]) { _.label = t[1]; t = op; break; }
                    if (t && _.label < t[2]) { _.label = t[2]; _.ops.push(op); break; }
                    if (t[2]) _.ops.pop();
                    _.trys.pop(); continue;
            }
            op = body.call(thisArg, _);
        } catch (e) { op = [6, e]; y = 0; } finally { f = t = 0; }
        if (op[0] & 5) throw op[1]; return { value: op[0] ? op[1] : void 0, done: true };
    }
};
Object.defineProperty(exports, "__esModule", { value: true });
var XlsxPopulate = require("xlsx-populate");
var worksheet_1 = require("./worksheet");
var Workbook = (function () {
    function Workbook(filePath) {
        this.filePath = filePath;
        this.workbook = null;
        this.worksheets = [];
    }
    Workbook.prototype.convertJson = function (jsonData, headings) {
        if (headings === void 0) { headings = []; }
        if (headings.length === 0) {
            headings = Object.keys(jsonData[0]);
        }
        jsonData = jsonData.map(sanitizeJsonRecord(headings));
        var aoaData = jsonData.map(function (jsonRecord) {
            var reduceFn = function (arrayRecord, heading) {
                arrayRecord.push(jsonRecord[heading]);
                return arrayRecord;
            };
            return headings.reduce(reduceFn, []);
        });
        return { headings: headings, aoaData: aoaData, jsonData: jsonData };
    };
    Workbook.prototype.data = function () {
        if (!this.workbook) {
            throw new Error('workbook is not ready');
        }
        return this.worksheets.map(this.getWsData);
    };
    Workbook.prototype.instantiateWorkbook = function (constructor, filePath) {
        return new constructor(filePath);
    };
    Workbook.prototype.initialize = function () {
        return __awaiter(this, void 0, void 0, function () {
            var _a, error_1;
            var _this = this;
            return __generator(this, function (_b) {
                switch (_b.label) {
                    case 0:
                        _b.trys.push([0, 2, , 3]);
                        _a = this;
                        return [4, XlsxPopulate.fromFileAsync(this.filePath)];
                    case 1:
                        _a.workbook = _b.sent();
                        this.worksheets = this.workbook
                            .sheets()
                            .reduce(function (worksheets, worksheet) {
                            worksheets.push(_this.instantiateWorksheet(worksheet_1.Worksheet, worksheet));
                            return worksheets;
                        }, []);
                        return [2, this];
                    case 2:
                        error_1 = _b.sent();
                        this.filePath = '';
                        this.workbook = null;
                        this.worksheets = [];
                        throw error_1;
                    case 3: return [2];
                }
            });
        });
    };
    Workbook.prototype.update = function (wsName, wsData) {
        return __awaiter(this, void 0, void 0, function () {
            var headings, jsonData, error_2;
            var _this = this;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        if (!this.workbook) {
                            throw new Error('workbook is not ready');
                        }
                        headings = wsData.headings, jsonData = wsData.jsonData;
                        wsData = this.convertJson(jsonData, headings);
                        this.worksheet(wsName).update(wsData);
                        _a.label = 1;
                    case 1:
                        _a.trys.push([1, 3, , 4]);
                        return [4, this.workbook.toFileAsync(this.filePath)];
                    case 2:
                        _a.sent();
                        return [2, this];
                    case 3:
                        error_2 = _a.sent();
                        this.worksheets = this.workbook
                            .sheets()
                            .reduce(function (worksheets, worksheet) {
                            worksheets.push(_this.instantiateWorksheet(worksheet_1.Worksheet, worksheet));
                            return worksheets;
                        }, []);
                        throw error_2;
                    case 4: return [2];
                }
            });
        });
    };
    Workbook.prototype.worksheet = function (wsName) {
        if (!this.workbook) {
            throw new Error('workbook is not ready');
        }
        var searchFn = this.worksheetNames();
        var wsIndex = searchFn.findIndex(this.findWorksheet(wsName));
        if (wsIndex !== -1) {
            return this.worksheets[wsIndex];
        }
        else {
            throw new Error("worksheet " + wsName + "} not found");
        }
    };
    Workbook.prototype.worksheetNames = function () {
        if (!this.workbook) {
            throw new Error('workbook is not ready');
        }
        return this.worksheets.map(function (worksheet) { return worksheet.name(); });
    };
    Workbook.prototype.getWsData = function (ws) {
        return ws.data();
    };
    Workbook.prototype.instantiateWorksheet = function (constructor, worksheet) {
        return new constructor(worksheet);
    };
    Workbook.prototype.findWorksheet = function (worksheetName) {
        return function (wsName) { return wsName === worksheetName; };
    };
    return Workbook;
}());
exports.Workbook = Workbook;
function sanitizeJsonRecord(properties) {
    return function (jsonRecord) {
        var originalRecord = JSON.parse(JSON.stringify(jsonRecord));
        return properties.reduce(sanitize(originalRecord), {});
    };
}
function sanitize(originalRecord) {
    return function (sanitized, property) {
        var isValid = originalRecord.hasOwnProperty(property);
        sanitized[property] = isValid ? originalRecord[property] : null;
        return sanitized;
    };
}
//# sourceMappingURL=workbook.js.map