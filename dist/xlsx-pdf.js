"use strict";
var __awaiter = (this && this.__awaiter) || function (thisArg, _arguments, P, generator) {
    function adopt(value) { return value instanceof P ? value : new P(function (resolve) { resolve(value); }); }
    return new (P || (P = Promise))(function (resolve, reject) {
        function fulfilled(value) { try { step(generator.next(value)); } catch (e) { reject(e); } }
        function rejected(value) { try { step(generator["throw"](value)); } catch (e) { reject(e); } }
        function step(result) { result.done ? resolve(result.value) : adopt(result.value).then(fulfilled, rejected); }
        step((generator = generator.apply(thisArg, _arguments || [])).next());
    });
};
var __importDefault = (this && this.__importDefault) || function (mod) {
    return (mod && mod.__esModule) ? mod : { "default": mod };
};
Object.defineProperty(exports, "__esModule", { value: true });
exports.XLSX_PDF = void 0;
const xlsx_1 = require("xlsx");
const pdfkit_table_1 = __importDefault(require("pdfkit-table"));
const fs_1 = require("fs");
const RESULT_PATH = "./test_files/";
function XLSX_PDF(path) {
    var _a;
    return __awaiter(this, void 0, void 0, function* () {
        const workBook = (0, xlsx_1.readFile)(path);
        for (const name of workBook.SheetNames) {
            const sheet = workBook.Sheets[name];
            const jsonData = (_a = xlsx_1.utils.sheet_to_json(sheet, {})) !== null && _a !== void 0 ? _a : [];
            const pdfDoc = new pdfkit_table_1.default({ margin: 30, size: "A4" });
            pdfDoc.pipe((0, fs_1.createWriteStream)(RESULT_PATH + "./result.pdf"));
            yield createTable(pdfDoc, jsonData);
        }
    });
}
exports.XLSX_PDF = XLSX_PDF;
function transformer(data) {
    if (data.length === 0) {
        return { headers: [], rows: [] };
    }
    let headers = [], rows = [];
    for (const key of Object.keys(data[0]))
        headers.push(key);
    for (let i = 0; i < data.length; i++) {
        if (!data[i])
            continue;
        rows.push(Object.values(data[i]).map((v) => String(v)));
    }
    return { headers, rows };
}
function createTable(doc, data) {
    return __awaiter(this, void 0, void 0, function* () {
        const table = transformer(data);
        yield doc.table(table, {
            prepareHeader: () => doc.font("./Roboto/Roboto-Regular.ttf", 12),
            prepareRow: () => doc.font("./Roboto/Roboto-Regular.ttf", 12),
        });
        doc.end();
    });
}
