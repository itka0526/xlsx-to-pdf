import { readFile, utils } from "xlsx";
import PDFDocument from "pdfkit-table";
import { createWriteStream } from "fs";

const RESULT_PATH = "./test_files/";

type COLUMN_NAME = string;
type ROW_VALUE = string;

type COLUMN_NAME_ROW_VALUE = { [k: COLUMN_NAME]: ROW_VALUE };

type JSON_DATA_FORMAT = COLUMN_NAME_ROW_VALUE[];
type PDF_TABLE_FORMAT = {
    headers: COLUMN_NAME[];
    rows: ROW_VALUE[][];
};

async function XLSX_PDF(path: string) {
    const workBook = readFile(path);

    for (const name of workBook.SheetNames) {
        const sheet = workBook.Sheets[name];

        const jsonData: JSON_DATA_FORMAT = utils.sheet_to_json(sheet, {}) ?? [];

        const pdfDoc = new PDFDocument({ margin: 30, size: "A4" });

        pdfDoc.pipe(createWriteStream(RESULT_PATH + "./result.pdf"));

        await createTable(pdfDoc, jsonData);
    }
}

function transformer(data: JSON_DATA_FORMAT): PDF_TABLE_FORMAT {
    if (data.length === 0) {
        return { headers: [], rows: [] };
    }

    let headers: COLUMN_NAME[] = [],
        rows: ROW_VALUE[][] = [];

    for (const key of Object.keys(data[0])) headers.push(key);

    for (let i = 0; i < data.length; i++) {
        if (!data[i]) continue;

        rows.push(Object.values(data[i]).map((v) => String(v)));
    }

    return { headers, rows };
}

async function createTable(doc: PDFDocument, data: JSON_DATA_FORMAT) {
    const table = transformer(data);

    await doc.table(table, {
        prepareHeader: () => doc.font("./Roboto/Roboto-Regular.ttf", 12),
        prepareRow: () => doc.font("./Roboto/Roboto-Regular.ttf", 12),
    });

    doc.end();
}

export { XLSX_PDF };
