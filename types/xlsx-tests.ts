import XLSX = require('xlsx');

const options: XLSX.ParsingOptions = {
    cellDates: true
};

const workbook: XLSX.WorkBook = XLSX.readFile('test.xlsx', options);
const otherworkbook: XLSX.WorkBook = XLSX.readFile('test.xlsx', {type: 'file'});

const author: string = workbook.Props.Author;

const firstsheet: string = workbook.SheetNames[0];
const firstworksheet: XLSX.WorkSheet = workbook.Sheets[firstsheet];
const WB1A1: XLSX.CellObject = (firstworksheet["A1"]);

interface Tester {
    name: string;
    age: number;
}

const jsonvalues: Tester[] = XLSX.utils.sheet_to_json<Tester>(firstworksheet);
const csv: string = XLSX.utils.sheet_to_csv(firstworksheet);
const txt: string = XLSX.utils.sheet_to_txt(firstworksheet);
const formulae: string[] = XLSX.utils.sheet_to_formulae(firstworksheet);
const aoa: any[][] = XLSX.utils.sheet_to_json<any[]>(firstworksheet, {raw:true, header:1});

const aoa2: XLSX.WorkSheet = XLSX.utils.aoa_to_sheet<number>([
    [1,2,3,4,5,6,7],
    [2,3,4,5,6,7,8]
]);

const js2ws: XLSX.WorkSheet = XLSX.utils.json_to_sheet<Tester>([
    {name:"Sheet", age: 12},
    {name:"JS", age: 24}
]);

const WBProps = workbook.Workbook;
const WBSheets = WBProps.Sheets;
const WBSheet0 = WBSheets[0];
console.log(WBSheet0.Hidden);

const fmt14 = XLSX.SSF.get_table()[14];
XLSX.SSF.load('"This is a custom format "0.000');

const newwb = XLSX.utils.book_new();
XLSX.utils.book_append_sheet(newwb, aoa2, "AOA");
XLSX.utils.book_append_sheet(newwb, js2ws, "JSON");
const bstrxlsx: string = XLSX.write(newwb, {type: "binary", bookType: "xlsx" });

const wb_1: XLSX.WorkBook = XLSX.read(XLSX.write(newwb, {type: "base64", bookType: "xlsx" }), {type: "base64"});
const wb_2: XLSX.WorkBook = XLSX.read(XLSX.write(newwb, {type: "binary", bookType: "xlsx" }), {type: "binary"});
const wb_3: XLSX.WorkBook = XLSX.read(XLSX.write(newwb, {type: "buffer", bookType: "xlsx" }), {type: "buffer"});
const wb_4: XLSX.WorkBook = XLSX.read(XLSX.write(newwb, {type: "file", bookType: "xlsx" }), {type: "file"});
const wb_5: XLSX.WorkBook = XLSX.read(XLSX.write(newwb, {type: "array", bookType: "xlsx" }), {type: "array"});
const wb_6: XLSX.WorkBook = XLSX.read(XLSX.write(newwb, {type: "string", bookType: "xlsx" }), {type: "string"});

function get_header_row(sheet: XLSX.WorkSheet) {
    const headers: string[] = [];
    const range = XLSX.utils.decode_range(sheet['!ref']);
    const R: number = range.s.r;
    for(let C = range.s.c; C <= range.e.c; ++C) {
        const cell: XLSX.CellObject = sheet[XLSX.utils.encode_cell({c:C, r:R})];
        let hdr = "UNKNOWN " + C;
        if(cell && cell.t) hdr = XLSX.utils.format_cell(cell);
        headers.push(hdr);
    }
    return headers;
}

const headers: string[] = get_header_row(aoa2);

const CFB = XLSX.CFB;
const vbawb = XLSX.readFile("test.xlsm", {bookVBA:true});
if(vbawb.vbaraw) {
    const cfb: any /* XLSX.CFB.CFB$Container */ = CFB.read(vbawb.vbaraw, {type: "buffer"});
}

// read
{
    XLSX.read(new Uint8Array(0));
    XLSX.read(new ArrayBuffer(0));

    XLSX.read(new Uint8Array(0), { type: 'buffer' });
    XLSX.read(new ArrayBuffer(0), { type: 'array' });

    XLSX.read('', { type: 'base64' });

    // throw at runtime as the path is empty, but type checking should succeed
    assertThrows(() => XLSX.read('', { type: 'file' }));
    XLSX.read('', { type: 'string' });
    XLSX.read('', { type: 'binary' });

    // @ts-expect-error plain object can't be parsed
    XLSX.read({});
    // @ts-expect-error array can't be parsed
    XLSX.read([]);
}

// write
{
    const wb = XLSX.read(new Uint8Array(0));

    const u8 = XLSX.write(wb, { type: 'buffer' });
    const arrayBuffer = XLSX.write(wb, { type: 'array' });
    const b64 = XLSX.write(wb, { type: 'base64' });

    assert(u8 instanceof Uint8Array);
    assert(arrayBuffer instanceof ArrayBuffer);
    assert(typeof b64 === 'string');

    // @ts-expect-error this will fail both at compile time and runtime without `opts.type`
    assertThrows(() => XLSX.write(wb));
    // @ts-expect-error this will fail both at compile time and runtime without `opts.type`
    assertThrows(() => XLSX.write(wb, {}));
}

// sheet_to_json (compile-time type checking only)
{
    const wb = XLSX.read(new Uint8Array(0));
    const ws = wb.Sheets[wb.SheetNames[0]];

    const _1 = XLSX.utils.sheet_to_json(ws, { header: 1 });
    _1[0] = [1, undefined, 'str', true];

    const _a = XLSX.utils.sheet_to_json(ws, { header: ['a', 'b', 'c', 'd'] });
    _a[0] = { a: 1, b: 2, c: 3, d: undefined };
    // @ts-expect-error unrecognized key `e`
    _a[0] = { a: 1, b: 2, c: 3, d: 4, e: 5 };
    // @ts-expect-error missing `d`
    _a[0] = { a: 1, b: 2, c: 3 };

    const header2 = ['a', 'b', 'c', 'd'] as const;
    const _a2 = XLSX.utils.sheet_to_json(ws, { header: header2 });
    _a2[0] = { a: 1, b: 2, c: 3, d: undefined };
    // @ts-expect-error unrecognized key `e`
    _a2[0] = { a: 1, b: 2, c: 3, d: 4, e: 5 };
    // @ts-expect-error missing `d`
    _a2[0] = { a: 1, b: 2, c: 3 };

    const header3 = ['a', 'b', 'c', 'd'];
    const _a3 = XLSX.utils.sheet_to_json(ws, { header: header3 });
    _a3[0] = { key: 'val' };
    // @ts-expect-error array not object
    _a3[0] = ['val'];

    const _A = XLSX.utils.sheet_to_json(ws, { header: 'A' });
    _A[0] = { key: 'val' };
    // @ts-expect-error array not object
    _A[0] = ['val'];

    const _u = XLSX.utils.sheet_to_json(ws, { header: undefined });
    _u[0] = { key: 'val' };
    // @ts-expect-error array not object
    _u[0] = ['val'];

    const _u2 = XLSX.utils.sheet_to_json(ws);
    _u2[0] = { key: 'val' };
    // @ts-expect-error array not object
    _u2[0] = ['val'];

    const _u3 = XLSX.utils.sheet_to_json(ws, {});
    _u3[0] = { key: 'val' };
    // @ts-expect-error array not object
    _u3[0] = ['val'];

    // manually supplying a type acts as a type assertion
    const _x = XLSX.utils.sheet_to_json<[string, number]>(ws);
    _x[0] = ['str', 1];
    // @ts-expect-error wrong column order
    _x[0] = [1, 'str'];

    const _defVal = XLSX.utils.sheet_to_json(ws, { defval: new Date(0) });
    _defVal[0] = { key: new Date(0), key2: 'str' };
    // @ts-ignore - this will only error when strict null checking enabled `undefined` not allowed
    _defVal[0] = { key: undefined };

    const _defVal2 = XLSX.utils.sheet_to_json(ws, { header: 1, defval: new Date(0) });
    _defVal2[0] = [new Date(0), 'str'];
    // @ts-ignore - this will only error when strict null checking enabled `undefined` not allowed
    _defVal2[0] = [new Date(0), 'str', undefined];

    const _defVal3 = XLSX.utils.sheet_to_json(ws, { header: ['a', 'b'], defval: new Date(0) });
    _defVal3[0] = { a: new Date(0), b: 'str' };
    // @ts-expect-error unrecognized key `c`
    _defVal3[0] = { a: new Date(0), b: 2, c: new Date(0) };
    // @ts-expect-error missing `b`
    _defVal3[0] = { a: 1 };
    // @ts-expect-error `undefined` not allowed
    _defVal3[0] = { a: undefined };
}

function assert(condition: boolean, message?: string) {
    if (!condition) {
        throw new Error(message ?? 'condition returned false');
    }
}

function assertThrows(fn: (() => any), message?: string) {
    let thrown = false;
    try {
        fn();
    } catch {
        thrown = true;
    }

    if (!thrown) throw new Error(message ?? 'function failed to throw');
}
