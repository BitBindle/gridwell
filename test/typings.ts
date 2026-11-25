// Checks the public typings in src/jspreadsheet.d.ts against what the library returns. npm run typecheck checks this
// file; mocha does not run it. assertType fails the check when a type is not exactly the expected type.
import type jspreadsheet from '../src/jspreadsheet';

type Equals<A, B> = (<T>() => T extends A ? 1 : 2) extends <T>() => T extends B ? 1 : 2 ? true : false;
declare function assertType<T extends true>(): void;

declare const worksheet: jspreadsheet.WorksheetInstance;

// getData returns rows of values, text with a delimiter, or objects with asJson
const data = worksheet.getData();
assertType<Equals<typeof data, jspreadsheet.CellValue[][]>>();
const highlightedData = worksheet.getData(true, true);
assertType<Equals<typeof highlightedData, jspreadsheet.CellValue[][]>>();
const csv = worksheet.getData(false, false, ';');
assertType<Equals<typeof csv, string>>();
const json = worksheet.getData(false, false, undefined, true);
assertType<Equals<typeof json, Record<number, jspreadsheet.CellValue>[]>>();
