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

// getSelected returns the records of the selected cells, or their names
const selected = worksheet.getSelected();
assertType<Equals<typeof selected, { element: HTMLTableCellElement; x: number; y: number }[]>>();
const selectedNames = worksheet.getSelected(true);
assertType<Equals<typeof selectedNames, string[]>>();

declare const spreadsheet: jspreadsheet.SpreadsheetInstance;

// The factory gives each spreadsheet getConfig
const config = spreadsheet.getConfig();
assertType<Equals<typeof config, jspreadsheet.SpreadsheetOptions>>();

declare const library: jspreadsheet.JSpreadsheet;

// version returns the version information, which has no type
const versionInfo = library.version();
assertType<Equals<typeof versionInfo.version, string>>();
// @ts-expect-error The version information has no type property
const versionType = versionInfo.type;
