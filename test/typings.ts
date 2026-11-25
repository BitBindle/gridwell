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

// getSelection returns null when no cell is selected
const selection = worksheet.getSelection();
assertType<Equals<typeof selection, [number, number, number, number] | null>>();

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

// setComments removes the comment of a cell when the comment is a falsy value, null included
worksheet.setComments('A1', null);
worksheet.setComments({ A1: 'Comment', B1: null, C1: undefined, D1: '' });

declare const shadowRoot: ShadowRoot;

// root, plugins and style are spreadsheet options. A number in a worksheet style is the index of a spreadsheet style.
const pluginFactory = function (spreadsheetInstance: jspreadsheet.SpreadsheetInstance, options: object, spreadsheetConfig: jspreadsheet.SpreadsheetOptions) {
    return {
        getConfig: () => spreadsheetConfig,
        onevent: () => {},
        persistence: () => {},
        contextMenu: (instance: jspreadsheet.WorksheetInstance, x: number | null, y: number | null, e: PointerEvent, items: jspreadsheet.ContextMenuItem[]) =>
            items,
    } as unknown as jspreadsheet.Plugin;
};
library(document.createElement('div'), {
    root: shadowRoot,
    plugins: { pluginFactory },
    style: ['background-color: orange;'],
    worksheets: [{ style: { A1: 0, B1: 'font-weight: bold;' } }],
});

// A context menu line has no title
const contextMenuItems: jspreadsheet.ContextMenuItem[] = [{ title: 'Copy', onclick: () => {} }, { type: 'line' }];

// insertBefore is read as true or false, so both a boolean and a number work
worksheet.insertRow(1, 0, true);
worksheet.insertColumn(1, 0, 1);

// A worksheet name that the namespace does not have gives undefined
const worksheetByName = library.getWorksheetInstanceByName('Countries', 'namespace');
assertType<Equals<typeof worksheetByName, jspreadsheet.WorksheetInstance | Record<string, jspreadsheet.WorksheetInstance> | undefined>>();

// The selection coordinates are numbers
assertType<Equals<typeof worksheet.selectedCell, undefined | null | [number, number, number, number]>>();
