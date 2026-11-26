// Checks the public typings in src/gridwell.d.ts against what the library returns. npm run typecheck checks this
// file; mocha does not run it. assertType fails the check when a type is not exactly the expected type.
import type gridwell from '../src/gridwell';

type Equals<A, B> = (<T>() => T extends A ? 1 : 2) extends <T>() => T extends B ? 1 : 2 ? true : false;
declare function assertType<T extends true>(): void;

declare const worksheet: gridwell.WorksheetInstance;

// getData returns rows of values, text with a delimiter, or objects with asJson
const data = worksheet.getData();
assertType<Equals<typeof data, gridwell.CellValue[][]>>();
const highlightedData = worksheet.getData(true, true);
assertType<Equals<typeof highlightedData, gridwell.CellValue[][]>>();
const csv = worksheet.getData(false, false, ';');
assertType<Equals<typeof csv, string>>();
const json = worksheet.getData(false, false, undefined, true);
assertType<Equals<typeof json, Record<number, gridwell.CellValue>[]>>();

// getSelected returns the records of the selected cells, or their names
const selected = worksheet.getSelected();
assertType<Equals<typeof selected, { element: HTMLTableCellElement; x: number; y: number }[]>>();
const selectedNames = worksheet.getSelected(true);
assertType<Equals<typeof selectedNames, string[]>>();

// getSelection returns null when no cell is selected
const selection = worksheet.getSelection();
assertType<Equals<typeof selection, [number, number, number, number] | null>>();

// getMeta returns the meta information of a cell, or null, or the meta information of all cells
const cellMeta = worksheet.getMeta('A1');
assertType<Equals<typeof cellMeta, gridwell.MetaInformation | null>>();
const allMeta = worksheet.getMeta();
assertType<Equals<typeof allMeta, Record<string, gridwell.MetaInformation> | undefined>>();

// resetStyle takes the new styles by cell name, as setStyle does
type ResetStyles = Parameters<gridwell.WorksheetInstance['resetStyle']>[0];
assertType<Equals<ResetStyles, Record<string, string | string[]>>>();

declare const spreadsheet: gridwell.SpreadsheetInstance;

// The factory gives each spreadsheet getConfig
const config = spreadsheet.getConfig();
assertType<Equals<typeof config, gridwell.SpreadsheetOptions>>();

declare const library: gridwell.Gridwell;

// invert gives the keys by value
const inverted = library.helpers.invert({ a: 0, b: 'x' });
assertType<Equals<typeof inverted, string[] & Record<string, string>>>();
assertType<Equals<Parameters<typeof library.helpers.getCaretIndex>[0], Node>>();

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
function pluginFactory(spreadsheetInstance: gridwell.SpreadsheetInstance, options: object, spreadsheetConfig: gridwell.SpreadsheetOptions) {
    return {
        getConfig: () => spreadsheetConfig,
        onevent: () => {},
        persistence: () => {},
        contextMenu: (instance: gridwell.WorksheetInstance, x: number | null, y: number | null, e: PointerEvent, items: gridwell.ContextMenuItem[]) => items,
    } as unknown as gridwell.Plugin;
}
library(document.createElement('div'), {
    root: shadowRoot,
    plugins: { pluginFactory },
    style: ['background-color: orange;'],
    worksheets: [{ style: { A1: 0, B1: 'font-weight: bold;' } }],
});

// A plugin needs none of its methods
library(document.createElement('div'), { plugins: { empty: () => ({}) }, worksheets: [{}] });

// The about option can be the text that the "About" item shows
library(document.createElement('div'), { about: 'My application', worksheets: [{}] });

// A context menu line has no title
const contextMenuItems: gridwell.ContextMenuItem[] = [{ title: 'Copy', onclick: () => {} }, { type: 'line' }];

// insertBefore is read as true or false, so both a boolean and a number work
worksheet.insertRow(1, 0, true);
worksheet.insertColumn(1, 0, 1);

// A worksheet name that the namespace does not have gives undefined
const worksheetByName = library.getWorksheetInstanceByName('Countries', 'namespace');
assertType<Equals<typeof worksheetByName, gridwell.WorksheetInstance | Record<string, gridwell.WorksheetInstance> | undefined>>();

// The selection coordinates are numbers, and the range is null when nothing is selected
assertType<Equals<typeof worksheet.selection.range, null | [number, number, number, number]>>();
// @ts-expect-error The range is selection.range
void worksheet.selectedCell;
// @ts-expect-error The selected cells are selection.highlighted
void worksheet.highlighted;

// Without a direction, orderBy toggles the direction of the column
worksheet.orderBy(0);

// The indexes of a drag or a resize are numbers
assertType<Equals<gridwell.DragColumnInfo['column'], number>>();
assertType<Equals<gridwell.DragRowInfo['row'], number>>();
assertType<Equals<gridwell.ResizeColumnInfo['column'], number>>();
assertType<Equals<gridwell.ResizeRowInfo['row'], number>>();

type Handler<E extends keyof gridwell.SpreadsheetOptions> = Parameters<Extract<gridwell.SpreadsheetOptions[E], (...args: never[]) => unknown>>;

// onpaste gets one item for each pasted cell
assertType<Equals<Handler<'onpaste'>[1], { x: number; y: number; value: gridwell.CellValue }[]>>();

// onbeforepaste can return the rows to paste, of { value } objects as it got them or of values
assertType<
    Equals<
        ReturnType<Extract<gridwell.SpreadsheetOptions['onbeforepaste'], (...args: never[]) => unknown>>,
        undefined | boolean | string | ({ value: gridwell.CellValue } | gridwell.CellValue)[][]
    >
>();

// oneditionend gets undefined as the value when the edit was not saved
assertType<Equals<Handler<'oneditionend'>[4], gridwell.CellValue | undefined>>();

// setValue without a value gives onbeforechange and onchange undefined as the new value
assertType<Equals<Handler<'onbeforechange'>[4], gridwell.CellValue | undefined>>();
assertType<Equals<Handler<'onchange'>[4], gridwell.CellValue | undefined>>();

// The change events get the column and row as numbers
assertType<Equals<Handler<'onbeforechange'>[2], number>>();
assertType<Equals<Handler<'onchange'>[3], number>>();
assertType<Equals<gridwell.CellChange['x'], number>>();
assertType<Equals<gridwell.CellChange['y'], number>>();

// oncomments gets the comments as setComments got them, where undefined removes a comment
assertType<Equals<Handler<'oncomments'>[1], Record<string, string | null | undefined>>>();

// onupdateresult gets the shown rows after a search or a filter
assertType<Equals<Handler<'onupdateresult'>, [instance: gridwell.WorksheetInstance, results: number[] | null]>>();

// The persistence of a plugin gets the changes as onafterchanges gets them
assertType<Equals<Parameters<NonNullable<gridwell.Plugin['persistence']>>[2], { data: gridwell.CellChange[] }>>();

// onbeforesave gets the values of a row by column name, or by index for a column without a name
assertType<Equals<Handler<'onbeforesave'>[2], { row: number; data: Record<string, gridwell.CellValue> }[]>>();

// A toolbar function gets the settings of the default toolbar, with the items, and returns settings
assertType<Equals<Parameters<Extract<gridwell.SpreadsheetOptions['toolbar'], (...args: never[]) => unknown>>[0], gridwell.ToolbarOptions>>();
assertType<Equals<ReturnType<Extract<gridwell.SpreadsheetOptions['toolbar'], (...args: never[]) => unknown>>, gridwell.ToolbarOptions>>();
assertType<Equals<Parameters<NonNullable<gridwell.Plugin['toolbar']>>[0], gridwell.ToolbarOptions>>();

// A select item can do without an icon
const fontItem: gridwell.ToolbarItem = { type: 'select', options: ['Arial', 'Verdana'] };

// The value of a select item is the index of the option it shows first
const alignItem: gridwell.ToolbarItem = { type: 'select', options: ['top', 'middle', 'bottom'], value: 1 };

// The undo history is an object with the records, the position and the flag
assertType<Equals<typeof worksheet.history.records, gridwell.HistoryRecord[]>>();
assertType<Equals<ReturnType<typeof worksheet.history.canUndo>, boolean>>();
assertType<Equals<ReturnType<typeof worksheet.history.back>, gridwell.HistoryRecord | undefined>>();
// @ts-expect-error The position is in history.index
void worksheet.historyIndex;
// @ts-expect-error The flag is history.ignore
void worksheet.ignoreHistory;

// What a worksheet copied is in its clipboard
assertType<Equals<typeof worksheet.clipboard.data, string>>();
assertType<Equals<typeof worksheet.clipboard.textarea, HTMLTextAreaElement>>();
// @ts-expect-error The copied text is clipboard.data
void worksheet.data;
// @ts-expect-error The textarea is clipboard.textarea
void worksheet.textarea;

// The cell editor keeps the cell that is being edited
assertType<Equals<typeof worksheet.editor.cell, HTMLTableCellElement | null>>();
if (worksheet.editor.isOpen()) {
    assertType<Equals<typeof worksheet.editor.cell, HTMLTableCellElement>>();
}
// @ts-expect-error The cell is editor.cell
void worksheet.edition;

// The formula chain is in formulas.chain
assertType<Equals<typeof worksheet.formulas.chain, Record<string, string[]>>>();
// @ts-expect-error The chain is formulas.chain
void worksheet.formula;

// The rows that the worksheet shows are in view
assertType<Equals<typeof worksheet.view.results, number[] | null>>();
assertType<Equals<typeof worksheet.view.pageNumber, number>>();
// @ts-expect-error The results are view.results
void worksheet.results;
// @ts-expect-error The page is view.pageNumber
void worksheet.pageNumber;

// The state of a mouse gesture is in pointer, and the pressed header and row number are in selection
assertType<Equals<typeof worksheet.pointer.corner, boolean>>();
assertType<Equals<typeof worksheet.pointer.cursor, HTMLElement | null>>();
assertType<Equals<typeof worksheet.selection.column, number | null>>();
assertType<Equals<typeof worksheet.selection.row, number | null>>();
// @ts-expect-error The resize is pointer.resizing
void worksheet.resizing;
// @ts-expect-error The move is pointer.dragging
void worksheet.dragging;
// @ts-expect-error The element is pointer.cursor
void worksheet.cursor;

// The fields that only the library uses are not public
// @ts-expect-error The col tags are internal
void worksheet.cols;
// @ts-expect-error The colgroup is internal
void worksheet.colgroupContainer;
// @ts-expect-error The selection corner is internal
void worksheet.corner;
// @ts-expect-error The header row is internal; the header cells are headers
void worksheet.headerContainer;
// @ts-expect-error The link below the table is internal
void worksheet.ads;

// The values that the application or a formula gives have the type unknown, not any
assertType<Equals<gridwell.MetaInformation[string], unknown>>();
assertType<Equals<gridwell.HistoryRecord['records'], unknown>>();
const formulaResult = worksheet.executeFormula('=1+1');
assertType<Equals<typeof formulaResult, unknown>>();
type SaveData = Parameters<NonNullable<gridwell.SpreadsheetOptions['onsave']>>[2];
assertType<Equals<SaveData, unknown>>();

// dispatch takes the arguments of the handler of the event
const dispatched = worksheet.dispatch('onselection', worksheet, 0, 0, 1, 1, undefined);
assertType<Equals<typeof dispatched, unknown>>();
// @ts-expect-error onselection gets the worksheet first
worksheet.dispatch('onselection', 0);
// @ts-expect-error The event must be an event of SpreadsheetOptions
worksheet.dispatch('oncustom');

// onevent gets the name of the event, then its arguments. A handler can name fewer parameters than an event gives (issue #133)
const shortOptions: gridwell.SpreadsheetOptions = {
    worksheets: [],
    onevent: (event) => {
        assertType<Equals<typeof event, gridwell.EventName>>();
    },
    plugins: {
        test: () => ({
            onevent(event, instance) {
                assertType<Equals<typeof event, gridwell.EventName>>();
                // The second argument of the events: a worksheet, or the spreadsheet
                assertType<Equals<typeof instance, gridwell.WorksheetInstance | gridwell.SpreadsheetInstance>>();
            },
        }),
    },
};
void shortOptions;
// Taken as EventArguments, a test of the name in args[0] narrows the arguments
const eventOptions: gridwell.SpreadsheetOptions = {
    worksheets: [],
    onevent(...args: gridwell.EventArguments) {
        if (args[0] === 'onselection') {
            assertType<Equals<(typeof args)[1], gridwell.WorksheetInstance>>();
            assertType<Equals<(typeof args)[2], number>>();
        }
    },
};
void eventOptions;
// A parameter that the shortest events do not give is optional
const optionalOptions: gridwell.SpreadsheetOptions = { worksheets: [], onevent: (event, instance, cell?) => void [event, instance, cell] };
void optionalOptions;
// @ts-expect-error onblur gives only the worksheet, so a third parameter must be optional
const requiredOptions: gridwell.SpreadsheetOptions = { worksheets: [], onevent: (event, instance, cell) => void [event, instance, cell] };
void requiredOptions;
// A handler that takes any arguments still fits, but not one that expects a worksheet for every event
function looseOnevent(event: string, ...rest: unknown[]) {
    return [event, rest];
}
const looseOptions: gridwell.SpreadsheetOptions = { worksheets: [], onevent: looseOnevent };
void looseOptions;
function worksheetOnevent(event: string, instance: gridwell.WorksheetInstance) {
    return [event, instance];
}
// @ts-expect-error onbeforesave and others give the spreadsheet first
const worksheetOptions: gridwell.SpreadsheetOptions = { worksheets: [], onevent: worksheetOnevent };
void worksheetOptions;

// A context menu item gets the element of the item and the release of the mouse or the key, and the toolbar items get
// the toolbar and the picker
type ContextMenuClick = Parameters<NonNullable<gridwell.ContextMenuOption['onclick']>>;
assertType<Equals<ContextMenuClick, [element: HTMLDivElement, e: MouseEvent | KeyboardEvent]>>();
// A handler that an application wrote for a mouse event still fits
const mouseItem: gridwell.ContextMenuOption = { title: 'Item', onclick: (element: HTMLDivElement, e: MouseEvent) => void [element, e] };
void mouseItem;
type ToolbarClick = Parameters<NonNullable<gridwell.ToolbarIconItem['onclick']>>[1];
assertType<Equals<ToolbarClick, gridwell.ToolbarInstance>>();
declare const toolbarInstance: gridwell.ToolbarInstance;
toolbarInstance.update();
// @ts-expect-error The toolbar does not move the items that do not fit into a menu
toolbarInstance.refresh();
declare const pickerInstance: gridwell.PickerInstance;
pickerInstance.close();
assertType<Equals<ReturnType<typeof pickerInstance.isOpened>, boolean>>();
// render gets the element of the option in the list, or null for the option that the picker shows, and can give an element
type PickerRender = NonNullable<gridwell.ToolbarSelectItem['render']>;
assertType<Equals<Parameters<PickerRender>, [option: string, element: HTMLElement | null]>>();
assertType<Equals<ReturnType<PickerRender>, string | Element>>();

// The library has the static methods that it declares, and no others
// @ts-expect-error The library has no index signature
void library.custom;
// An event of the application's own needs a cast of its name, as UPGRADING.md says
worksheet.dispatch('oncustom' as gridwell.EventName, worksheet);

// getDataFromRange gives the data without processed, and the shown values with it (issue #99)
const rangeData = worksheet.getDataFromRange('A1:C1');
assertType<Equals<typeof rangeData, gridwell.CellValue[][]>>();
worksheet.getDataFromRange('A1:C1', false);
worksheet.getDataFromRange('A1:C1', true);

// An empty delimiter counts as none, so getData gives rows (issue #39)
const emptyDelimiterData = worksheet.getData(false, false, '');
assertType<Equals<typeof emptyDelimiterData, gridwell.CellValue[][]>>();
const emptyDelimiterJson = worksheet.getData(false, false, '', true);
assertType<Equals<typeof emptyDelimiterJson, Record<number, gridwell.CellValue>[]>>();

// A column can have a format, a locale with the options of Intl.NumberFormat, allowEmpty, the filter and the delimiter of a
// dropdown, and the valid range and the first day of the week of a calendar
const columns: gridwell.Column[] = [
    { type: 'numeric', format: '0.0' },
    { type: 'numeric', locale: 'en-US', options: { style: 'currency', currency: 'USD' }, allowEmpty: true },
    { type: 'dropdown', source: ['a'], filter: (element, cell, x, y, source) => source ?? [] },
    { type: 'dropdown', source: ['a', 'b'], multiple: true, delimiter: ', ' },
    { type: 'calendar', options: { format: 'DD/MM/YYYY', validRange: ['2024-01-01', null], startingDay: 1 } },
];

// setConfig takes the options to change: of the worksheet, or of the spreadsheet with spreadsheetLevel
worksheet.setConfig({ columnSorting: false });
worksheet.setConfig({ tabs: true }, true);
// @ts-expect-error -- tabs is an option of the spreadsheet
worksheet.setConfig({ tabs: true });
worksheet.parent.setConfig({ toolbar: true });

// A footer can show numbers
library(document.createElement('div'), { worksheets: [{ footers: [['Total', 3, '=SUM(A1:A2)']] }] });

// A table needs no options and no worksheets, because the first worksheet is created from the table
declare const table: HTMLTableElement;
library(table);
library(table, { tabs: true });
assertType<Equals<ReturnType<typeof library.helpers.createFromTable>, gridwell.WorksheetOptions>>();
library.helpers.createFromTable(table);
// @ts-expect-error -- a div needs the options with the worksheets
library(document.createElement('div'));

// The cells that onbeforepaste gets have the pasted text
library(document.createElement('div'), {
    worksheets: [{}],
    onbeforepaste: (instance, copiedText) => copiedText.map((row) => row.map((cell) => cell.value.trim())),
});

// setFormulaEngine takes the engine of @jspreadsheet/formula, or null to remove it
library.setFormulaEngine(null);
type Engine = Parameters<gridwell.Gridwell['setFormulaEngine']>[0];
assertType<Equals<Engine, gridwell.FormulaEngine | null>>();
