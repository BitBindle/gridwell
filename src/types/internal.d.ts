/**
 * Types the library code uses internally. They are not published: the public typings are in src/jspreadsheet.d.ts.
 *
 * Import them with `import type { Worksheet } from '../types/internal';` and give a function that uses the worksheet
 * as `this` a `this: Worksheet` parameter.
 */
import type jspreadsheet from '../jspreadsheet';
import type { Tabs } from '../utils/tabs';

/**
 * A worksheet as the library sees it: the public instance and the state that only the library uses.
 */
export interface Worksheet extends jspreadsheet.WorksheetInstance {
    parent: Spreadsheet;

    /** The events find the worksheet of an element through jssWorksheet. */
    element: jspreadsheet.JworksheetInstanceElement & { jssWorksheet: Worksheet };

    options: Omit<jspreadsheet.WorksheetOptions, 'data' | 'columns' | 'minDimensions'> & {
        /**
         * setData changes rows that are objects to arrays, so the library code always gets arrays. setData also sets
         * an empty array when there is no data.
         */
        data: jspreadsheet.CellValue[][];
        /**
         * Column is a union by type, and the code reads the options of one type after a test of the type that the
         * type check cannot follow. The index signature allows that. The worksheet is set up with an empty array when
         * there are no columns.
         */
        columns: (jspreadsheet.Column & Record<string, any>)[];
        /** The worksheet is set up with [0, 0] when the option is missing. */
        minDimensions: [number, number];
    };

    /** The public typings allow any record. The library stores the records in WorksheetHistoryRecord. */
    history: WorksheetHistoryRecord[];

    /** The cell that has the roving tabindex (the tab stop of the grid). */
    activeCell?: HTMLTableCellElement | null;

    /** Whether the keyboard focus is in the grid or in its editor. */
    focusInGrid?: boolean;

    /** The select for the number of rows per page, when the worksheet has pagination options. */
    paginationDropdown?: HTMLSelectElement;

    /** True while a paste updates many cells, so that formula references are updated once at the end. */
    skipUpdateTableReferences?: boolean;

    /** The footer of the table, when the worksheet has footers. */
    tfoot?: HTMLTableSectionElement;

    /** True while the mouse drags the selection corner. */
    selectedCorner?: boolean;

    /** The column whose header the mouse selects, or null or false when there is none. It is text from data-x. */
    selectedHeader?: any;

    /** The row whose number the mouse selects, or null or false when there is none. It is text from data-y. */
    selectedRow?: any;

    /**
     * The public typings use unions of the row and column information. The code tests one property and then reads the
     * others, which the type check cannot follow, so the library sees these as any.
     */
    resizing: any;
    dragging: any;
}

/**
 * A spreadsheet as the library sees it.
 */
export interface Spreadsheet extends jspreadsheet.SpreadsheetInstance {
    /** A spreadsheet has no parent. Code that gets a worksheet or a spreadsheet checks this property. */
    parent?: undefined;

    worksheets: Worksheet[];

    /** The tabs of the worksheets, which createTabs keeps on the element. */
    element: jspreadsheet.JspreadsheetInstanceElement & { tabs?: Tabs };

    /** The jsuites context menu, which jsuites keeps on its element. */
    contextMenu: HTMLDivElement & {
        contextmenu: {
            open(e: MouseEvent, items: jspreadsheet.ContextMenuItem[]): void;
            /** jsuites takes no argument. The factory passes false, which jsuites ignores. */
            close(ignored?: boolean): void;
        };
    };

    /** The jsuites toolbar, which jsuites keeps on its element. update calls updateState of each item with the worksheet. */
    toolbar: HTMLDivElement & { toolbar: { update(worksheet: Worksheet): void } };

    /** True while createWorksheet adds a tab, so that the tab does not create a second worksheet. */
    creationThroughJss?: boolean;
}

/**
 * The type of a function that implements a public worksheet method, with the worksheet as `this`.
 *
 *     export const getData: WorksheetMethod<'getData'> = function (highlighted, processed) { ... };
 *
 * The parameters get their types from WorksheetInstance in src/jspreadsheet.d.ts, so the code and the public typings agree.
 */
/** A merged range in options.mergeCells: the colspan, the rowspan and, after setMerge, the cells that the range hides. */
export type MergeInfo = [number, number, HTMLTableCellElement[]?];

export type WorksheetMethod<K extends keyof jspreadsheet.WorksheetInstance> = jspreadsheet.WorksheetInstance[K] extends (...args: infer A) => infer R
    ? (this: Worksheet, ...args: A) => R
    : never;

/**
 * The jspreadsheet function, with the state the library keeps on it.
 */
export interface Library extends jspreadsheet.JSpreadsheet {
    /** The worksheet that gets the keyboard and clipboard events. */
    current: Worksheet | null;

    spreadsheet: Spreadsheet[];
}

/** A worksheet that has a selection. The functions that move the selection assume one. */
export interface SelectedWorksheet extends Worksheet {
    selectedCell: [number, number, number, number];
}

/**
 * A DOM event whose target the code reads as an element. The handlers are on the document and on the elements of the
 * spreadsheet, and the code reads the target without a check, as it did in JavaScript.
 */
export type TargetEvent<E extends Event = Event> = E & { target: HTMLElement };

/** A cell that a setValue record of the history changed. */
export type HistoryCellChange = {
    x: number;
    y: number;
    col: number;
    row: number;
    /** Missing when the cell did not change, for example because it is read-only. */
    value?: jspreadsheet.CellValue;
    oldValue?: jspreadsheet.CellValue;
};

/**
 * A record of the undo history, by its action. Undo and redo read the fields of the action. Type aliases, not
 * interfaces, so that a record is a public HistoryRecord too.
 */
export type WorksheetHistoryRecord =
    | {
          action: 'insertColumn' | 'deleteColumn';
          columnNumber: number;
          numOfColumns: number;
          /** insertColumn keeps its argument. deleteColumn stores 1, so that undo inserts the columns before columnNumber. */
          insertBefore: boolean | number | undefined;
          /** deleteColumn stores undefined when the worksheet has no column options, which setup always gives it. */
          columns: Worksheet['options']['columns'] | undefined;
          headers: Worksheet['headers'];
          cols: Worksheet['cols'];
          records: Worksheet['records'];
          footers: string[][];
          data: jspreadsheet.CellValue[][];
      }
    | {
          action: 'insertRow' | 'deleteRow';
          rowNumber: number;
          numOfRows: number;
          /** insertRow keeps its argument. deleteRow stores 1, so that undo inserts the rows before rowNumber. */
          insertBefore: boolean | number | undefined;
          rowRecords: Worksheet['records'];
          rowData: jspreadsheet.CellValue[][];
          rowNode: Worksheet['rows'];
      }
    | { action: 'moveColumn' | 'moveRow'; oldValue: number; newValue: number }
    | { action: 'setWidth'; column: number | number[]; oldValue: number | number[]; newValue: number | number[] }
    | {
          action: 'setHeight';
          row: number;
          /** The height attribute of the row as text, or the measured height when the row has no attribute. */
          oldValue: number | string;
          newValue: number;
      }
    | { action: 'setHeader'; column: number; oldValue: string; newValue: string }
    | { action: 'setComments'; oldValue: Record<string, string | null>; newValue: Record<string, string | null | undefined> }
    | {
          action: 'setValue';
          records: HistoryCellChange[];
          selection: Worksheet['selectedCell'];
          /** The style that a paste set and the style it replaced. A paste stores both. */
          newStyle?: Record<string, string>;
          oldStyle?: Record<string, string>;
      }
    | { action: 'setMerge'; column: string; colspan: number; rowspan: number; data: jspreadsheet.CellValue[] }
    | { action: 'orderBy'; rows: number[]; column: number; order: boolean | number }
    /** The styles by cell name, each as 'property:value' entries joined with ';' */
    | { action: 'setStyle'; oldValue: Record<string, string>; newValue: Record<string, string> };
