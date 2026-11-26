/**
 * Types the library code uses internally. They are not published: the public typings are in src/jspreadsheet.d.ts.
 *
 * Import them with `import type { Worksheet } from '../types/internal';` and give a function that uses the worksheet
 * as `this` a `this: Worksheet` parameter.
 */
import type jspreadsheet from '../jspreadsheet';
import type { Tabs } from '../utils/tabs';
import type { Spreadsheet } from '../Spreadsheet';
import type { Worksheet } from '../Worksheet';
import type { Cells } from '../components/Cells';
import type { Clipboard } from '../components/Clipboard';
import type { Comments } from '../components/Comments';
import type { Editor } from '../components/Editor';
import type { Formulas } from '../components/Formulas';
import type { History } from '../components/History';
import type { Merges } from '../components/Merges';
import type { Meta } from '../components/Meta';
import type { Pointer } from '../components/Pointer';
import type { Selection } from '../components/Selection';
import type { Styles } from '../components/Styles';
import type { View } from '../components/View';

/** The classes. Their fields are in WorksheetState and SpreadsheetState below. */
export type { Spreadsheet, Worksheet };

/**
 * The fields of a worksheet as the library sees it: the public instance and the state that only the library uses.
 */
export interface WorksheetState extends jspreadsheet.WorksheetInstance {
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
    history: History;

    clipboard: Clipboard;

    editor: Editor;

    selection: Selection;

    formulas: Formulas;

    view: View;

    pointer: Pointer;

    /** Finds the td of a cell and its label, and makes cells read-only. */
    cells: Cells;

    /** The styles of the cells. */
    styles: Styles;

    /** The comments of the cells. */
    comments: Comments;

    /** The meta information of the cells. */
    meta: Meta;

    /** The merged cells. */
    merges: Merges;

    /** The div that holds the link to the Jspreadsheet site, below the table. */
    ads: HTMLDivElement;

    /** The col tags of the table, one for each column, with the index of the column. */
    cols: {
        colElement: HTMLTableColElement;
        x: number;
    }[];

    /** The colgroup tag of the table, which holds the col tags. */
    colgroupContainer: HTMLElement;

    /** The square in the lower-right corner of the selection, which fills the cells it is dragged over. */
    corner: HTMLDivElement;

    /** The row of the thead that holds the headers of the columns. */
    headerContainer: HTMLTableRowElement;

    /** The select for the number of rows per page, when the worksheet has pagination options. */
    paginationDropdown?: HTMLSelectElement;

    /** True while a paste updates many cells, so that formula references are updated once at the end. */
    skipUpdateTableReferences?: boolean;

    /** The footer of the table, when the worksheet has footers. */
    tfoot?: HTMLTableSectionElement;

    /**
     * The scrollLeft of the content at the last scroll. A scroll places the frozen columns only when scrollLeft changed,
     * not when the content scrolled up or down.
     */
    lastScrollLeft: number;
}

/**
 * The fields of a spreadsheet as the library sees it.
 */
export interface SpreadsheetState extends jspreadsheet.SpreadsheetInstance {
    worksheets: Worksheet[];

    /** The tabs of the worksheets, which the Tabs constructor keeps on the element. */
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

/** A merged range in options.mergeCells: the colspan, the rowspan and, after setMerge, the cells that the range hides. */
export type MergeInfo = [number, number, HTMLTableCellElement[]?];

/**
 * The type of a function that implements a public worksheet method, with the worksheet as `this`.
 *
 *     export const getData: WorksheetMethod<'getData'> = function (highlighted, processed) { ... };
 *
 * The parameters get their types from WorksheetInstance in src/jspreadsheet.d.ts, so the code and the public typings agree.
 */
export type WorksheetMethod<K extends keyof jspreadsheet.WorksheetInstance> = jspreadsheet.WorksheetInstance[K] extends (...args: infer A) => infer R
    ? (this: Worksheet, ...args: A) => R
    : never;

/** The handlers of the events, by event name (see src/jspreadsheet.d.ts). */
export type EventHandlers = jspreadsheet.EventHandlers;

/** The name of an event that dispatch can trigger. */
export type EventName = jspreadsheet.EventName;

/**
 * dispatch, typed by event: the arguments are those of the handler that SpreadsheetOptions declares for the event, and
 * the result is what that handler returns. null means that no handler ran. onevent and the onevent of plugins can also
 * return any value, so a caller still checks what it gets.
 *
 * The code calls it as dispatch.call(worksheet, 'onchange', ...), so call is typed too.
 */
export interface Dispatch {
    <E extends EventName>(this: Worksheet | Spreadsheet, event: E, ...args: Parameters<EventHandlers[E]>): ReturnType<EventHandlers[E]> | null;
    call<E extends EventName>(thisArg: Worksheet | Spreadsheet, event: E, ...args: Parameters<EventHandlers[E]>): ReturnType<EventHandlers[E]> | null;
}

/**
 * The jspreadsheet function, with the state the library keeps on it.
 */
export interface Library extends jspreadsheet.JSpreadsheet {
    /** The worksheet that gets the keyboard and clipboard events. */
    current: Worksheet | null;

    spreadsheet: Spreadsheet[];

    /**
     * The timer of a long press on a touch screen, which opens the context menu, of a press on a selected header, which
     * renames the column, and of the search as the user types.
     */
    timeControl: null | number;

    /**
     * True from a press on a worksheet, while no editor is open, until the button is released. The mouse move and over
     * handlers resize, move or extend the selection only while it is true. Undefined before the first press.
     */
    isMouseAction?: boolean;

    /**
     * The cell of a long press on a touch screen, which opened its editor. When the touch ends, the input of the editor
     * gets the focus. Null for a color cell.
     */
    tmpElement?: HTMLElement | null;

    /** The timer that loads more rows when the content scrolls near its top or its end, with lazyLoading. */
    timeControlLoading: null | number;
}

/** A worksheet whose editor is open. The functions for the keys of an open editor assume one. */
export interface EditingWorksheet extends Worksheet {
    editor: Editor & { cell: HTMLTableCellElement };
}

/** A worksheet that has a selection. The functions that move the selection assume one. */
export interface SelectedWorksheet extends Worksheet {
    selection: Selection & { range: [number, number, number, number] };
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
          /** The height attribute of the row, or the measured height when the row has no attribute. */
          oldValue: number;
          newValue: number;
      }
    | { action: 'setHeader'; column: number; oldValue: string; newValue: string }
    | { action: 'setComments'; oldValue: Record<string, string | null>; newValue: Record<string, string | null | undefined> }
    | {
          action: 'setValue';
          records: HistoryCellChange[];
          selection: Selection['range'];
          /** The style that a paste set and the style it replaced. A paste stores both. */
          newStyle?: Record<string, string>;
          oldStyle?: Record<string, string>;
      }
    | { action: 'setMerge'; column: string; colspan: number; rowspan: number; data: jspreadsheet.CellValue[] }
    | { action: 'orderBy'; rows: number[]; column: number; order: boolean | number }
    /** The styles by cell name, each as 'property:value' entries joined with ';' */
    | { action: 'setStyle'; oldValue: Record<string, string>; newValue: Record<string, string> };
