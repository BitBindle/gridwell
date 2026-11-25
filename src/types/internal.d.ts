/**
 * Types the library code uses internally. They are not published: the public typings are in src/jspreadsheet.d.ts.
 *
 * In a JavaScript module, import them with a JSDoc tag:
 *
 *     /** @import { Worksheet } from '../types/internal' *\/
 *
 * and give a function that uses the worksheet as `this` the tag `@this {Worksheet}`.
 */
import type jspreadsheet from '../jspreadsheet';

/**
 * A worksheet as the library sees it: the public instance and the state that only the library uses.
 */
export interface Worksheet extends jspreadsheet.WorksheetInstance {
    parent: Spreadsheet;

    options: Omit<jspreadsheet.WorksheetOptions, 'data' | 'columns'> & {
        /** setData changes rows that are objects to arrays, so the library code always gets arrays. */
        data?: jspreadsheet.CellValue[][];
        /**
         * Column is a union by type, and the code reads the options of one type after a test of the type that the
         * type check cannot follow. The index signature allows that.
         */
        columns?: (jspreadsheet.Column & Record<string, any>)[];
    };

    /** The public typings allow text, but the library only stores numbers. */
    selectedCell: undefined | null | [number, number, number, number];

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

    /** True while createWorksheet adds a tab, so that the tab does not create a second worksheet. */
    creationThroughJss?: boolean;

    /** Returns the spreadsheet settings. The factory sets it, but the public typings do not declare it yet. */
    getConfig: () => jspreadsheet.SpreadsheetOptions;
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
