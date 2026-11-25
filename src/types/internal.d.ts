/**
 * Types the library code uses internally. They are not published: the public typings are in src/index.d.ts.
 *
 * In a JavaScript module, import them with a JSDoc tag:
 *
 *     /** @import { Worksheet } from '../types/internal' *\/
 *
 * and give a function that uses the worksheet as `this` the tag `@this {Worksheet}`.
 */
import type jspreadsheet from '../index';

/**
 * A worksheet as the library sees it: the public instance and the state that only the library uses.
 */
export interface Worksheet extends jspreadsheet.WorksheetInstance {
    parent: Spreadsheet;

    /** setData changes rows that are objects to arrays, so the library code always gets arrays. */
    options: Omit<jspreadsheet.WorksheetOptions, 'data'> & { data?: jspreadsheet.CellValue[][] };

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
}

/**
 * The type of a function that implements a public worksheet method, with the worksheet as `this`.
 *
 *     export const getData: WorksheetMethod<'getData'> = function (highlighted, processed) { ... };
 *
 * The parameters get their types from WorksheetInstance in src/index.d.ts, so the code and the public typings agree.
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
