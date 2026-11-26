import jSuites from 'jsuites';
import type jspreadsheet from './jspreadsheet';
import { Accessibility } from './components/Accessibility.js';
import { Cells } from './components/Cells.js';
import { Clipboard } from './components/Clipboard.js';
import { Columns } from './components/Columns.js';
import { Comments } from './components/Comments.js';
import { Data } from './components/Data.js';
import { Editor } from './components/Editor.js';
import { Formulas } from './components/Formulas.js';
import { History } from './components/History.js';
import { Keyboard } from './components/Keyboard.js';
import { Merges } from './components/Merges.js';
import { Meta } from './components/Meta.js';
import { Pointer } from './components/Pointer.js';
import { Rows } from './components/Rows.js';
import { Selection } from './components/Selection.js';
import { Styles } from './components/Styles.js';
import { View } from './components/View.js';
import type { EventHandlers, EventName, Spreadsheet, WorksheetState } from './types/internal';

/** The rows that persistence posts to the server, the data of onbeforesave */
type SaveData = Parameters<NonNullable<jspreadsheet.SpreadsheetOptions['onbeforesave']>>[2];

/** A record of onafterchanges. The coordinates can be text (issue #52), and a read-only cell has no value (issue #58). */
type ChangeRecord = { x: number; y: number; value?: jspreadsheet.CellValue };

/**
 * The fields of a worksheet and its public methods. The fields are in WorksheetState (src/types/internal.d.ts). The
 * public methods are methods of the class, or functions in the modules under src/utils, which src/utils/worksheets.ts
 * puts on the prototype.
 */
// eslint-disable-next-line @typescript-eslint/no-empty-object-type, @typescript-eslint/no-unsafe-declaration-merging
export interface Worksheet extends WorksheetState {}

/**
 * A worksheet: one table of a spreadsheet, with its options, its DOM and its state.
 *
 * The constructor sets up the state that the worksheet has before it is built. buildWorksheet (src/utils/worksheets.ts)
 * then creates the DOM from the options.
 */
// eslint-disable-next-line @typescript-eslint/no-unsafe-declaration-merging
export class Worksheet {
    constructor(spreadsheet: Spreadsheet, options: jspreadsheet.WorksheetOptions, element?: HTMLElement) {
        this.parent = spreadsheet;
        // buildWorksheet fills in the options that the internal type says are always there
        this.options = options as Worksheet['options'];
        if (element) {
            // buildWorksheet adds jssWorksheet
            this.element = element as Worksheet['element'];
        }
        this.history = new History(this);
        this.clipboard = new Clipboard(this);
        this.editor = new Editor(this);
        this.selection = new Selection(this);
        this.formulas = new Formulas(this);
        this.view = new View(this);
        this.pointer = new Pointer(this);
        this.cells = new Cells(this);
        this.styles = new Styles(this);
        this.comments = new Comments(this);
        this.meta = new Meta(this);
        this.merges = new Merges(this);
        this.columns = new Columns(this);
        this.rowManager = new Rows(this);
        this.data = new Data(this);
        this.accessibility = new Accessibility(this);
        this.keyboard = new Keyboard(this);
        this.lastScrollLeft = 0;
    }

    /**
     * Trigger an event of the worksheet. The handlers get the worksheet as `this`. After onafterchanges, the plugins
     * and the persistence option save the changes.
     *
     * The first argument is the name of the event, for example "onchange". The other arguments go to the event handlers.
     * It returns the result of the last handler, or null when no handler ran.
     */
    dispatch<E extends EventName>(event: E, ...args: Parameters<EventHandlers[E]>): ReturnType<EventHandlers[E]> | null {
        const worksheet = this;
        const spreadsheet = worksheet.parent;

        const ret = spreadsheet.callHandlers(worksheet, event, args);

        if (event == 'onafterchanges') {
            if (typeof spreadsheet.plugins === 'object') {
                Object.entries(spreadsheet.plugins).forEach(function ([, plugin]) {
                    if (typeof plugin.persistence === 'function') {
                        // The changes of onafterchanges
                        plugin.persistence(worksheet, 'setValue', { data: args[1] as jspreadsheet.CellChange[] });
                    }
                });
            }

            if (worksheet.options.persistence) {
                const url = worksheet.options.persistence == true ? worksheet.options.url : worksheet.options.persistence;
                const data = this.prepareJson(args[1] as ChangeRecord[]);
                this.save(url, data);
            }
        }

        // The result of a handler of the event, as the type of dispatch says
        return ret as ReturnType<EventHandlers[E]> | null;
    }

    /**
     * Prepare JSON in the correct format
     * @param data - The changed cells.
     */
    private prepareJson(data: ChangeRecord[]) {
        const obj = this;

        const rows: SaveData = [];
        for (let i = 0; i < data.length; i++) {
            const x = data[i].x;
            const y = data[i].y;
            const k = obj.options.columns[x].name ? obj.options.columns[x].name : x;

            // Create row
            if (!rows[y]) {
                rows[y] = {
                    row: y,
                    data: {},
                };
            }
            // A read-only cell has no value (issue #58)
            rows[y].data[k] = data[i].value!;
        }

        // Filter rows
        return rows.filter(function (el) {
            return el != null;
        });
    }

    /**
     * Post json to a remote server
     */
    // The url is undefined when persistence is true and the worksheet has no url option
    private save(url: string | undefined, data: SaveData) {
        const obj = this;

        // Parse anything in the data before sending to the server
        const ret = obj.parent.dispatch('onbeforesave', obj.parent, obj, data);
        if (ret) {
            // The application can return any data to send in place of the rows
            data = ret as SaveData;
        } else {
            if (ret === false) {
                return false;
            }
        }

        // Remove update
        jSuites.ajax({
            url: url,
            method: 'POST',
            dataType: 'json',
            data: { data: JSON.stringify(data) },
            success: function (result: unknown) {
                // Event
                obj.dispatch('onsave', obj.parent, obj, data);
            },
        });
    }

    /**
     * Get the options of the worksheet.
     */
    getConfig() {
        const obj = this;

        return obj.options;
    }

    /**
     * Change options of the worksheet, or of its spreadsheet when spreadsheetLevel is true.
     * @param config - The options to change.
     * @param spreadsheetLevel - Change the options of the spreadsheet of the worksheet.
     */
    setConfig(config: jspreadsheet.SpreadsheetOptions | jspreadsheet.WorksheetOptions, spreadsheetLevel?: boolean) {
        const obj = this;

        if (spreadsheetLevel) {
            obj.parent.setConfig(config);
            return;
        }

        const keys = Object.keys(config);

        keys.forEach(function (key) {
            (obj.options as Record<string, unknown>)[key] = (config as Record<string, unknown>)[key];
        });
    }

    /**
     * The index of the worksheet that the spreadsheet shows.
     */
    getWorksheetActive(): number {
        return this.parent.getWorksheetActive();
    }
}
