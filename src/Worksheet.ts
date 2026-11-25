import type jspreadsheet from './jspreadsheet';
import { Clipboard } from './components/Clipboard.js';
import { Editor } from './components/Editor.js';
import { Formulas } from './components/Formulas.js';
import { History } from './components/History.js';
import { Pointer } from './components/Pointer.js';
import { Selection } from './components/Selection.js';
import { View } from './components/View.js';
import type { Spreadsheet, WorksheetState } from './types/internal';

/**
 * The fields of a worksheet and its public methods. The fields are in WorksheetState (src/types/internal.d.ts). The
 * public methods are functions in the modules under src/utils, which src/utils/worksheets.ts puts on the prototype.
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
        this.history = new History();
        this.clipboard = new Clipboard();
        this.editor = new Editor();
        this.selection = new Selection();
        this.formulas = new Formulas();
        this.view = new View();
        this.pointer = new Pointer();
        this.lastScrollLeft = 0;
    }
}
