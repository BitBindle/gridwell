import type jspreadsheet from './jspreadsheet';
import { Cells } from './components/Cells.js';
import { Clipboard } from './components/Clipboard.js';
import { Columns } from './components/Columns.js';
import { Comments } from './components/Comments.js';
import { Data } from './components/Data.js';
import { Editor } from './components/Editor.js';
import { Formulas } from './components/Formulas.js';
import { History } from './components/History.js';
import { Merges } from './components/Merges.js';
import { Meta } from './components/Meta.js';
import { Pointer } from './components/Pointer.js';
import { Rows } from './components/Rows.js';
import { Selection } from './components/Selection.js';
import { Styles } from './components/Styles.js';
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
        this.lastScrollLeft = 0;
    }
}
