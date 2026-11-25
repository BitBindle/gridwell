import type jspreadsheet from './jspreadsheet';
import type { SpreadsheetState, Worksheet } from './types/internal';

/**
 * The fields of a spreadsheet and its public methods. The fields are in SpreadsheetState (src/types/internal.d.ts). The
 * public methods are functions in the modules under src/utils, which src/utils/factory.ts puts on the prototype.
 */
// eslint-disable-next-line @typescript-eslint/no-empty-object-type, @typescript-eslint/no-unsafe-declaration-merging
export interface Spreadsheet extends SpreadsheetState {}

/**
 * A spreadsheet: the worksheets in their tabs, the toolbar, the context menu and the plugins.
 *
 * The constructor sets up the state that the spreadsheet has before it is built. factory.spreadsheet
 * (src/utils/factory.ts) then creates the worksheets, the toolbar and the context menu.
 */
// eslint-disable-next-line @typescript-eslint/no-unsafe-declaration-merging
export class Spreadsheet {
    constructor(element: HTMLElement, config: jspreadsheet.SpreadsheetOptions, worksheets: Worksheet[]) {
        // jspreadsheet() returns this array before the worksheets are built, and the factory fills it
        this.worksheets = worksheets;
        this.config = config;
        // The element gets the spreadsheet in its spreadsheet property when the spreadsheet is built
        this.element = element as Spreadsheet['element'];
        this.el = this.element;

        // The context menu container. jSuites.contextmenu adds its instance when the spreadsheet is built.
        this.contextMenu = document.createElement('div') as Spreadsheet['contextMenu'];
        this.contextMenu.className = 'jss_contextmenu';
    }
}
