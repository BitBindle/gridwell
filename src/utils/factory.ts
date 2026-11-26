import { createFromTable } from './helpers.js';
import type jspreadsheet from '../jspreadsheet';
import { Spreadsheet } from '../Spreadsheet.js';
import type { Worksheet } from '../types/internal';

/**
 * Create a spreadsheet in an element and build it. A table is replaced by a div. The div gets the spreadsheet in its
 * spreadsheet property.
 */
export async function createSpreadsheet(el: HTMLElement & { spreadsheet?: Spreadsheet }, options: jspreadsheet.SpreadsheetOptions, worksheets: Worksheet[]) {
    if (el.tagName == 'TABLE') {
        if (!options) {
            // The worksheets are added below
            options = {} as jspreadsheet.SpreadsheetOptions;
        }

        if (!options.worksheets) {
            options.worksheets = [];
        }

        const tableOptions = createFromTable(el, options.worksheets[0]);

        // el is a table, so the options were created
        options.worksheets[0] = tableOptions!;

        const div = document.createElement('div');
        // A table outside the document has no parent, and this throws (issue #55)
        el.parentNode!.insertBefore(div, el);
        el.remove();
        el = div;
    }

    // el gets the spreadsheet at the end
    const spreadsheet = new Spreadsheet(el, options, worksheets);

    await spreadsheet.build();

    return spreadsheet;
}
