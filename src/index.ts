import libraryBase from './utils/libraryBase.js';
import { setDictionary } from './utils/dictionary.js';

import { createSpreadsheet } from './utils/factory.js';
import { destroyEvents } from './events/index.js';

import * as helpers from './utils/helpers.js';
import version from './utils/version.js';
import type { Library, Spreadsheet, Worksheet } from './types/internal';

/**
 * Create a spreadsheet in an element. It returns the worksheets, which the spreadsheet builds after it returns.
 */
function jspreadsheet(el: HTMLDivElement | HTMLTableElement, options: jspreadsheet.SpreadsheetOptions) {
    // A table is replaced by a div in its parent. Without a parent, the spreadsheet could not be reached, and the error
    // would be lost in the promise of createSpreadsheet.
    if (el.tagName === 'TABLE' && !el.parentNode) {
        throw new Error('Jspreadsheet: the table must be in a parent element, which gets the spreadsheet in its place');
    }

    const worksheets: Worksheet[] = [];

    // Create spreadsheet
    createSpreadsheet(el, options, worksheets).then((spreadsheet) => {
        libraryBase.jspreadsheet.spreadsheet.push(spreadsheet);

        // Global onload event
        spreadsheet.dispatch('onload', spreadsheet);
    });

    return worksheets;
}

// The other properties are added below
libraryBase.jspreadsheet = jspreadsheet as unknown as Library;

libraryBase.jspreadsheet.getWorksheetInstanceByName = (worksheetName: string | null | undefined, namespace: string) => {
    const targetSpreadsheet = libraryBase.jspreadsheet.spreadsheet.find((spreadsheet) => {
        return spreadsheet.config.namespace === namespace;
    });

    if (!targetSpreadsheet) {
        return {};
    }

    if (typeof worksheetName === 'undefined' || worksheetName === null) {
        const namespaceEntries = targetSpreadsheet.worksheets.map((worksheet) => {
            return [worksheet.options.worksheetName, worksheet];
        });

        return Object.fromEntries(namespaceEntries);
    }

    return targetSpreadsheet.worksheets.find((worksheet) => {
        return worksheet.options.worksheetName === worksheetName;
    });
};

libraryBase.jspreadsheet.setFormulaEngine = (engine: jspreadsheet.FormulaEngine | null) => {
    if (engine !== null && typeof engine !== 'function') {
        throw new TypeError('Jspreadsheet: the formula engine must be a function, such as the default export of @jspreadsheet/formula');
    }

    libraryBase.formulaEngine = engine;
};

// Define dictionary
libraryBase.jspreadsheet.setDictionary = (o: Record<string, string>) => {
    setDictionary(o);
};

// destroy empties the spreadsheet property of the element
libraryBase.jspreadsheet.destroy = (
    element: Omit<jspreadsheet.JspreadsheetInstanceElement, 'spreadsheet'> & { spreadsheet: jspreadsheet.SpreadsheetInstance | null },
    destroyEventHandlers?: boolean
) => {
    if (element.spreadsheet) {
        const spreadsheetIndex = libraryBase.jspreadsheet.spreadsheet.indexOf(element.spreadsheet as Spreadsheet);
        libraryBase.jspreadsheet.spreadsheet.splice(spreadsheetIndex, 1);

        const root = element.spreadsheet.config.root || document;

        // The active worksheet would keep the destroyed spreadsheet in memory. A worksheet that deleteWorksheet removed
        // is not in the list of worksheets, but its parent is still the spreadsheet.
        if (libraryBase.jspreadsheet.current?.parent === element.spreadsheet) {
            libraryBase.jspreadsheet.current = null;

            // A timer of a search, a column rename or a touch hold would read the active worksheet
            if (libraryBase.jspreadsheet.timeControl) {
                clearTimeout(libraryBase.jspreadsheet.timeControl);
                libraryBase.jspreadsheet.timeControl = null;
            }
        }

        // An open context menu listens to the document, and its items hold the spreadsheet
        (element.spreadsheet as Spreadsheet).contextMenu.contextmenu?.close();

        // The size observer of the frozen rows holds the worksheet
        for (const worksheet of (element.spreadsheet as Spreadsheet).worksheets) {
            worksheet.rowManager.stopFreeze();
            worksheet.rowSource?.stop();
        }

        element.spreadsheet = null;
        element.innerHTML = '';

        if (destroyEventHandlers) {
            destroyEvents(root);
        }
    }
};

libraryBase.jspreadsheet.destroyAll = () => {
    // destroy removes the spreadsheet from the list, so the loop goes over a copy
    for (const spreadsheet of [...libraryBase.jspreadsheet.spreadsheet]) {
        libraryBase.jspreadsheet.destroy(spreadsheet.element);
    }
};

libraryBase.jspreadsheet.current = null;

libraryBase.jspreadsheet.spreadsheet = [];

libraryBase.jspreadsheet.helpers = {} as jspreadsheet.JssHelpers;

libraryBase.jspreadsheet.version = () => version;

for (const [key, value] of Object.entries(helpers)) {
    (libraryBase.jspreadsheet.helpers as unknown as Record<string, unknown>)[key] = value;
}

export default libraryBase.jspreadsheet;
