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
function gridwell(el: HTMLDivElement | HTMLTableElement, options: gridwell.SpreadsheetOptions) {
    // A table is replaced by a div in its parent. Without a parent, the spreadsheet could not be reached, and the error
    // would be lost in the promise of createSpreadsheet.
    if (el.tagName === 'TABLE' && !el.parentNode) {
        throw new Error('Gridwell: the table must be in a parent element, which gets the spreadsheet in its place');
    }

    const worksheets: Worksheet[] = [];

    // Create spreadsheet
    createSpreadsheet(el, options, worksheets).then((spreadsheet) => {
        libraryBase.gridwell.spreadsheet.push(spreadsheet);

        // Global onload event
        spreadsheet.dispatch('onload', spreadsheet);
    });

    return worksheets;
}

// The other properties are added below
libraryBase.gridwell = gridwell as unknown as Library;

libraryBase.gridwell.getWorksheetInstanceByName = (worksheetName: string | null | undefined, namespace: string) => {
    const targetSpreadsheet = libraryBase.gridwell.spreadsheet.find((spreadsheet) => {
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

libraryBase.gridwell.setFormulaEngine = (engine: gridwell.FormulaEngine | null) => {
    if (engine !== null && typeof engine !== 'function') {
        throw new TypeError('Gridwell: the formula engine must be a function, such as the default export of @jspreadsheet/formula');
    }

    libraryBase.formulaEngine = engine;
};

// Define dictionary
libraryBase.gridwell.setDictionary = (o: Record<string, string>) => {
    setDictionary(o);
};

// destroy empties the spreadsheet property of the element
libraryBase.gridwell.destroy = (
    element: Omit<gridwell.GridwellInstanceElement, 'spreadsheet'> & { spreadsheet: gridwell.SpreadsheetInstance | null },
    destroyEventHandlers?: boolean
) => {
    if (element.spreadsheet) {
        const spreadsheetIndex = libraryBase.gridwell.spreadsheet.indexOf(element.spreadsheet as Spreadsheet);
        libraryBase.gridwell.spreadsheet.splice(spreadsheetIndex, 1);

        const root = element.spreadsheet.config.root || document;

        element.spreadsheet = null;
        element.innerHTML = '';

        if (destroyEventHandlers) {
            destroyEvents(root);
        }
    }
};

libraryBase.gridwell.destroyAll = () => {
    // destroy removes the spreadsheet from the list, so the loop goes over a copy
    for (const spreadsheet of [...libraryBase.gridwell.spreadsheet]) {
        libraryBase.gridwell.destroy(spreadsheet.element);
    }
};

libraryBase.gridwell.current = null;

libraryBase.gridwell.spreadsheet = [];

libraryBase.gridwell.helpers = {} as gridwell.JssHelpers;

libraryBase.gridwell.version = () => version;

for (const [key, value] of Object.entries(helpers)) {
    (libraryBase.gridwell.helpers as unknown as Record<string, unknown>)[key] = value;
}

export default libraryBase.gridwell;
