import jSuites from 'jsuites';

import libraryBase from './utils/libraryBase.js';

import Factory from './utils/factory.js';
import { destroyEvents } from './utils/events.js';

import * as helpers from './utils/helpers.js';
import dispatch from './utils/dispatch.js';
import version from './utils/version.js';
import type { Library, Spreadsheet, Worksheet } from './types/internal';

libraryBase.jspreadsheet = function (el: HTMLDivElement | HTMLTableElement, options: jspreadsheet.SpreadsheetOptions) {
    try {
        const worksheets: Worksheet[] = [];

        // Create spreadsheet
        Factory.spreadsheet(el, options, worksheets).then((spreadsheet) => {
            libraryBase.jspreadsheet.spreadsheet.push(spreadsheet);

            // Global onload event
            dispatch.call(spreadsheet, 'onload', spreadsheet);
        });

        return worksheets;
    } catch (e) {
        console.error(e);
    }
    // The other properties are added below. Factory.spreadsheet is async, so it never throws here and the function always returns the worksheets.
} as unknown as Library;

libraryBase.jspreadsheet.getWorksheetInstanceByName = function (worksheetName: string | null | undefined, namespace: string) {
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

// Define dictionary
libraryBase.jspreadsheet.setDictionary = function (o: Record<string, string>) {
    jSuites.setDictionary(o);
};

// destroy empties the spreadsheet property of the element
libraryBase.jspreadsheet.destroy = function (
    element: Omit<jspreadsheet.JspreadsheetInstanceElement, 'spreadsheet'> & { spreadsheet: jspreadsheet.SpreadsheetInstance | null },
    destroyEventHandlers?: boolean
) {
    if (element.spreadsheet) {
        const spreadsheetIndex = libraryBase.jspreadsheet.spreadsheet.indexOf(element.spreadsheet as Spreadsheet);
        libraryBase.jspreadsheet.spreadsheet.splice(spreadsheetIndex, 1);

        const root = element.spreadsheet.config.root || document;

        element.spreadsheet = null;
        element.innerHTML = '';

        if (destroyEventHandlers) {
            destroyEvents(root);
        }
    }
};

libraryBase.jspreadsheet.destroyAll = function () {
    for (let spreadsheetIndex = 0; spreadsheetIndex < libraryBase.jspreadsheet.spreadsheet.length; spreadsheetIndex++) {
        const spreadsheet = libraryBase.jspreadsheet.spreadsheet[spreadsheetIndex];

        libraryBase.jspreadsheet.destroy(spreadsheet.element);
    }
};

libraryBase.jspreadsheet.current = null;

libraryBase.jspreadsheet.spreadsheet = [];

libraryBase.jspreadsheet.helpers = {} as jspreadsheet.JssHelpers;

libraryBase.jspreadsheet.version = function () {
    return version;
};

Object.entries(helpers).forEach(([key, value]) => {
    (libraryBase.jspreadsheet.helpers as unknown as Record<string, unknown>)[key] = value;
});

export default libraryBase.jspreadsheet;
