import jSuites from 'jsuites';

import libraryBase from './utils/libraryBase.js';

import Factory from './utils/factory.js';
import { destroyEvents } from './utils/events.js';

import * as helpers from './utils/helpers.js';
import dispatch from './utils/dispatch.js';
import version from './utils/version.js';
import type { Library } from './types/internal';

libraryBase.jspreadsheet = function (el: any, options: any) {
    try {
        const worksheets: any[] = [];

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
} as Library;

libraryBase.jspreadsheet.getWorksheetInstanceByName = function (worksheetName: any, namespace: any): any {
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
libraryBase.jspreadsheet.setDictionary = function (o: any) {
    jSuites.setDictionary(o);
};

libraryBase.jspreadsheet.destroy = function (element: any, destroyEventHandlers?: boolean) {
    if (element.spreadsheet) {
        const spreadsheetIndex = libraryBase.jspreadsheet.spreadsheet.indexOf(element.spreadsheet);
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

libraryBase.jspreadsheet.helpers = {} as any;

libraryBase.jspreadsheet.version = function () {
    return version;
};

Object.entries(helpers).forEach(([key, value]) => {
    (libraryBase.jspreadsheet.helpers as any)[key] = value;
});

export default libraryBase.jspreadsheet;
