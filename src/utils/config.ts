import type jspreadsheet from '../jspreadsheet';
import { Spreadsheet } from '../Spreadsheet.js';
import type { Worksheet } from '../types/internal';
/**
 * Get table config information
 */
export const getWorksheetConfig = function (this: Worksheet) {
    const obj = this;

    return obj.options;
};

export const getSpreadsheetConfig = function (this: Spreadsheet) {
    const spreadsheet = this;

    return spreadsheet.config;
};

/**
 * Change options of the worksheet, or of the spreadsheet when this is a spreadsheet or spreadsheetLevel is true.
 * @param config - The options to change.
 * @param spreadsheetLevel - Change the options of the spreadsheet of the worksheet.
 */
export const setConfig = function (
    this: Worksheet | Spreadsheet,
    config: jspreadsheet.SpreadsheetOptions | jspreadsheet.WorksheetOptions,
    spreadsheetLevel?: boolean
) {
    const obj = this;

    const keys = Object.keys(config);

    let spreadsheet: Spreadsheet;

    if (obj instanceof Spreadsheet) {
        spreadsheetLevel = true;

        spreadsheet = obj;
    } else {
        spreadsheet = obj.parent;
    }

    keys.forEach(function (key) {
        if (spreadsheetLevel) {
            // The options are set by name, and SpreadsheetOptions has no index signature
            (spreadsheet.config as unknown as Record<string, unknown>)[key] = (config as Record<string, unknown>)[key];

            if (key === 'toolbar') {
                if ((config as Record<string, unknown>)[key] === true) {
                    spreadsheet.showToolbar();
                } else if ((config as Record<string, unknown>)[key] === false) {
                    spreadsheet.hideToolbar();
                }
            }
        } else {
            ((obj as Worksheet).options as Record<string, unknown>)[key] = (config as Record<string, unknown>)[key];
        }
    });
};
