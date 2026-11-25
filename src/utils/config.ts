import type { Worksheet, Spreadsheet } from '../types/internal';
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
export const setConfig = function (this: Worksheet | Spreadsheet, config: Record<string, any>, spreadsheetLevel?: boolean) {
    const obj = this;

    const keys = Object.keys(config);

    let spreadsheet: Spreadsheet;

    if (!obj.parent) {
        spreadsheetLevel = true;

        spreadsheet = obj as Spreadsheet;
    } else {
        spreadsheet = obj.parent;
    }

    keys.forEach(function (key) {
        if (spreadsheetLevel) {
            (spreadsheet.config as Record<string, any>)[key] = config[key];

            if (key === 'toolbar') {
                if (config[key] === true) {
                    spreadsheet.showToolbar();
                } else if (config[key] === false) {
                    spreadsheet.hideToolbar();
                }
            }
        } else {
            ((obj as Worksheet).options as Record<string, any>)[key] = config[key];
        }
    });
};
