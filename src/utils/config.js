// @ts-check
/** @import { Worksheet, Spreadsheet } from '../types/internal' */
/**
 * Get table config information
 * @this {Worksheet}
 */
export const getWorksheetConfig = function () {
    const obj = this;

    return obj.options;
};

/** @this {Spreadsheet} */
export const getSpreadsheetConfig = function () {
    const spreadsheet = this;

    return spreadsheet.config;
};

/**
 * Change options of the worksheet, or of the spreadsheet when this is a spreadsheet or spreadsheetLevel is true.
 * @this {Worksheet | Spreadsheet}
 * @param {Record<string, any>} config - The options to change.
 * @param {boolean} [spreadsheetLevel] - Change the options of the spreadsheet of the worksheet.
 */
export const setConfig = function (config, spreadsheetLevel) {
    const obj = this;

    const keys = Object.keys(config);

    /** @type {Spreadsheet} */
    let spreadsheet;

    if (!obj.parent) {
        spreadsheetLevel = true;

        spreadsheet = /** @type {Spreadsheet} */ (obj);
    } else {
        spreadsheet = obj.parent;
    }

    keys.forEach(function (key) {
        if (spreadsheetLevel) {
            /** @type {Record<string, any>} */ (spreadsheet.config)[key] = config[key];

            if (key === 'toolbar') {
                if (config[key] === true) {
                    spreadsheet.showToolbar();
                } else if (config[key] === false) {
                    spreadsheet.hideToolbar();
                }
            }
        } else {
            /** @type {Record<string, any>} */ (/** @type {Worksheet} */ (obj).options)[key] = config[key];
        }
    });
};
