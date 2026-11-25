import { Worksheet } from '../Worksheet.js';
import type { Spreadsheet } from '../types/internal';

// The worksheet that a spreadsheet shows. These functions read only the spreadsheet, so the toolbar and the factory can
// import them without an import cycle.

// Spreadsheets call it too
export const getWorksheetActive = function (this: Worksheet | Spreadsheet): number {
    const spreadsheet = this instanceof Worksheet ? this.parent : this;

    // The element of a spreadsheet with tabs has the tabs object
    // getActive returns false only when no tab is open, and a tab is open while the spreadsheet has worksheets
    return spreadsheet.element.tabs ? (spreadsheet.element.tabs.getActive() as number) : 0;
};

export const getWorksheetInstance = function (this: Spreadsheet, index?: number) {
    const spreadsheet = this;

    const worksheetIndex = typeof index !== 'undefined' ? index : getWorksheetActive.call(spreadsheet);

    return spreadsheet.worksheets[worksheetIndex];
};
