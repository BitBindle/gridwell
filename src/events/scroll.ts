import libraryBase from '../utils/libraryBase.js';

/**
 * A scroll of the page, or of an element around a worksheet, moves the cells under the sticky header rows and the
 * frozen rows. Each worksheet in the scrolled element checks again if they cover its fill handle. View.scroll checks it
 * for a scroll of the content of the worksheet.
 */
export function scrollControls(e: Event) {
    const target = e.target as Node | null;
    if (!target) {
        return;
    }
    for (const spreadsheet of libraryBase.jspreadsheet.spreadsheet) {
        for (const worksheet of spreadsheet.worksheets) {
            if (target !== worksheet.content && (target === document || target.contains(worksheet.element))) {
                worksheet.selection.updateCornerCover();
            }
        }
    }
}
