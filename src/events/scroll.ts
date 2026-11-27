import libraryBase from '../utils/libraryBase.js';

/**
 * A scroll of the page, or of an element around the worksheet, moves the cells under the sticky header rows and the
 * frozen rows. The active worksheet checks again if they cover its fill handle. View.scroll checks it for a scroll of
 * the content of the worksheet.
 */
export function scrollControls(e: Event) {
    const current = libraryBase.jspreadsheet.current;
    if (current && e.target !== current.content) {
        current.selection.updateCornerCover();
    }
}
