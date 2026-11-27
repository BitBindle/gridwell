import libraryBase from '../utils/libraryBase.js';
import { getElement, getIndex } from './helpers.js';
import type { TargetEvent } from '../types/internal';

export function touchStartControls(e: TargetEvent<TouchEvent>) {
    const jssTable = getElement(e.target);

    // A touch on a worksheet makes it the active worksheet. A touch outside the spreadsheets leaves no active worksheet.
    const previous = libraryBase.jspreadsheet.current;
    if (jssTable[0]) {
        if (previous !== jssTable[0].jssWorksheet) {
            if (previous) {
                previous.resetSelection();
            }
            libraryBase.jspreadsheet.current = jssTable[0].jssWorksheet;
        }
    } else {
        if (previous && !e.target.closest('.jss_object')) {
            previous.resetSelection();
            libraryBase.jspreadsheet.current = null;
        }
    }

    const current = libraryBase.jspreadsheet.current;

    if (current && jssTable[0]) {
        if (!current.editor.isOpen()) {
            const columnId = getIndex(e.target, 'data-x');
            const rowId = getIndex(e.target, 'data-y');

            if (columnId != null && rowId != null) {
                current.selection.updateFromCoords(columnId, rowId, undefined, undefined, e);

                if (current.options.editable === false) {
                    return;
                }

                // The active worksheet when the timer runs
                libraryBase.jspreadsheet.timeControl = setTimeout(() => {
                    // Keep temporary reference to the element
                    const column = libraryBase.jspreadsheet.current!.options.columns[libraryBase.jspreadsheet.current!.cells.getColumnIndex(columnId, rowId)];
                    if (column.type === 'color') {
                        libraryBase.jspreadsheet.tmpElement = null;
                    } else {
                        libraryBase.jspreadsheet.tmpElement = e.target;
                    }
                    libraryBase.jspreadsheet.current!.editor.open(e.target as HTMLTableCellElement, false, e);
                }, 500);
            }
        }
    }
}

export function touchEndControls(e: TargetEvent<TouchEvent>) {
    // Clear any time control
    if (libraryBase.jspreadsheet.timeControl) {
        clearTimeout(libraryBase.jspreadsheet.timeControl);
        libraryBase.jspreadsheet.timeControl = null;
        // Element
        const editor = libraryBase.jspreadsheet.tmpElement?.children[0];
        if (editor?.tagName === 'INPUT' || editor?.tagName === 'TEXTAREA') {
            (editor as HTMLElement).focus();
        }
        libraryBase.jspreadsheet.tmpElement = null;
    }
}
