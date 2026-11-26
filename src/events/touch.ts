import libraryBase from '../utils/libraryBase.js';
import { getElement, getIndex } from './helpers.js';
import type { TargetEvent } from '../types/internal';

export function touchStartControls(e: TargetEvent<TouchEvent>) {
    const jssTable = getElement(e.target);

    // A touch on a worksheet makes it the active worksheet. A touch outside the spreadsheets leaves no active worksheet.
    const previous = libraryBase.gridwell.current;
    if (jssTable[0]) {
        if (previous !== jssTable[0].jssWorksheet) {
            if (previous) {
                previous.resetSelection();
            }
            libraryBase.gridwell.current = jssTable[0].jssWorksheet;
        }
    } else {
        if (previous && !e.target.closest('.jss_object')) {
            previous.resetSelection();
            libraryBase.gridwell.current = null;
        }
    }

    const current = libraryBase.gridwell.current;

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
                libraryBase.gridwell.timeControl = setTimeout(() => {
                    // Keep temporary reference to the element
                    const column = libraryBase.gridwell.current!.options.columns[libraryBase.gridwell.current!.cells.getColumnIndex(columnId, rowId)];
                    if (column.type === 'color') {
                        libraryBase.gridwell.tmpElement = null;
                    } else {
                        libraryBase.gridwell.tmpElement = e.target;
                    }
                    libraryBase.gridwell.current!.editor.open(e.target as HTMLTableCellElement, false, e);
                }, 500);
            }
        }
    }
}

export function touchEndControls(e: TargetEvent<TouchEvent>) {
    // Clear any time control
    if (libraryBase.gridwell.timeControl) {
        clearTimeout(libraryBase.gridwell.timeControl);
        libraryBase.gridwell.timeControl = null;
        // Element
        if (libraryBase.gridwell.tmpElement?.children[0].tagName === 'INPUT') {
            // The test of the tag name above
            (libraryBase.gridwell.tmpElement.children[0] as HTMLInputElement).focus();
        }
        libraryBase.gridwell.tmpElement = null;
    }
}
