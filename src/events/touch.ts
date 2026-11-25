import { openEditor } from '../utils/editor.js';
import libraryBase from '../utils/libraryBase.js';
import { getCellColumnIndex } from '../utils/values.js';
import { updateSelectionFromCoords } from '../utils/selection.js';
import { getElement, getIndex } from './helpers.js';
import type { TargetEvent } from '../types/internal';

export const touchStartControls = function (e: TargetEvent<TouchEvent>) {
    const jssTable = getElement(e.target);

    if (jssTable[0]) {
        if (libraryBase.jspreadsheet.current != jssTable[0].jssWorksheet) {
            if (libraryBase.jspreadsheet.current) {
                libraryBase.jspreadsheet.current.resetSelection();
            }
            libraryBase.jspreadsheet.current = jssTable[0].jssWorksheet;
        }
    } else {
        if (libraryBase.jspreadsheet.current && !e.target.closest('.jss_object')) {
            libraryBase.jspreadsheet.current.resetSelection();
            libraryBase.jspreadsheet.current = null;
        }
    }

    if (libraryBase.jspreadsheet.current && jssTable[0]) {
        if (!libraryBase.jspreadsheet.current.editor.isOpen()) {
            const columnId = getIndex(e.target, 'data-x');
            const rowId = getIndex(e.target, 'data-y');

            if (columnId != null && rowId != null) {
                updateSelectionFromCoords.call(libraryBase.jspreadsheet.current, columnId, rowId, undefined, undefined, e);

                if (libraryBase.jspreadsheet.current.options.editable == false) {
                    return;
                }

                libraryBase.jspreadsheet.timeControl = setTimeout(function () {
                    // Keep temporary reference to the element
                    const column =
                        libraryBase.jspreadsheet.current!.options.columns[getCellColumnIndex.call(libraryBase.jspreadsheet.current!, columnId, rowId)];
                    if (column.type == 'color') {
                        libraryBase.jspreadsheet.tmpElement = null;
                    } else {
                        libraryBase.jspreadsheet.tmpElement = e.target;
                    }
                    openEditor.call(libraryBase.jspreadsheet.current!, e.target as HTMLTableCellElement, false, e);
                }, 500);
            }
        }
    }
};

export const touchEndControls = function (e: TargetEvent<TouchEvent>) {
    // Clear any time control
    if (libraryBase.jspreadsheet.timeControl) {
        clearTimeout(libraryBase.jspreadsheet.timeControl);
        libraryBase.jspreadsheet.timeControl = null;
        // Element
        if (libraryBase.jspreadsheet.tmpElement && libraryBase.jspreadsheet.tmpElement.children[0].tagName == 'INPUT') {
            libraryBase.jspreadsheet.tmpElement.children[0].focus();
        }
        libraryBase.jspreadsheet.tmpElement = null;
    }
};
