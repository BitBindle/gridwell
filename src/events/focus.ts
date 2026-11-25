import { closeEditor } from '../utils/editor.js';
import libraryBase from '../utils/libraryBase.js';
import { getIndex } from './helpers.js';
import type { TargetEvent, Worksheet } from '../types/internal';

/**
 * When the focus moves into another worksheet, for example with Tab, that worksheet becomes the active worksheet, so the
 * keys go to the grid that has the focus. The worksheet that was active closes its editor and loses its selection, as
 * after a press on the other worksheet. A focused cell of a worksheet without a selection becomes the selection. The
 * focus on the tabs or the toolbar of a spreadsheet, which are outside its worksheets, changes nothing.
 */
export const focusInControls = function (e: TargetEvent<FocusEvent>) {
    // The target can be a node without classList, such as the document
    if (!e.target || e.target.nodeType !== 1) {
        return;
    }

    const element = e.target.closest('.jss_container') as Worksheet['element'] | null;
    if (!element || libraryBase.jspreadsheet.current === element.jssWorksheet) {
        return;
    }

    const previous = libraryBase.jspreadsheet.current;
    if (previous) {
        if (previous.editor.isOpen()) {
            closeEditor.call(previous, previous.editor.cell, true);
        }
        previous.resetSelection();
    }

    const worksheet = element.jssWorksheet;
    libraryBase.jspreadsheet.current = worksheet;

    if (!worksheet.selection.range && e.target.tagName === 'TD' && worksheet.tbody.contains(e.target)) {
        const x = getIndex(e.target, 'data-x');
        const y = getIndex(e.target, 'data-y');
        if (x !== null && y !== null) {
            worksheet.updateSelectionFromCoords(x, y);
        }
    }
};
