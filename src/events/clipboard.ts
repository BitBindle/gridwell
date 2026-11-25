import libraryBase from '../utils/libraryBase.js';
import { copy, paste } from '../utils/copyPaste.js';
import { getElement } from './helpers.js';
import type { TargetEvent } from '../types/internal';

export const pasteControls = function (e?: TargetEvent<ClipboardEvent>) {
    // A paste on a control with the class jss_object outside the worksheet goes to that control
    if (e && e.target && typeof e.target.closest === 'function' && !getElement(e.target)[0] && e.target.closest('.jss_object')) {
        return;
    }

    if (libraryBase.jspreadsheet.current && libraryBase.jspreadsheet.current.selection.range) {
        if (!libraryBase.jspreadsheet.current.editor.isOpen()) {
            if (libraryBase.jspreadsheet.current.options.editable != false) {
                if (e && e.clipboardData) {
                    paste.call(
                        libraryBase.jspreadsheet.current,
                        libraryBase.jspreadsheet.current.selection.range[0],
                        libraryBase.jspreadsheet.current.selection.range[1],
                        e.clipboardData.getData('text')
                    );
                    e.preventDefault();
                } else if ((window as Window & { clipboardData?: DataTransfer }).clipboardData) {
                    // Internet Explorer has the clipboard on the window
                    paste.call(
                        libraryBase.jspreadsheet.current,
                        libraryBase.jspreadsheet.current.selection.range[0],
                        libraryBase.jspreadsheet.current.selection.range[1],
                        (window as Window & { clipboardData?: DataTransfer }).clipboardData!.getData('text')
                    );
                }
            }
        }
    }
};

export const cutControls = function (e?: TargetEvent<ClipboardEvent>) {
    if (libraryBase.jspreadsheet.current) {
        // Without a selection, there is nothing to cut
        if (!libraryBase.jspreadsheet.current.editor.isOpen() && libraryBase.jspreadsheet.current.selection.range) {
            const copied = copy.call(libraryBase.jspreadsheet.current, true, undefined, undefined, undefined, undefined, true);
            // oncopy can cancel the cut
            if (copied !== false && libraryBase.jspreadsheet.current.options.editable != false) {
                libraryBase.jspreadsheet.current.setValue(
                    libraryBase.jspreadsheet.current.selection.highlighted.map(function (record) {
                        return record.element;
                    }),
                    ''
                );
            }
        }
    }
};

export const copyControls = function (e?: TargetEvent<ClipboardEvent>) {
    if (libraryBase.jspreadsheet.current) {
        if (!libraryBase.jspreadsheet.current.editor.isOpen()) {
            copy.call(libraryBase.jspreadsheet.current, true);
        }
    }
};
