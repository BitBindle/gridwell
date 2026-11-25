import libraryBase from '../utils/libraryBase.js';
import { copy, paste } from '../utils/copyPaste.js';
import { getElement } from './helpers.js';
import type { TargetEvent } from '../types/internal';

export const pasteControls = function (e?: TargetEvent<ClipboardEvent>) {
    // A paste on a control with the class jss_object outside the worksheet goes to that control
    if (e && e.target && typeof e.target.closest === 'function' && !getElement(e.target)[0] && e.target.closest('.jss_object')) {
        return;
    }

    const current = libraryBase.jspreadsheet.current;

    if (current && current.selection.range) {
        if (!current.editor.isOpen()) {
            if (current.options.editable != false) {
                if (e && e.clipboardData) {
                    paste.call(current, current.selection.range[0], current.selection.range[1], e.clipboardData.getData('text'));
                    e.preventDefault();
                } else if ((window as Window & { clipboardData?: DataTransfer }).clipboardData) {
                    // Internet Explorer has the clipboard on the window
                    paste.call(
                        current,
                        current.selection.range[0],
                        current.selection.range[1],
                        (window as Window & { clipboardData?: DataTransfer }).clipboardData!.getData('text')
                    );
                }
            }
        }
    }
};

export const cutControls = function (e?: TargetEvent<ClipboardEvent>) {
    const current = libraryBase.jspreadsheet.current;

    if (current) {
        // Without a selection, there is nothing to cut
        if (!current.editor.isOpen() && current.selection.range) {
            const copied = copy.call(current, true, undefined, undefined, undefined, undefined, true);
            // oncopy can cancel the cut
            if (copied !== false && current.options.editable != false) {
                current.setValue(
                    current.selection.highlighted.map(function (record) {
                        return record.element;
                    }),
                    ''
                );
            }
        }
    }
};

export const copyControls = function (e?: TargetEvent<ClipboardEvent>) {
    const current = libraryBase.jspreadsheet.current;

    if (current) {
        if (!current.editor.isOpen()) {
            copy.call(current, true);
        }
    }
};
