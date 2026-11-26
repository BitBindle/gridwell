import libraryBase from '../utils/libraryBase.js';
import { getElement } from './helpers.js';
import type { TargetEvent } from '../types/internal';

export function pasteControls(e?: TargetEvent<ClipboardEvent>) {
    // A paste on a control with the class jss_object outside the worksheet goes to that control
    if (e?.target && typeof e.target.closest === 'function' && !getElement(e.target)[0] && e.target.closest('.jss_object')) {
        return;
    }

    const current = libraryBase.jspreadsheet.current;

    if (current?.selection.range) {
        if (!current.editor.isOpen()) {
            if (current.options.editable !== false) {
                if (e?.clipboardData) {
                    current.clipboard.paste(current.selection.range[0], current.selection.range[1], e.clipboardData.getData('text'));
                    e.preventDefault();
                } else if ((window as Window & { clipboardData?: DataTransfer }).clipboardData) {
                    // Internet Explorer has the clipboard on the window
                    current.clipboard.paste(
                        current.selection.range[0],
                        current.selection.range[1],
                        (window as Window & { clipboardData?: DataTransfer }).clipboardData!.getData('text')
                    );
                }
            }
        }
    }
}

export function cutControls(e?: TargetEvent<ClipboardEvent>) {
    const current = libraryBase.jspreadsheet.current;

    if (current) {
        // Without a selection, there is nothing to cut
        if (!current.editor.isOpen() && current.selection.range) {
            const copied = current.clipboard.copy(true, undefined, undefined, undefined, undefined, true);
            // oncopy can cancel the cut
            if (copied !== false && current.options.editable !== false) {
                current.setValue(
                    current.selection.highlighted.map((record) => {
                        return record.element;
                    }),
                    ''
                );
            }
        }
    }
}

export function copyControls(e?: TargetEvent<ClipboardEvent>) {
    const current = libraryBase.jspreadsheet.current;

    if (current) {
        if (!current.editor.isOpen()) {
            current.clipboard.copy(true);
        }
    }
}
