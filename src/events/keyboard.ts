import libraryBase from '../utils/libraryBase.js';
import { getElement } from './helpers.js';
import type { TargetEvent } from '../types/internal';

export function keyDownControls(e: TargetEvent<KeyboardEvent>) {
    // Keys on the worksheet tabs, on the button to add a worksheet, on the toolbar and on the context menu do not go to the
    // grid
    const keyPath = typeof e.composedPath === 'function' ? e.composedPath() : [];
    // The first node of the path can be a node without closest, such as the document. The code checks for closest.
    const keyTarget = (keyPath.length ? keyPath[0] : e.target) as HTMLElement;
    if (
        keyTarget &&
        typeof keyTarget.closest === 'function' &&
        keyTarget.closest('.jss_spreadsheet > .jtabs-headers-container, .jss_spreadsheet > .jss_toolbar, .jss_spreadsheet > .jss_contextmenu')
    ) {
        return;
    }

    const current = libraryBase.gridwell.current;

    if (current) {
        // Keys on an element outside the spreadsheet, for example an input or a button next to the grid that has the focus
        // after Tab, also in an open shadow root, do not go to the grid. Keys with the focus on the body, where it falls when
        // the focused element is removed, still go to the grid, and so do keys on an element that the key removed, such as
        // a dropdown editor that Tab closed. An editor can have controls outside the spreadsheet, such as a jss_object
        // control, so this applies only when no editor is open.
        if (!current.editor.isOpen()) {
            if (
                keyTarget &&
                typeof keyTarget.closest === 'function' &&
                keyTarget.isConnected &&
                keyTarget !== document.body &&
                keyTarget !== document.documentElement &&
                !getElement(keyTarget)[0]
            ) {
                return;
            }
        }

        if (current.editor.isOpen()) {
            current.keyboard.editorKeys(e);
        }

        if (!current.editor.isOpen() && current.selection.range) {
            current.keyboard.gridKeys(e);
        } else {
            if (e.target.classList.contains('jss_search')) {
                if (libraryBase.gridwell.timeControl) {
                    clearTimeout(libraryBase.gridwell.timeControl);
                }

                // The active worksheet when the timer runs
                libraryBase.gridwell.timeControl = setTimeout(() => {
                    libraryBase.gridwell.current!.search((e.target as HTMLInputElement).value);
                }, 200);
            }
        }
    }
}
