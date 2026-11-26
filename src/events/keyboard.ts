import jSuites from 'jsuites';
import libraryBase from '../utils/libraryBase.js';
import { down, first, last, left, right, up } from '../utils/keys.js';
import { getCellColumnIndex } from '../utils/values.js';
import { selectAll } from '../utils/selection.js';
import { copy } from '../utils/copyPaste.js';
import { copyControls, cutControls, pasteControls } from './clipboard.js';
import { getElement, isCtrl } from './helpers.js';
import type { EditingWorksheet, SelectedWorksheet, TargetEvent } from '../types/internal';

/**
 * The keys while an editor is open: Escape closes it without saving, Enter and Tab close it and save, and Alt+Enter adds
 * a line to a text editor that wraps.
 */
const editorKeys = function (this: EditingWorksheet, e: TargetEvent<KeyboardEvent>) {
    const obj = this;
    const column = obj.options.columns && obj.options.columns[getCellColumnIndex.call(obj, obj.editor.x, obj.editor.y)];

    if (e.which == 27) {
        // Escape
        if (obj.editor.isOpen()) {
            // Exit without saving
            obj.editor.close(obj.editor.cell, false);
        }
        e.preventDefault();
    } else if (e.which == 13) {
        // Enter
        if (column && column.type == 'calendar') {
            obj.editor.close(obj.editor.cell, true);
        } else if (column && column.type == 'dropdown') {
            // Do nothing
        } else {
            // Alt enter -> do not close editor
            if (
                (obj.options.wordWrap == true ||
                    (column && column.wordWrap == true) ||
                    (obj.options.data[obj.editor.y][obj.editor.x] && (obj.options.data[obj.editor.y][obj.editor.x] as string).length > 200)) &&
                e.altKey
            ) {
                // Add new line to the editor
                const editorTextarea = obj.editor.cell.children[0] as HTMLTextAreaElement;
                let editorValue = (obj.editor.cell.children[0] as HTMLTextAreaElement).value;
                const editorIndexOf = editorTextarea.selectionStart;
                editorValue = editorValue.slice(0, editorIndexOf) + '\n' + editorValue.slice(editorIndexOf);
                editorTextarea.value = editorValue;
                editorTextarea.focus();
                editorTextarea.selectionStart = editorIndexOf + 1;
                editorTextarea.selectionEnd = editorIndexOf + 1;
            } else {
                (obj.editor.cell.children[0] as HTMLElement).blur();
            }
        }
    } else if (e.which == 9) {
        // Tab
        if (column && ['calendar', 'html'].includes(column.type as string)) {
            obj.editor.close(obj.editor.cell, true);
        } else {
            (obj.editor.cell.children[0] as HTMLElement).blur();
        }
    }
};

/**
 * Delete and Backspace: delete the selected rows or columns after a confirmation, or clear the selected cells.
 */
const deleteKey = function (this: SelectedWorksheet) {
    const obj = this;

    if (obj.options.editable != false) {
        if (obj.selection.row != null) {
            if (obj.options.allowDeleteRow != false) {
                if (confirm(jSuites.translate('Are you sure to delete the selected rows?'))) {
                    obj.deleteRow();
                }
            }
        } else if (obj.selection.column != null) {
            if (obj.options.allowDeleteColumn != false) {
                if (confirm(jSuites.translate('Are you sure to delete the selected columns?'))) {
                    obj.deleteColumn();
                }
            }
        } else {
            // Change value
            obj.setValue(
                obj.selection.highlighted.map(function (record) {
                    return record.element;
                }),
                ''
            );
        }
    }
};

/**
 * Enter moves the selection down, and adds a row below the last row. Shift+Enter moves it up.
 */
const enterKey = function (this: SelectedWorksheet, e: TargetEvent<KeyboardEvent>) {
    const obj = this;

    // Move cursor
    if (e.shiftKey) {
        up.call(obj);
    } else {
        if (obj.options.allowInsertRow != false) {
            if (obj.options.allowManualInsertRow != false) {
                if (obj.selection.range[1] == obj.options.data.length - 1) {
                    // New record in case the selection is in the last row
                    obj.insertRow();
                }
            }
        }

        down.call(obj);
    }
    e.preventDefault();
};

/**
 * Tab moves the selection right, and adds a column after the last column. Shift+Tab moves it left.
 */
const tabKey = function (this: SelectedWorksheet, e: TargetEvent<KeyboardEvent>) {
    const obj = this;

    // Tab
    if (e.shiftKey) {
        left.call(obj);
    } else {
        if (obj.options.allowInsertColumn != false) {
            if (obj.options.allowManualInsertColumn != false) {
                if (obj.selection.range[0] == obj.options.data[0].length - 1) {
                    // New record in case the selection is in the last column
                    obj.insertColumn();
                }
            }
        }

        right.call(obj);
    }
    e.preventDefault();
};

/**
 * Ctrl or Cmd with a letter: select all, save, redo, undo, copy, cut and paste.
 */
const shortcutKeys = function (this: SelectedWorksheet, e: TargetEvent<KeyboardEvent>) {
    const obj = this;

    if (e.which == 65) {
        // Ctrl + A
        selectAll.call(obj);
        e.preventDefault();
    } else if (e.which == 83) {
        // Ctrl + S
        obj.download();
        e.preventDefault();
    } else if (e.which == 89) {
        // Ctrl + Y
        obj.redo();
        e.preventDefault();
    } else if (e.which == 90) {
        // Ctrl + Z
        obj.undo();
        e.preventDefault();
    } else if (e.which == 67) {
        // Ctrl + C
        copy.call(obj, true);
        e.preventDefault();
    } else if (e.which == 88) {
        // Ctrl + X
        if (obj.options.editable != false) {
            cutControls();
        } else {
            copyControls();
        }
        e.preventDefault();
    } else if (e.which == 86) {
        // Ctrl + V
        pasteControls();
    }
};

/**
 * A key that starts an edition: Space opens the editor or toggles a checkbox or a radio, F2 opens the editor with the
 * value of the cell, and a character opens the editor with that character.
 */
const typingKeys = function (this: SelectedWorksheet, e: TargetEvent<KeyboardEvent>) {
    const obj = this;

    if (obj.options.editable != false) {
        const rowId = obj.selection.range[1];
        const columnId = obj.selection.range[0];
        const column = obj.options.columns[getCellColumnIndex.call(obj, columnId, rowId)];

        // Characters able to start a edition
        if (e.keyCode == 32) {
            // Space
            e.preventDefault();
            if (column.type == 'checkbox' || column.type == 'radio') {
                obj.editor.setCheckRadioValue();
            } else {
                // Start edition
                obj.editor.open(obj.records[rowId][columnId].element, true, e);
            }
        } else if (e.keyCode == 113) {
            // Start edition with current content F2
            obj.editor.open(obj.records[rowId][columnId].element, false, e);
        } else if ((e.key.length === 1 || e.key === 'Process') && !(e.altKey || isCtrl(e))) {
            // Start edition
            obj.editor.open(obj.records[rowId][columnId].element, true, e);
            // Prevent entries in the calendar
            if (column && column.type == 'calendar') {
                e.preventDefault();
            }
        }
    }
};

/**
 * The keys on the grid while no editor is open: the arrows, Home and End move the selection, and the other keys
 * delete, insert, run a shortcut or start an edition.
 */
const gridKeys = function (this: SelectedWorksheet, e: TargetEvent<KeyboardEvent>) {
    const obj = this;

    // Which key
    if (e.which == 37) {
        left.call(obj, e.shiftKey, e.ctrlKey);
        e.preventDefault();
    } else if (e.which == 39) {
        right.call(obj, e.shiftKey, e.ctrlKey);
        e.preventDefault();
    } else if (e.which == 38) {
        up.call(obj, e.shiftKey, e.ctrlKey);
        e.preventDefault();
    } else if (e.which == 40) {
        down.call(obj, e.shiftKey, e.ctrlKey);
        e.preventDefault();
    } else if (e.which == 36) {
        first.call(obj, e.shiftKey, e.ctrlKey);
        e.preventDefault();
    } else if (e.which == 35) {
        last.call(obj, e.shiftKey, e.ctrlKey);
        e.preventDefault();
    } else if (e.which == 46 || e.which == 8) {
        // Delete
        deleteKey.call(obj);
    } else if (e.which == 13) {
        enterKey.call(obj, e);
    } else if (e.which == 9) {
        tabKey.call(obj, e);
    } else {
        if ((e.ctrlKey || e.metaKey) && !e.shiftKey) {
            shortcutKeys.call(obj, e);
        } else {
            typingKeys.call(obj, e);
        }
    }
};

export const keyDownControls = function (e: TargetEvent<KeyboardEvent>) {
    // Keys on the worksheet tabs and on the button to add a worksheet do not go to the grid
    const keyPath = typeof e.composedPath === 'function' ? e.composedPath() : [];
    // The first node of the path can be a node without closest, such as the document. The code checks for closest.
    const keyTarget = (keyPath.length ? keyPath[0] : e.target) as HTMLElement;
    if (keyTarget && typeof keyTarget.closest === 'function' && keyTarget.closest('.jss_spreadsheet > .jtabs-headers-container')) {
        return;
    }

    const current = libraryBase.jspreadsheet.current;

    if (current) {
        // Keys on an element outside the spreadsheet, for example an input or a button next to the grid that has the focus
        // after Tab, also in an open shadow root, do not go to the grid. Keys with the focus on the body, where it falls when
        // the focused element is removed, still go to the grid. An editor can have controls outside the spreadsheet, such as
        // a jss_object control, so this applies only when no editor is open.
        if (!current.editor.isOpen()) {
            if (
                keyTarget &&
                typeof keyTarget.closest === 'function' &&
                keyTarget !== document.body &&
                keyTarget !== document.documentElement &&
                !getElement(keyTarget)[0]
            ) {
                return;
            }
        }

        if (current.editor.isOpen()) {
            // isOpen narrows the editor, not the worksheet
            editorKeys.call(current as EditingWorksheet, e);
        }

        if (!current.editor.isOpen() && current.selection.range) {
            // The test of the range narrows the selection, not the worksheet
            gridKeys.call(current as SelectedWorksheet, e);
        } else {
            if (e.target.classList.contains('jss_search')) {
                if (libraryBase.jspreadsheet.timeControl) {
                    clearTimeout(libraryBase.jspreadsheet.timeControl);
                }

                // The active worksheet when the timer runs
                libraryBase.jspreadsheet.timeControl = setTimeout(function () {
                    libraryBase.jspreadsheet.current!.search((e.target as HTMLInputElement).value);
                }, 200);
            }
        }
    }
};
