import jSuites from 'jsuites';
import { copyControls, cutControls, pasteControls } from '../events/clipboard.js';
import { isCtrl } from '../events/helpers.js';
import type { EditingWorksheet, SelectedWorksheet, TargetEvent, Worksheet } from '../types/internal';

/**
 * The keys on a worksheet: the keys of an open editor, and the keys of the grid, which move the selection, delete,
 * copy, paste and start the editor when a key is typed.
 *
 * Each worksheet has one, as worksheet.keyboard. The handler of keydown in src/events/keyboard.ts finds the worksheet
 * and calls it.
 */
export class Keyboard {
    readonly worksheet: Worksheet;

    constructor(worksheet: Worksheet) {
        this.worksheet = worksheet;
    }

    /**
     * The keys while an editor is open: Escape closes it without saving, Enter and Tab close it and save, and Alt+Enter adds
     * a line to a text editor that wraps.
     */
    editorKeys(e: TargetEvent<KeyboardEvent>) {
        // keyDownControls calls it only while an editor is open
        const obj = this.worksheet as EditingWorksheet;
        const column = obj.options.columns && obj.options.columns[obj.cells.getColumnIndex(obj.editor.x, obj.editor.y)];

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
    }

    /**
     * Delete and Backspace: delete the selected rows or columns after a confirmation, or clear the selected cells.
     */
    private deleteKey() {
        // gridKeys calls it, when the worksheet has a selection
        const obj = this.worksheet as SelectedWorksheet;

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
                    obj.selection.highlighted.map((record) => {
                        return record.element;
                    }),
                    ''
                );
            }
        }
    }

    /**
     * Enter moves the selection down, and adds a row below the last row. Shift+Enter moves it up.
     */
    private enterKey(e: TargetEvent<KeyboardEvent>) {
        // gridKeys calls it, when the worksheet has a selection
        const obj = this.worksheet as SelectedWorksheet;

        // Move cursor
        if (e.shiftKey) {
            obj.selection.up();
        } else {
            if (obj.options.allowInsertRow != false) {
                if (obj.options.allowManualInsertRow != false) {
                    if (obj.selection.range[1] == obj.options.data.length - 1) {
                        // New record in case the selection is in the last row
                        obj.insertRow();
                    }
                }
            }

            obj.selection.down();
        }
        e.preventDefault();
    }

    /**
     * Tab moves the selection right, and adds a column after the last column. Shift+Tab moves it left.
     */
    private tabKey(e: TargetEvent<KeyboardEvent>) {
        // gridKeys calls it, when the worksheet has a selection
        const obj = this.worksheet as SelectedWorksheet;

        // Tab
        if (e.shiftKey) {
            obj.selection.left();
        } else {
            if (obj.options.allowInsertColumn != false) {
                if (obj.options.allowManualInsertColumn != false) {
                    if (obj.selection.range[0] == obj.options.data[0].length - 1) {
                        // New record in case the selection is in the last column
                        obj.insertColumn();
                    }
                }
            }

            obj.selection.right();
        }
        e.preventDefault();
    }

    /**
     * Ctrl or Cmd with a letter: select all, save, redo, undo, copy, cut and paste.
     */
    private shortcutKeys(e: TargetEvent<KeyboardEvent>) {
        // gridKeys calls it, when the worksheet has a selection
        const obj = this.worksheet as SelectedWorksheet;

        if (e.which == 65) {
            // Ctrl + A
            obj.selection.selectAll();
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
            obj.clipboard.copy(true);
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
    }

    /**
     * A key that starts an edition: Space opens the editor or toggles a checkbox or a radio, F2 opens the editor with the
     * value of the cell, and a character opens the editor with that character.
     */
    private typingKeys(e: TargetEvent<KeyboardEvent>) {
        // gridKeys calls it, when the worksheet has a selection
        const obj = this.worksheet as SelectedWorksheet;

        if (obj.options.editable != false) {
            const rowId = obj.selection.range[1];
            const columnId = obj.selection.range[0];
            const column = obj.options.columns[obj.cells.getColumnIndex(columnId, rowId)];

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
    }

    /**
     * The keys on the grid while no editor is open: the arrows, Home and End move the selection, and the other keys
     * delete, insert, run a shortcut or start an edition.
     */
    gridKeys(e: TargetEvent<KeyboardEvent>) {
        // keyDownControls calls it only when the worksheet has a selection
        const obj = this.worksheet as SelectedWorksheet;

        // Which key
        if (e.which == 37) {
            obj.selection.left(e.shiftKey, e.ctrlKey);
            e.preventDefault();
        } else if (e.which == 39) {
            obj.selection.right(e.shiftKey, e.ctrlKey);
            e.preventDefault();
        } else if (e.which == 38) {
            obj.selection.up(e.shiftKey, e.ctrlKey);
            e.preventDefault();
        } else if (e.which == 40) {
            obj.selection.down(e.shiftKey, e.ctrlKey);
            e.preventDefault();
        } else if (e.which == 36) {
            obj.selection.first(e.shiftKey, e.ctrlKey);
            e.preventDefault();
        } else if (e.which == 35) {
            obj.selection.last(e.shiftKey, e.ctrlKey);
            e.preventDefault();
        } else if (e.which == 46 || e.which == 8) {
            // Delete
            this.deleteKey();
        } else if (e.which == 13) {
            this.enterKey(e);
        } else if (e.which == 9) {
            this.tabKey(e);
        } else {
            if ((e.ctrlKey || e.metaKey) && !e.shiftKey) {
                this.shortcutKeys(e);
            } else {
                this.typingKeys(e);
            }
        }
    }
}
