import { translate } from '../utils/dictionary.js';
import { copyControls, cutControls, pasteControls } from '../events/clipboard.js';
import { isCtrl } from '../events/helpers.js';
import type { EditingWorksheet, SelectedWorksheet, TargetEvent, Worksheet } from '../types/internal';
import { askConfirm } from '../utils/dialogs.js';

/**
 * Type text at the caret of a focused field with the insertText command, as the user does. False when the browser does
 * not do the command.
 */
function insertText(field: HTMLTextAreaElement, text: string) {
    if (document.activeElement !== field || typeof document.execCommand !== 'function') {
        return false;
    }
    const value = field.value;
    const start = field.selectionStart;
    try {
        // jsdom gives true and changes nothing
        return document.execCommand('insertText', false, text) && (field.value !== value || field.selectionStart !== start);
    } catch {
        return false;
    }
}

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
     * a line break to a textarea editor.
     */
    editorKeys(e: TargetEvent<KeyboardEvent>) {
        // keyDownControls calls it only while an editor is open
        const worksheet = this.worksheet as EditingWorksheet;
        const column = worksheet.options.columns?.[worksheet.cells.getColumnIndex(worksheet.editor.x, worksheet.editor.y)];

        if (e.which === 27) {
            // Escape
            if (worksheet.editor.isOpen()) {
                // Exit without saving
                worksheet.editor.close(worksheet.editor.cell, false);
            }
            e.preventDefault();
        } else if (e.which === 13) {
            // Enter
            if (column?.type === 'calendar') {
                worksheet.editor.close(worksheet.editor.cell, true);
            } else if (column?.type === 'dropdown') {
                // Do nothing
            } else {
                // Alt+Enter adds a line break at the caret of a textarea, and the editor stays open. The insertText command
                // goes into the undo of the browser and shows the caret. Without it, setRangeText adds the line break.
                const editor = worksheet.editor.cell.children[0] as HTMLTextAreaElement;
                if (e.altKey && editor.tagName === 'TEXTAREA') {
                    if (!insertText(editor, '\n')) {
                        editor.setRangeText('\n', editor.selectionStart, editor.selectionEnd, 'end');
                        editor.dispatchEvent(new Event('input', { bubbles: true }));
                    }
                    // Each line of a one-line editor has the height of the editor. Chromium does not scroll to the new line.
                    if (editor.getAttribute('wrap') === 'off') {
                        const line = editor.value.slice(0, editor.selectionStart).split('\n').length - 1;
                        editor.scrollTop = line * editor.clientHeight;
                    }
                    e.preventDefault();
                } else {
                    editor.blur();
                }
            }
        } else if (e.which === 9) {
            // Tab
            if (column && ['calendar', 'html'].includes(column.type as string)) {
                worksheet.editor.close(worksheet.editor.cell, true);
            } else {
                (worksheet.editor.cell.children[0] as HTMLElement).blur();
            }
        }
    }

    /**
     * Delete and Backspace: delete the selected rows or columns after a confirmation, or clear the selected cells.
     */
    private deleteKey() {
        // gridKeys calls it, when the worksheet has a selection
        const worksheet = this.worksheet as SelectedWorksheet;

        if (worksheet.options.editable !== false) {
            if (worksheet.selection.row != null) {
                if (worksheet.options.allowDeleteRow !== false) {
                    if (askConfirm(worksheet.parent.config, translate('Are you sure to delete the selected rows?'))) {
                        worksheet.deleteRow();
                    }
                }
            } else if (worksheet.selection.column != null) {
                if (worksheet.options.allowDeleteColumn !== false) {
                    if (askConfirm(worksheet.parent.config, translate('Are you sure to delete the selected columns?'))) {
                        worksheet.deleteColumn();
                    }
                }
            } else {
                // Change value
                worksheet.setValue(
                    worksheet.selection.highlighted.map((record) => {
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
        const worksheet = this.worksheet as SelectedWorksheet;

        // Move cursor
        if (e.shiftKey) {
            worksheet.selection.up();
        } else {
            if (worksheet.options.allowInsertRow !== false) {
                if (worksheet.options.allowManualInsertRow !== false) {
                    if (worksheet.selection.range[1] === worksheet.options.data.length - 1) {
                        // New record in case the selection is in the last row
                        worksheet.insertRow();
                    }
                }
            }

            worksheet.selection.down();
        }
        e.preventDefault();
    }

    /**
     * Tab moves the selection right, and Shift+Tab moves it left. At the end of the row, where the selection does not move,
     * the key keeps its default action and moves the focus out of the grid, as in the WAI-ARIA grid pattern.
     */
    private tabKey(e: TargetEvent<KeyboardEvent>) {
        // gridKeys calls it, when the worksheet has a selection
        const worksheet = this.worksheet as SelectedWorksheet;

        const before = worksheet.selection.range.join();

        if (e.shiftKey) {
            worksheet.selection.left();
        } else {
            worksheet.selection.right();
        }

        if (worksheet.selection.range?.join() !== before) {
            e.preventDefault();
        }
    }

    /**
     * Ctrl or Cmd with a letter: select all, save, redo, undo, copy, cut and paste.
     */
    private shortcutKeys(e: TargetEvent<KeyboardEvent>) {
        // gridKeys calls it, when the worksheet has a selection
        const worksheet = this.worksheet as SelectedWorksheet;

        if (e.which === 65) {
            // Ctrl + A
            worksheet.selection.selectAll();
            e.preventDefault();
        } else if (e.which === 83) {
            // Ctrl + S
            worksheet.download();
            e.preventDefault();
        } else if (e.which === 89) {
            // Ctrl + Y
            worksheet.redo();
            e.preventDefault();
        } else if (e.which === 90) {
            // Ctrl + Z
            worksheet.undo();
            e.preventDefault();
        } else if (e.which === 67) {
            // Ctrl + C
            worksheet.clipboard.copy(true);
            e.preventDefault();
        } else if (e.which === 88) {
            // Ctrl + X
            if (worksheet.options.editable !== false) {
                cutControls();
            } else {
                copyControls();
            }
            e.preventDefault();
        } else if (e.which === 86) {
            // Ctrl + V
            pasteControls();
        }
    }

    /**
     * Open the editor of the cell (x, y). In the row source mode, the window scrolls to the row of the cell when it has no
     * record.
     */
    private openEditor(y: number, x: number, empty: boolean, e: KeyboardEvent) {
        const worksheet = this.worksheet;
        const element = worksheet.rowSource ? worksheet.rowSource.reveal(x, y) : worksheet.records[y][x].element;
        if (element) {
            worksheet.editor.open(element as HTMLTableCellElement, empty, e);
        }
    }

    /**
     * A key that starts an edition: Space opens the editor or toggles a checkbox or a radio, F2 opens the editor with the
     * value of the cell, and a character opens the editor with that character.
     */
    private typingKeys(e: TargetEvent<KeyboardEvent>) {
        // gridKeys calls it, when the worksheet has a selection
        const worksheet = this.worksheet as SelectedWorksheet;

        if (worksheet.options.editable !== false) {
            const rowId = worksheet.selection.range[1];
            const columnId = worksheet.selection.range[0];
            const column = worksheet.options.columns[worksheet.cells.getColumnIndex(columnId, rowId)];

            // Characters able to start a edition
            if (e.keyCode === 32) {
                // Space
                e.preventDefault();
                if (column.type === 'checkbox' || column.type === 'radio') {
                    worksheet.editor.setCheckRadioValue();
                } else {
                    // Start edition
                    this.openEditor(rowId, columnId, true, e);
                }
            } else if (e.keyCode === 113) {
                // Start edition with current content F2
                this.openEditor(rowId, columnId, false, e);
            } else if ((e.key.length === 1 || e.key === 'Process') && !(e.altKey || isCtrl(e))) {
                // Start edition
                this.openEditor(rowId, columnId, true, e);
                // Prevent entries in the calendar
                if (column?.type === 'calendar') {
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
        const worksheet = this.worksheet as SelectedWorksheet;

        // Which key
        if (e.which === 37) {
            worksheet.selection.left(e.shiftKey, e.ctrlKey);
            e.preventDefault();
        } else if (e.which === 39) {
            worksheet.selection.right(e.shiftKey, e.ctrlKey);
            e.preventDefault();
        } else if (e.which === 38) {
            worksheet.selection.up(e.shiftKey, e.ctrlKey);
            e.preventDefault();
        } else if (e.which === 40) {
            worksheet.selection.down(e.shiftKey, e.ctrlKey);
            e.preventDefault();
        } else if (e.which === 36) {
            worksheet.selection.first(e.shiftKey, e.ctrlKey);
            e.preventDefault();
        } else if (e.which === 35) {
            worksheet.selection.last(e.shiftKey, e.ctrlKey);
            e.preventDefault();
        } else if (e.which === 46 || e.which === 8) {
            // Delete
            this.deleteKey();
        } else if (e.which === 13) {
            this.enterKey(e);
        } else if (e.which === 9) {
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
