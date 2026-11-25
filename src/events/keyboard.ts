import jSuites from 'jsuites';
import { closeEditor, openEditor, setCheckRadioValue } from '../utils/editor.js';
import libraryBase from '../utils/libraryBase.js';
import { down, first, last, left, right, up } from '../utils/keys.js';
import { getCellColumnIndex } from '../utils/internal.js';
import { selectAll } from '../utils/selection.js';
import { copy } from '../utils/copyPaste.js';
import { copyControls, cutControls, pasteControls } from './clipboard.js';
import { getElement, isCtrl } from './helpers.js';
import type { TargetEvent } from '../types/internal';

export const keyDownControls = function (e: TargetEvent<KeyboardEvent>) {
    // Keys on the worksheet tabs and on the button to add a worksheet do not go to the grid
    const keyPath = typeof e.composedPath === 'function' ? e.composedPath() : [];
    // The first node of the path can be a node without closest, such as the document. The code checks for closest.
    const keyTarget = (keyPath.length ? keyPath[0] : e.target) as HTMLElement;
    if (keyTarget && typeof keyTarget.closest === 'function' && keyTarget.closest('.jss_spreadsheet > .jtabs-headers-container')) {
        return;
    }

    if (libraryBase.jspreadsheet.current) {
        // Keys on an element outside the spreadsheet, for example an input or a button next to the grid that has the focus
        // after Tab, also in an open shadow root, do not go to the grid. Keys with the focus on the body, where it falls when
        // the focused element is removed, still go to the grid. An editor can have controls outside the spreadsheet, such as
        // a jss_object control, so this applies only when no editor is open.
        if (!libraryBase.jspreadsheet.current.editor.isOpen()) {
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

        if (libraryBase.jspreadsheet.current.editor.isOpen()) {
            const current = libraryBase.jspreadsheet.current;
            const column = current.options.columns && current.options.columns[getCellColumnIndex.call(current, current.editor.x, current.editor.y)];

            if (e.which == 27) {
                // Escape
                if (libraryBase.jspreadsheet.current.editor.isOpen()) {
                    // Exit without saving
                    closeEditor.call(libraryBase.jspreadsheet.current, libraryBase.jspreadsheet.current.editor.cell, false);
                }
                e.preventDefault();
            } else if (e.which == 13) {
                // Enter
                if (column && column.type == 'calendar') {
                    closeEditor.call(libraryBase.jspreadsheet.current, libraryBase.jspreadsheet.current.editor.cell, true);
                } else if (column && column.type == 'dropdown') {
                    // Do nothing
                } else {
                    // Alt enter -> do not close editor
                    if (
                        (libraryBase.jspreadsheet.current.options.wordWrap == true ||
                            (column && column.wordWrap == true) ||
                            (libraryBase.jspreadsheet.current.options.data[libraryBase.jspreadsheet.current.editor.y][
                                libraryBase.jspreadsheet.current.editor.x
                            ] &&
                                (
                                    libraryBase.jspreadsheet.current.options.data[libraryBase.jspreadsheet.current.editor.y][
                                        libraryBase.jspreadsheet.current.editor.x
                                    ] as string
                                ).length > 200)) &&
                        e.altKey
                    ) {
                        // Add new line to the editor
                        const editorTextarea = libraryBase.jspreadsheet.current.editor.cell.children[0] as HTMLTextAreaElement;
                        let editorValue = (libraryBase.jspreadsheet.current.editor.cell.children[0] as HTMLTextAreaElement).value;
                        const editorIndexOf = editorTextarea.selectionStart;
                        editorValue = editorValue.slice(0, editorIndexOf) + '\n' + editorValue.slice(editorIndexOf);
                        editorTextarea.value = editorValue;
                        editorTextarea.focus();
                        editorTextarea.selectionStart = editorIndexOf + 1;
                        editorTextarea.selectionEnd = editorIndexOf + 1;
                    } else {
                        (libraryBase.jspreadsheet.current.editor.cell.children[0] as HTMLElement).blur();
                    }
                }
            } else if (e.which == 9) {
                // Tab
                if (column && ['calendar', 'html'].includes(column.type as string)) {
                    closeEditor.call(libraryBase.jspreadsheet.current, libraryBase.jspreadsheet.current.editor.cell, true);
                } else {
                    (libraryBase.jspreadsheet.current.editor.cell.children[0] as HTMLElement).blur();
                }
            }
        }

        if (!libraryBase.jspreadsheet.current.editor.isOpen() && libraryBase.jspreadsheet.current.selection.range) {
            // Which key
            if (e.which == 37) {
                left.call(libraryBase.jspreadsheet.current, e.shiftKey, e.ctrlKey);
                e.preventDefault();
            } else if (e.which == 39) {
                right.call(libraryBase.jspreadsheet.current, e.shiftKey, e.ctrlKey);
                e.preventDefault();
            } else if (e.which == 38) {
                up.call(libraryBase.jspreadsheet.current, e.shiftKey, e.ctrlKey);
                e.preventDefault();
            } else if (e.which == 40) {
                down.call(libraryBase.jspreadsheet.current, e.shiftKey, e.ctrlKey);
                e.preventDefault();
            } else if (e.which == 36) {
                first.call(libraryBase.jspreadsheet.current, e.shiftKey, e.ctrlKey);
                e.preventDefault();
            } else if (e.which == 35) {
                last.call(libraryBase.jspreadsheet.current, e.shiftKey, e.ctrlKey);
                e.preventDefault();
            } else if (e.which == 46 || e.which == 8) {
                // Delete
                if (libraryBase.jspreadsheet.current.options.editable != false) {
                    if (libraryBase.jspreadsheet.current.selection.row != null) {
                        if (libraryBase.jspreadsheet.current.options.allowDeleteRow != false) {
                            if (confirm(jSuites.translate('Are you sure to delete the selected rows?'))) {
                                libraryBase.jspreadsheet.current.deleteRow();
                            }
                        }
                    } else if (libraryBase.jspreadsheet.current.selection.column != null) {
                        if (libraryBase.jspreadsheet.current.options.allowDeleteColumn != false) {
                            if (confirm(jSuites.translate('Are you sure to delete the selected columns?'))) {
                                libraryBase.jspreadsheet.current.deleteColumn();
                            }
                        }
                    } else {
                        // Change value
                        libraryBase.jspreadsheet.current.setValue(
                            libraryBase.jspreadsheet.current.selection.highlighted.map(function (record) {
                                return record.element;
                            }),
                            ''
                        );
                    }
                }
            } else if (e.which == 13) {
                // Move cursor
                if (e.shiftKey) {
                    up.call(libraryBase.jspreadsheet.current);
                } else {
                    if (libraryBase.jspreadsheet.current.options.allowInsertRow != false) {
                        if (libraryBase.jspreadsheet.current.options.allowManualInsertRow != false) {
                            if (libraryBase.jspreadsheet.current.selection.range[1] == libraryBase.jspreadsheet.current.options.data.length - 1) {
                                // New record in case the selection is in the last row
                                libraryBase.jspreadsheet.current.insertRow();
                            }
                        }
                    }

                    down.call(libraryBase.jspreadsheet.current);
                }
                e.preventDefault();
            } else if (e.which == 9) {
                // Tab
                if (e.shiftKey) {
                    left.call(libraryBase.jspreadsheet.current);
                } else {
                    if (libraryBase.jspreadsheet.current.options.allowInsertColumn != false) {
                        if (libraryBase.jspreadsheet.current.options.allowManualInsertColumn != false) {
                            if (libraryBase.jspreadsheet.current.selection.range[0] == libraryBase.jspreadsheet.current.options.data[0].length - 1) {
                                // New record in case the selection is in the last column
                                libraryBase.jspreadsheet.current.insertColumn();
                            }
                        }
                    }

                    right.call(libraryBase.jspreadsheet.current);
                }
                e.preventDefault();
            } else {
                if ((e.ctrlKey || e.metaKey) && !e.shiftKey) {
                    if (e.which == 65) {
                        // Ctrl + A
                        selectAll.call(libraryBase.jspreadsheet.current);
                        e.preventDefault();
                    } else if (e.which == 83) {
                        // Ctrl + S
                        libraryBase.jspreadsheet.current.download();
                        e.preventDefault();
                    } else if (e.which == 89) {
                        // Ctrl + Y
                        libraryBase.jspreadsheet.current.redo();
                        e.preventDefault();
                    } else if (e.which == 90) {
                        // Ctrl + Z
                        libraryBase.jspreadsheet.current.undo();
                        e.preventDefault();
                    } else if (e.which == 67) {
                        // Ctrl + C
                        copy.call(libraryBase.jspreadsheet.current, true);
                        e.preventDefault();
                    } else if (e.which == 88) {
                        // Ctrl + X
                        if (libraryBase.jspreadsheet.current.options.editable != false) {
                            cutControls();
                        } else {
                            copyControls();
                        }
                        e.preventDefault();
                    } else if (e.which == 86) {
                        // Ctrl + V
                        pasteControls();
                    }
                } else {
                    if (libraryBase.jspreadsheet.current.selection.range) {
                        if (libraryBase.jspreadsheet.current.options.editable != false) {
                            const rowId = libraryBase.jspreadsheet.current.selection.range[1];
                            const columnId = libraryBase.jspreadsheet.current.selection.range[0];
                            const column =
                                libraryBase.jspreadsheet.current.options.columns[getCellColumnIndex.call(libraryBase.jspreadsheet.current, columnId, rowId)];

                            // Characters able to start a edition
                            if (e.keyCode == 32) {
                                // Space
                                e.preventDefault();
                                if (column.type == 'checkbox' || column.type == 'radio') {
                                    setCheckRadioValue.call(libraryBase.jspreadsheet.current);
                                } else {
                                    // Start edition
                                    openEditor.call(
                                        libraryBase.jspreadsheet.current,
                                        libraryBase.jspreadsheet.current.records[rowId][columnId].element,
                                        true,
                                        e
                                    );
                                }
                            } else if (e.keyCode == 113) {
                                // Start edition with current content F2
                                openEditor.call(libraryBase.jspreadsheet.current, libraryBase.jspreadsheet.current.records[rowId][columnId].element, false, e);
                            } else if ((e.key.length === 1 || e.key === 'Process') && !(e.altKey || isCtrl(e))) {
                                // Start edition
                                openEditor.call(libraryBase.jspreadsheet.current, libraryBase.jspreadsheet.current.records[rowId][columnId].element, true, e);
                                // Prevent entries in the calendar
                                if (column && column.type == 'calendar') {
                                    e.preventDefault();
                                }
                            }
                        }
                    }
                }
            }
        } else {
            if (e.target.classList.contains('jss_search')) {
                if (libraryBase.jspreadsheet.timeControl) {
                    clearTimeout(libraryBase.jspreadsheet.timeControl);
                }

                libraryBase.jspreadsheet.timeControl = setTimeout(function () {
                    libraryBase.jspreadsheet.current!.search((e.target as HTMLInputElement).value);
                }, 200);
            }
        }
    }
};
