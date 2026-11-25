import { closeEditor, openEditor } from '../utils/editor.js';
import libraryBase from '../utils/libraryBase.js';
import { isColMerged, isRowMerged } from '../utils/merges.js';
import { copyData, resetSelection, selectAll, updateCornerPosition, updateSelectionFromCoords } from '../utils/selection.js';
import { openFilter } from '../utils/filter.js';
import { setWidth } from '../utils/columns.js';
import { moveRow, setHeight } from '../utils/rows.js';
import { getElement, getIndex } from './helpers.js';
import type { TargetEvent, Worksheet } from '../types/internal';

export const mouseUpControls = function (e: TargetEvent<MouseEvent>) {
    if (libraryBase.jspreadsheet.current) {
        // Update cell size
        if (libraryBase.jspreadsheet.current.pointer.resizing) {
            // Columns to be updated
            if (libraryBase.jspreadsheet.current.pointer.resizing.column != null) {
                // New width
                const newWidth = parseInt(
                    libraryBase.jspreadsheet.current.cols[libraryBase.jspreadsheet.current.pointer.resizing.column].colElement.getAttribute('width')
                );
                // Columns
                const columns = libraryBase.jspreadsheet.current.getSelectedColumns();
                if (columns.length > 1) {
                    const currentWidth = [];
                    for (let i = 0; i < columns.length; i++) {
                        currentWidth.push(parseInt(libraryBase.jspreadsheet.current.cols[columns[i]].colElement.getAttribute('width')));
                    }
                    // Previous width
                    const index = columns.indexOf(parseInt(libraryBase.jspreadsheet.current.pointer.resizing.column));
                    currentWidth[index] = libraryBase.jspreadsheet.current.pointer.resizing.width;
                    setWidth.call(libraryBase.jspreadsheet.current, columns, newWidth, currentWidth);
                } else {
                    setWidth.call(
                        libraryBase.jspreadsheet.current,
                        parseInt(libraryBase.jspreadsheet.current.pointer.resizing.column),
                        newWidth,
                        libraryBase.jspreadsheet.current.pointer.resizing.width
                    );
                }
                // Remove border
                libraryBase.jspreadsheet.current.headers[libraryBase.jspreadsheet.current.pointer.resizing.column].classList.remove('resizing');
                for (let j = 0; j < libraryBase.jspreadsheet.current.records.length; j++) {
                    if (libraryBase.jspreadsheet.current.records[j][libraryBase.jspreadsheet.current.pointer.resizing.column]) {
                        libraryBase.jspreadsheet.current.records[j][libraryBase.jspreadsheet.current.pointer.resizing.column].element.classList.remove(
                            'resizing'
                        );
                    }
                }
            } else {
                // Remove Class
                libraryBase.jspreadsheet.current.rows[libraryBase.jspreadsheet.current.pointer.resizing.row].element.children[0].classList.remove('resizing');
                // setHeight parses the height, and it ignores a missing attribute (NaN) as it ignored null
                const newHeight = parseInt(
                    libraryBase.jspreadsheet.current.rows[libraryBase.jspreadsheet.current.pointer.resizing.row].element.getAttribute('height')
                );
                setHeight.call(
                    libraryBase.jspreadsheet.current,
                    libraryBase.jspreadsheet.current.pointer.resizing.row,
                    newHeight,
                    libraryBase.jspreadsheet.current.pointer.resizing.height
                );
                // Remove border
                libraryBase.jspreadsheet.current.pointer.resizing.element.classList.remove('resizing');
            }
            // Reset resizing helper
            libraryBase.jspreadsheet.current.pointer.resizing = null;
        } else if (libraryBase.jspreadsheet.current.pointer.dragging) {
            // Reset dragging helper
            if (libraryBase.jspreadsheet.current.pointer.dragging) {
                if (libraryBase.jspreadsheet.current.pointer.dragging.column != null) {
                    // Target
                    const columnId = getIndex(e.target, 'data-x');
                    // Remove move style
                    libraryBase.jspreadsheet.current.headers[libraryBase.jspreadsheet.current.pointer.dragging.column].classList.remove('dragging');
                    for (let j = 0; j < libraryBase.jspreadsheet.current.rows.length; j++) {
                        if (libraryBase.jspreadsheet.current.records[j][libraryBase.jspreadsheet.current.pointer.dragging.column]) {
                            libraryBase.jspreadsheet.current.records[j][libraryBase.jspreadsheet.current.pointer.dragging.column].element.classList.remove(
                                'dragging'
                            );
                        }
                    }
                    for (let i = 0; i < libraryBase.jspreadsheet.current.headers.length; i++) {
                        libraryBase.jspreadsheet.current.headers[i].classList.remove('dragging-left');
                        libraryBase.jspreadsheet.current.headers[i].classList.remove('dragging-right');
                    }
                    // Update position
                    if (columnId != null) {
                        if (libraryBase.jspreadsheet.current.pointer.dragging.column != libraryBase.jspreadsheet.current.pointer.dragging.destination) {
                            libraryBase.jspreadsheet.current.moveColumn(
                                libraryBase.jspreadsheet.current.pointer.dragging.column,
                                libraryBase.jspreadsheet.current.pointer.dragging.destination
                            );
                        }
                    }
                } else {
                    // A row that moved has a sibling. A row that did not move can be the only row.
                    if (libraryBase.jspreadsheet.current.pointer.dragging.row != libraryBase.jspreadsheet.current.pointer.dragging.destination) {
                        let position;

                        if (libraryBase.jspreadsheet.current.pointer.dragging.element.nextSibling) {
                            position = parseInt((libraryBase.jspreadsheet.current.pointer.dragging.element.nextSibling as HTMLElement).getAttribute('data-y'));
                            if (libraryBase.jspreadsheet.current.pointer.dragging.row < position) {
                                position -= 1;
                            }
                        } else {
                            position = parseInt(
                                (libraryBase.jspreadsheet.current.pointer.dragging.element.previousSibling as HTMLElement).getAttribute('data-y')
                            );
                        }

                        moveRow.call(libraryBase.jspreadsheet.current, libraryBase.jspreadsheet.current.pointer.dragging.row, position, true);
                    }
                    libraryBase.jspreadsheet.current.pointer.dragging.element.classList.remove('dragging');
                }
                libraryBase.jspreadsheet.current.pointer.dragging = null;
            }
        } else {
            // Close any corner selection
            if (libraryBase.jspreadsheet.current.pointer.corner) {
                libraryBase.jspreadsheet.current.pointer.corner = false;

                // Data to be copied
                if (libraryBase.jspreadsheet.current.selection.fill.length > 0) {
                    // Copy data
                    copyData.call(
                        libraryBase.jspreadsheet.current,
                        libraryBase.jspreadsheet.current.selection.fill[0],
                        libraryBase.jspreadsheet.current.selection.fill[libraryBase.jspreadsheet.current.selection.fill.length - 1]
                    );

                    // Remove selection
                    libraryBase.jspreadsheet.current.selection.clearFill();
                }
            }
        }
    }

    // Clear any time control
    if (libraryBase.jspreadsheet.timeControl) {
        clearTimeout(libraryBase.jspreadsheet.timeControl);
        libraryBase.jspreadsheet.timeControl = null;
    }

    // Mouse up
    libraryBase.jspreadsheet.isMouseAction = false;
};

export const mouseDownControls = function (e: TargetEvent<MouseEvent>) {
    e = e || window.event;

    let mouseButton;

    if (e.buttons) {
        mouseButton = e.buttons;
    } else if (e.button) {
        mouseButton = e.button;
    } else {
        mouseButton = e.which;
    }

    // Get elements
    const jssTable = getElement(e.target);

    if (jssTable[0]) {
        if (libraryBase.jspreadsheet.current != jssTable[0].jssWorksheet) {
            if (libraryBase.jspreadsheet.current) {
                if (libraryBase.jspreadsheet.current.editor.isOpen()) {
                    closeEditor.call(libraryBase.jspreadsheet.current, libraryBase.jspreadsheet.current.editor.cell, true);
                }
                libraryBase.jspreadsheet.current.resetSelection();
            }
            libraryBase.jspreadsheet.current = jssTable[0].jssWorksheet;
        }
    } else {
        if (libraryBase.jspreadsheet.current) {
            if (libraryBase.jspreadsheet.current.editor.isOpen()) {
                closeEditor.call(libraryBase.jspreadsheet.current, libraryBase.jspreadsheet.current.editor.cell, true);
            }

            if (!e.target.closest('.jss_object')) {
                resetSelection.call(libraryBase.jspreadsheet.current, true);
                libraryBase.jspreadsheet.current = null;
            }
        }
    }

    if (libraryBase.jspreadsheet.current && mouseButton == 1) {
        if (e.target.classList.contains('jss_selectall')) {
            if (libraryBase.jspreadsheet.current) {
                selectAll.call(libraryBase.jspreadsheet.current);
            }
        } else if (e.target.classList.contains('jss_corner')) {
            if (libraryBase.jspreadsheet.current.options.editable != false) {
                libraryBase.jspreadsheet.current.pointer.corner = true;
            }
        } else {
            // Header found
            if (jssTable[1] == 1) {
                const columnText = e.target.getAttribute('data-x');
                if (columnText) {
                    const columnId = parseInt(columnText);
                    // Update cursor
                    const info = e.target.getBoundingClientRect();
                    if (libraryBase.jspreadsheet.current.options.columnResize != false && info.width - e.offsetX < 6) {
                        // Resize helper
                        libraryBase.jspreadsheet.current.pointer.resizing = {
                            mousePosition: e.pageX,
                            column: columnId,
                            width: info.width,
                        };

                        // Border indication
                        libraryBase.jspreadsheet.current.headers[columnId].classList.add('resizing');
                        for (let j = 0; j < libraryBase.jspreadsheet.current.records.length; j++) {
                            if (libraryBase.jspreadsheet.current.records[j][columnId]) {
                                libraryBase.jspreadsheet.current.records[j][columnId].element.classList.add('resizing');
                            }
                        }
                    } else if (libraryBase.jspreadsheet.current.options.columnDrag != false && info.height - e.offsetY < 6) {
                        if (isColMerged.call(libraryBase.jspreadsheet.current, columnId).length) {
                            console.error('Jspreadsheet: This column is part of a merged cell.');
                        } else {
                            // Reset selection
                            libraryBase.jspreadsheet.current.resetSelection();
                            // Drag helper
                            libraryBase.jspreadsheet.current.pointer.dragging = {
                                element: e.target as HTMLTableCellElement,
                                column: columnId,
                                destination: columnId,
                            };
                            // Border indication
                            libraryBase.jspreadsheet.current.headers[columnId].classList.add('dragging');
                            for (let j = 0; j < libraryBase.jspreadsheet.current.records.length; j++) {
                                if (libraryBase.jspreadsheet.current.records[j][columnId]) {
                                    libraryBase.jspreadsheet.current.records[j][columnId].element.classList.add('dragging');
                                }
                            }
                        }
                    } else {
                        let o, d;

                        if (libraryBase.jspreadsheet.current.selection.column != null && (e.shiftKey || e.ctrlKey)) {
                            o = libraryBase.jspreadsheet.current.selection.column;
                            d = columnId;
                        } else {
                            // Press to rename
                            if (
                                libraryBase.jspreadsheet.current.selection.column == columnId &&
                                libraryBase.jspreadsheet.current.options.allowRenameColumn != false
                            ) {
                                libraryBase.jspreadsheet.timeControl = setTimeout(function () {
                                    libraryBase.jspreadsheet.current!.setHeader(columnId);
                                }, 800);
                            }

                            // Keep track of which header was selected first
                            libraryBase.jspreadsheet.current.selection.column = columnId;

                            // Update selection single column
                            o = columnId;
                            d = columnId;
                        }

                        // Update selection
                        updateSelectionFromCoords.call(libraryBase.jspreadsheet.current, o, 0, d, libraryBase.jspreadsheet.current.options.data.length - 1, e);
                    }
                } else {
                    if ((e.target.parentNode as HTMLElement).classList.contains('jss_nested')) {
                        let c1, c2;

                        if (e.target.getAttribute('data-column')) {
                            const column = e.target.getAttribute('data-column')!.split(',');
                            c1 = parseInt(column[0]);
                            c2 = parseInt(column[column.length - 1]);
                        } else {
                            c1 = 0;
                            c2 = libraryBase.jspreadsheet.current.options.columns.length - 1;
                        }
                        updateSelectionFromCoords.call(
                            libraryBase.jspreadsheet.current,
                            c1,
                            0,
                            c2,
                            libraryBase.jspreadsheet.current.options.data.length - 1,
                            e
                        );
                    }
                }
            } else {
                libraryBase.jspreadsheet.current.selection.column = null;
            }

            // Body found
            if (jssTable[1] == 2) {
                const rowId = parseInt(e.target.getAttribute('data-y'));

                if (e.target.classList.contains('jss_row')) {
                    const info = e.target.getBoundingClientRect();
                    if (libraryBase.jspreadsheet.current.options.rowResize != false && info.height - e.offsetY < 6) {
                        // Resize helper
                        libraryBase.jspreadsheet.current.pointer.resizing = {
                            element: e.target.parentNode as HTMLTableRowElement,
                            mousePosition: e.pageY,
                            row: rowId,
                            height: info.height,
                        };
                        // Border indication
                        (e.target.parentNode as HTMLElement).classList.add('resizing');
                    } else if (libraryBase.jspreadsheet.current.options.rowDrag != false && info.width - e.offsetX < 6) {
                        if (isRowMerged.call(libraryBase.jspreadsheet.current, rowId).length) {
                            console.error('Jspreadsheet: This row is part of a merged cell');
                        } else if (libraryBase.jspreadsheet.current.options.search == true && libraryBase.jspreadsheet.current.view.results) {
                            console.error('Jspreadsheet: Please clear your search before perform this action');
                        } else {
                            // Reset selection
                            libraryBase.jspreadsheet.current.resetSelection();
                            // Drag helper
                            libraryBase.jspreadsheet.current.pointer.dragging = {
                                element: e.target.parentNode as HTMLTableRowElement,
                                row: rowId,
                                destination: rowId,
                            };
                            // Border indication
                            (e.target.parentNode as HTMLElement).classList.add('dragging');
                        }
                    } else {
                        let o, d;
                        if (libraryBase.jspreadsheet.current.selection.row != null && (e.shiftKey || e.ctrlKey)) {
                            o = libraryBase.jspreadsheet.current.selection.row;
                            d = rowId;
                        } else {
                            // Keep track of which header was selected first
                            libraryBase.jspreadsheet.current.selection.row = rowId;

                            // Update selection single column
                            o = rowId;
                            d = rowId;
                        }

                        // Update selection
                        updateSelectionFromCoords.call(libraryBase.jspreadsheet.current, null, o, null, d, e);
                    }
                } else {
                    // Jclose
                    if (e.target.classList.contains('jclose') && e.target.clientWidth - e.offsetX < 50 && e.offsetY < 50) {
                        closeEditor.call(libraryBase.jspreadsheet.current, libraryBase.jspreadsheet.current.editor.cell!, true);
                    } else {
                        const getCellCoords = function (element: HTMLElement) {
                            const x = element.getAttribute('data-x');
                            const y = element.getAttribute('data-y');
                            if (x && y) {
                                return [parseInt(x), parseInt(y)];
                            } else {
                                if (element.parentNode) {
                                    return getCellCoords(element.parentNode as HTMLElement);
                                }
                            }
                        };

                        const position = getCellCoords(e.target);
                        if (position) {
                            const columnId = position[0];
                            const rowId = position[1];
                            // Close edition
                            if (libraryBase.jspreadsheet.current.editor.isOpen()) {
                                if (libraryBase.jspreadsheet.current.editor.x != columnId || libraryBase.jspreadsheet.current.editor.y != rowId) {
                                    closeEditor.call(libraryBase.jspreadsheet.current, libraryBase.jspreadsheet.current.editor.cell, true);
                                }
                            }

                            if (!libraryBase.jspreadsheet.current.editor.isOpen()) {
                                // Update cell selection
                                if (e.shiftKey) {
                                    updateSelectionFromCoords.call(
                                        libraryBase.jspreadsheet.current,
                                        libraryBase.jspreadsheet.current.selection.range![0],
                                        libraryBase.jspreadsheet.current.selection.range![1],
                                        columnId,
                                        rowId,
                                        e
                                    );
                                } else {
                                    updateSelectionFromCoords.call(libraryBase.jspreadsheet.current, columnId, rowId, undefined, undefined, e);
                                }
                            }

                            // No full row selected
                            libraryBase.jspreadsheet.current.selection.column = null;
                            libraryBase.jspreadsheet.current.selection.row = null;
                        }
                    }
                }
            } else {
                libraryBase.jspreadsheet.current.selection.row = null;
            }

            // Pagination
            if (e.target.classList.contains('jss_page')) {
                if (e.target.textContent == '<') {
                    libraryBase.jspreadsheet.current.page(0);
                } else if (e.target.textContent == '>') {
                    libraryBase.jspreadsheet.current.page(parseInt(e.target.getAttribute('title')) - 1);
                } else {
                    libraryBase.jspreadsheet.current.page(parseInt(e.target.textContent) - 1);
                }
            }
        }

        if (libraryBase.jspreadsheet.current.editor.isOpen()) {
            libraryBase.jspreadsheet.isMouseAction = false;
        } else {
            libraryBase.jspreadsheet.isMouseAction = true;
        }
    } else {
        libraryBase.jspreadsheet.isMouseAction = false;
    }
};

// Mouse move controls
export const mouseMoveControls = function (e: TargetEvent<MouseEvent>) {
    e = e || window.event;

    let mouseButton;

    if (e.buttons) {
        mouseButton = e.buttons;
    } else if (e.button) {
        mouseButton = e.button;
    } else {
        mouseButton = e.which;
    }

    if (!mouseButton) {
        libraryBase.jspreadsheet.isMouseAction = false;
    }

    if (libraryBase.jspreadsheet.current) {
        if (libraryBase.jspreadsheet.isMouseAction == true) {
            // Resizing is ongoing
            if (libraryBase.jspreadsheet.current.pointer.resizing) {
                if (libraryBase.jspreadsheet.current.pointer.resizing.column != null) {
                    const width = e.pageX - libraryBase.jspreadsheet.current.pointer.resizing.mousePosition;

                    if (libraryBase.jspreadsheet.current.pointer.resizing.width + width > 0) {
                        const tempWidth = libraryBase.jspreadsheet.current.pointer.resizing.width + width;
                        libraryBase.jspreadsheet.current.cols[libraryBase.jspreadsheet.current.pointer.resizing.column].colElement.setAttribute(
                            'width',
                            tempWidth
                        );

                        updateCornerPosition.call(libraryBase.jspreadsheet.current);
                    }
                } else {
                    const height = e.pageY - libraryBase.jspreadsheet.current.pointer.resizing.mousePosition;

                    if (libraryBase.jspreadsheet.current.pointer.resizing.height + height > 0) {
                        const tempHeight = libraryBase.jspreadsheet.current.pointer.resizing.height + height;
                        libraryBase.jspreadsheet.current.rows[libraryBase.jspreadsheet.current.pointer.resizing.row].element.setAttribute('height', tempHeight);

                        updateCornerPosition.call(libraryBase.jspreadsheet.current);
                    }
                }
            } else if (libraryBase.jspreadsheet.current.pointer.dragging) {
                if (libraryBase.jspreadsheet.current.pointer.dragging.column != null) {
                    const columnId = getIndex(e.target, 'data-x');
                    if (columnId != null) {
                        if (isColMerged.call(libraryBase.jspreadsheet.current, columnId).length) {
                            console.error('Jspreadsheet: This column is part of a merged cell.');
                        } else {
                            for (let i = 0; i < libraryBase.jspreadsheet.current.headers.length; i++) {
                                libraryBase.jspreadsheet.current.headers[i].classList.remove('dragging-left');
                                libraryBase.jspreadsheet.current.headers[i].classList.remove('dragging-right');
                            }

                            if (libraryBase.jspreadsheet.current.pointer.dragging.column == columnId) {
                                libraryBase.jspreadsheet.current.pointer.dragging.destination = columnId;
                            } else {
                                if (e.target.clientWidth / 2 > e.offsetX) {
                                    if (libraryBase.jspreadsheet.current.pointer.dragging.column < columnId) {
                                        libraryBase.jspreadsheet.current.pointer.dragging.destination = columnId - 1;
                                    } else {
                                        libraryBase.jspreadsheet.current.pointer.dragging.destination = columnId;
                                    }
                                    libraryBase.jspreadsheet.current.headers[columnId].classList.add('dragging-left');
                                } else {
                                    if (libraryBase.jspreadsheet.current.pointer.dragging.column < columnId) {
                                        libraryBase.jspreadsheet.current.pointer.dragging.destination = columnId;
                                    } else {
                                        libraryBase.jspreadsheet.current.pointer.dragging.destination = columnId + 1;
                                    }
                                    libraryBase.jspreadsheet.current.headers[columnId].classList.add('dragging-right');
                                }
                            }
                        }
                    }
                } else {
                    const rowId = getIndex(e.target, 'data-y');
                    if (rowId != null) {
                        if (isRowMerged.call(libraryBase.jspreadsheet.current, rowId).length) {
                            console.error('Jspreadsheet: This row is part of a merged cell.');
                        } else {
                            const target = e.target.clientHeight / 2 > e.offsetY ? e.target.parentNode!.nextSibling : e.target.parentNode;
                            if (libraryBase.jspreadsheet.current.pointer.dragging.element != target) {
                                e.target.parentNode!.parentNode!.insertBefore(libraryBase.jspreadsheet.current.pointer.dragging.element, target);
                                libraryBase.jspreadsheet.current.pointer.dragging.destination = Array.prototype.indexOf.call(
                                    libraryBase.jspreadsheet.current.pointer.dragging.element.parentNode!.children,
                                    libraryBase.jspreadsheet.current.pointer.dragging.element
                                );
                            }
                        }
                    }
                }
            }
        } else {
            const x = e.target.getAttribute('data-x');
            const y = e.target.getAttribute('data-y');
            const rect = e.target.getBoundingClientRect();

            if (libraryBase.jspreadsheet.current.pointer.cursor) {
                libraryBase.jspreadsheet.current.pointer.cursor.style.cursor = '';
                libraryBase.jspreadsheet.current.pointer.cursor = null;
            }

            if (e.target.parentNode!.parentNode && (e.target.parentNode!.parentNode as HTMLElement).className) {
                if ((e.target.parentNode!.parentNode as HTMLElement).classList.contains('resizable')) {
                    if (e.target && x && !y && rect.width - (e.clientX - rect.left) < 6) {
                        libraryBase.jspreadsheet.current.pointer.cursor = e.target;
                        libraryBase.jspreadsheet.current.pointer.cursor.style.cursor = 'col-resize';
                    } else if (e.target && !x && y && rect.height - (e.clientY - rect.top) < 6) {
                        libraryBase.jspreadsheet.current.pointer.cursor = e.target;
                        libraryBase.jspreadsheet.current.pointer.cursor.style.cursor = 'row-resize';
                    }
                }

                if ((e.target.parentNode!.parentNode as HTMLElement).classList.contains('draggable')) {
                    if (e.target && !x && y && rect.width - (e.clientX - rect.left) < 6) {
                        libraryBase.jspreadsheet.current.pointer.cursor = e.target;
                        libraryBase.jspreadsheet.current.pointer.cursor.style.cursor = 'move';
                    } else if (e.target && x && !y && rect.height - (e.clientY - rect.top) < 6) {
                        libraryBase.jspreadsheet.current.pointer.cursor = e.target;
                        libraryBase.jspreadsheet.current.pointer.cursor.style.cursor = 'move';
                    }
                }
            }
        }
    }
};

/**
 * Update copy selection
 *
 * @param int x, y
 * @return void
 */
const updateCopySelection = function (this: Worksheet, x3: number, y3: number) {
    const obj = this;

    // Remove selection
    obj.selection.clearFill();

    // Get elements first and last
    const x1 = obj.selection.bounds![0];
    const y1 = obj.selection.bounds![1];
    const x2 = obj.selection.bounds![2];
    const y2 = obj.selection.bounds![3];

    if (x3 != null && y3 != null) {
        let px, ux;

        if (x3 - x2 > 0) {
            px = parseInt(x2) + 1;
            ux = parseInt(x3);
        } else {
            px = parseInt(x3);
            ux = parseInt(x1) - 1;
        }

        let py, uy;

        if (y3 - y2 > 0) {
            py = parseInt(y2) + 1;
            uy = parseInt(y3);
        } else {
            py = parseInt(y3);
            uy = parseInt(y1) - 1;
        }

        if (ux - px <= uy - py) {
            px = parseInt(x1);
            ux = parseInt(x2);
        } else {
            py = parseInt(y1);
            uy = parseInt(y2);
        }

        for (let j = py; j <= uy; j++) {
            for (let i = px; i <= ux; i++) {
                if (obj.records[j][i] && obj.rows[j].element.style.display != 'none' && obj.records[j][i].element.style.display != 'none') {
                    obj.records[j][i].element.classList.add('selection');
                    obj.records[py][i].element.classList.add('selection-top');
                    obj.records[uy][i].element.classList.add('selection-bottom');
                    obj.records[j][px].element.classList.add('selection-left');
                    obj.records[j][ux].element.classList.add('selection-right');

                    // Persist selected elements
                    obj.selection.fill.push(obj.records[j][i].element);
                }
            }
        }
    }
};

export const mouseOverControls = function (e: TargetEvent<MouseEvent>) {
    e = e || window.event;

    let mouseButton;

    if (e.buttons) {
        mouseButton = e.buttons;
    } else if (e.button) {
        mouseButton = e.button;
    } else {
        mouseButton = e.which;
    }

    if (!mouseButton) {
        libraryBase.jspreadsheet.isMouseAction = false;
    }

    if (libraryBase.jspreadsheet.current && libraryBase.jspreadsheet.isMouseAction == true) {
        // Get elements
        const jssTable = getElement(e.target);

        if (jssTable[0]) {
            // Avoid cross reference
            if (libraryBase.jspreadsheet.current != jssTable[0].jssWorksheet) {
                if (libraryBase.jspreadsheet.current) {
                    return false;
                }
            }

            let columnId = getIndex(e.target, 'data-x');
            const rowId = getIndex(e.target, 'data-y');
            if (libraryBase.jspreadsheet.current.pointer.resizing || libraryBase.jspreadsheet.current.pointer.dragging) {
                // ignore
            } else {
                // Header found
                if (jssTable[1] == 1) {
                    if (libraryBase.jspreadsheet.current.selection.column != null) {
                        columnId = getIndex(e.target, 'data-x');
                        const o = libraryBase.jspreadsheet.current.selection.column;
                        const d = columnId;
                        // Update selection
                        updateSelectionFromCoords.call(libraryBase.jspreadsheet.current, o, 0, d, libraryBase.jspreadsheet.current.options.data.length - 1, e);
                    }
                }

                // Body found
                if (jssTable[1] == 2) {
                    if (e.target.classList.contains('jss_row')) {
                        if (libraryBase.jspreadsheet.current.selection.row != null) {
                            const o = libraryBase.jspreadsheet.current.selection.row;
                            const d = rowId;
                            // Update selection
                            updateSelectionFromCoords.call(
                                libraryBase.jspreadsheet.current,
                                0,
                                o,
                                libraryBase.jspreadsheet.current.options.data[0].length - 1,
                                d,
                                e
                            );
                        }
                    } else {
                        // Do not select edtion is in progress
                        if (!libraryBase.jspreadsheet.current.editor.isOpen()) {
                            if (columnId != null && rowId != null) {
                                if (libraryBase.jspreadsheet.current.pointer.corner) {
                                    updateCopySelection.call(libraryBase.jspreadsheet.current, columnId, rowId);
                                } else {
                                    if (libraryBase.jspreadsheet.current.selection.range) {
                                        updateSelectionFromCoords.call(
                                            libraryBase.jspreadsheet.current,
                                            libraryBase.jspreadsheet.current.selection.range[0],
                                            libraryBase.jspreadsheet.current.selection.range[1],
                                            columnId,
                                            rowId,
                                            e
                                        );
                                    }
                                }
                            }
                        }
                    }
                }
            }
        }
    }

    // Clear any time control
    if (libraryBase.jspreadsheet.timeControl) {
        clearTimeout(libraryBase.jspreadsheet.timeControl);
        libraryBase.jspreadsheet.timeControl = null;
    }
};

export const doubleClickControls = function (e: TargetEvent<MouseEvent>) {
    // Jss is selected
    if (libraryBase.jspreadsheet.current) {
        // Corner action
        if (e.target.classList.contains('jss_corner')) {
            // Any selected cells
            if (libraryBase.jspreadsheet.current.selection.highlighted.length > 0) {
                // Copy from this
                const x1 = parseInt(libraryBase.jspreadsheet.current.selection.highlighted[0].element.getAttribute('data-x'));
                const y1 =
                    parseInt(
                        libraryBase.jspreadsheet.current.selection.highlighted[
                            libraryBase.jspreadsheet.current.selection.highlighted.length - 1
                        ].element.getAttribute('data-y')
                    ) + 1;
                // Until this
                const x2 = parseInt(
                    libraryBase.jspreadsheet.current.selection.highlighted[
                        libraryBase.jspreadsheet.current.selection.highlighted.length - 1
                    ].element.getAttribute('data-x')
                );
                const y2 = libraryBase.jspreadsheet.current.records.length - 1;
                // Execute copy
                copyData.call(
                    libraryBase.jspreadsheet.current,
                    libraryBase.jspreadsheet.current.records[y1][x1].element,
                    libraryBase.jspreadsheet.current.records[y2][x2].element
                );
            }
        } else if (e.target.classList.contains('jss_column_filter')) {
            // Column
            const columnId = parseInt(e.target.getAttribute('data-x'));
            // Open filter
            openFilter.call(libraryBase.jspreadsheet.current, columnId);
        } else {
            // Get table
            const jssTable = getElement(e.target);

            // Double click over header
            if (jssTable[1] == 1 && libraryBase.jspreadsheet.current.options.columnSorting != false) {
                // Check valid column header coords
                const columnId = getIndex(e.target, 'data-x');
                if (columnId != null) {
                    libraryBase.jspreadsheet.current.orderBy(columnId);
                }
            }

            // Double click over body
            if (jssTable[1] == 2 && libraryBase.jspreadsheet.current.options.editable != false) {
                if (!libraryBase.jspreadsheet.current.editor.isOpen()) {
                    const getCellCoords = function (element: HTMLElement) {
                        if (element.parentNode) {
                            const x = element.getAttribute('data-x');
                            const y = element.getAttribute('data-y');
                            if (x && y) {
                                return element;
                            } else {
                                return getCellCoords(element.parentNode as HTMLElement);
                            }
                        }
                    };
                    const cell = getCellCoords(e.target);
                    if (cell && cell.classList.contains('highlight')) {
                        openEditor.call(libraryBase.jspreadsheet.current, cell as HTMLTableCellElement, undefined, e);
                    }
                }
            }
        }
    }
};
