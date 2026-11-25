import { closeEditor, openEditor } from '../utils/editor.js';
import libraryBase from '../utils/libraryBase.js';
import { isColMerged, isRowMerged } from '../utils/merges.js';
import { resetSelection, selectAll, updateCornerPosition, updateSelectionFromCoords } from '../utils/selection.js';
import { copyData } from '../utils/fill.js';
import { openFilter } from '../utils/filter.js';
import { setWidth } from '../utils/columns.js';
import { moveRow, setHeight } from '../utils/rows.js';
import { getElement, getIndex } from './helpers.js';
import type { TargetEvent, Worksheet } from '../types/internal';

/**
 * The mouse handlers act on the active worksheet, jspreadsheet.current. A press on a worksheet makes it the active
 * worksheet (mouseDownControls), and the other handlers read it once. Each gesture has its own function, with the
 * worksheet as `this`: a press on a header, a row number or a cell, the resize or the move of a column or a row, and
 * the fill with the selection corner.
 */

/** The button that is down: buttons, or button, or which in old browsers. */
const getMouseButton = function (e: MouseEvent) {
    if (e.buttons) {
        return e.buttons;
    } else if (e.button) {
        return e.button;
    } else {
        return e.which;
    }
};

/** The column and the row of the cell that contains the element, from its data-x and data-y attributes. */
const getCellPosition = function (element: HTMLElement): number[] | undefined {
    const x = element.getAttribute('data-x');
    const y = element.getAttribute('data-y');
    if (x && y) {
        return [parseInt(x), parseInt(y)];
    } else {
        if (element.parentNode) {
            return getCellPosition(element.parentNode as HTMLElement);
        }
    }
};

/** The cell that contains the element: the first element up the tree with data-x and data-y. */
const getCell = function (element: HTMLElement): HTMLElement | undefined {
    if (element.parentNode) {
        const x = element.getAttribute('data-x');
        const y = element.getAttribute('data-y');
        if (x && y) {
            return element;
        } else {
            return getCell(element.parentNode as HTMLElement);
        }
    }
};

/**
 * The end of the resize of a column or a row: set the width or the height that the mouse gave, with the history.
 */
const endResize = function (this: Worksheet) {
    const obj = this;
    const resizing = obj.pointer.resizing!;

    // Columns to be updated
    if (resizing.column != null) {
        // New width
        const newWidth = parseInt(obj.cols[resizing.column].colElement.getAttribute('width'));
        // Columns
        const columns = obj.getSelectedColumns();
        if (columns.length > 1) {
            const currentWidth = [];
            for (let i = 0; i < columns.length; i++) {
                currentWidth.push(parseInt(obj.cols[columns[i]].colElement.getAttribute('width')));
            }
            // Previous width
            const index = columns.indexOf(parseInt(resizing.column));
            currentWidth[index] = resizing.width;
            setWidth.call(obj, columns, newWidth, currentWidth);
        } else {
            setWidth.call(obj, parseInt(resizing.column), newWidth, resizing.width);
        }
        // Remove border
        obj.headers[resizing.column].classList.remove('resizing');
        for (let j = 0; j < obj.records.length; j++) {
            if (obj.records[j][resizing.column]) {
                obj.records[j][resizing.column].element.classList.remove('resizing');
            }
        }
    } else {
        // Remove Class
        obj.rows[resizing.row].element.children[0].classList.remove('resizing');
        // setHeight parses the height, and it ignores a missing attribute (NaN) as it ignored null
        const newHeight = parseInt(obj.rows[resizing.row].element.getAttribute('height'));
        setHeight.call(obj, resizing.row, newHeight, resizing.height);
        // Remove border
        resizing.element.classList.remove('resizing');
    }
    // Reset resizing helper
    obj.pointer.resizing = null;
};

/**
 * The end of the move of a column or a row: move it to the position that the mouse gave.
 */
const endDrag = function (this: Worksheet, e: TargetEvent<MouseEvent>) {
    const obj = this;
    const dragging = obj.pointer.dragging!;

    if (dragging.column != null) {
        // Target
        const columnId = getIndex(e.target, 'data-x');
        // Remove move style
        obj.headers[dragging.column].classList.remove('dragging');
        for (let j = 0; j < obj.rows.length; j++) {
            if (obj.records[j][dragging.column]) {
                obj.records[j][dragging.column].element.classList.remove('dragging');
            }
        }
        for (let i = 0; i < obj.headers.length; i++) {
            obj.headers[i].classList.remove('dragging-left');
            obj.headers[i].classList.remove('dragging-right');
        }
        // Update position
        if (columnId != null) {
            if (dragging.column != dragging.destination) {
                obj.moveColumn(dragging.column, dragging.destination);
            }
        }
    } else {
        // A row that moved has a sibling. A row that did not move can be the only row.
        if (dragging.row != dragging.destination) {
            let position;

            if (dragging.element.nextSibling) {
                position = parseInt((dragging.element.nextSibling as HTMLElement).getAttribute('data-y'));
                if (dragging.row < position) {
                    position -= 1;
                }
            } else {
                position = parseInt((dragging.element.previousSibling as HTMLElement).getAttribute('data-y'));
            }

            moveRow.call(obj, dragging.row, position, true);
        }
        dragging.element.classList.remove('dragging');
    }
    // Reset dragging helper
    obj.pointer.dragging = null;
};

/**
 * The end of a drag of the selection corner: fill the cells that the corner went over.
 */
const endFill = function (this: Worksheet) {
    const obj = this;

    // Close any corner selection
    if (obj.pointer.corner) {
        obj.pointer.corner = false;

        // Data to be copied
        if (obj.selection.fill.length > 0) {
            // Copy data
            copyData.call(obj, obj.selection.fill[0], obj.selection.fill[obj.selection.fill.length - 1]);

            // Remove selection
            obj.selection.clearFill();
        }
    }
};

export const mouseUpControls = function (e: TargetEvent<MouseEvent>) {
    const current = libraryBase.jspreadsheet.current;

    if (current) {
        // Update cell size
        if (current.pointer.resizing) {
            endResize.call(current);
        } else if (current.pointer.dragging) {
            endDrag.call(current, e);
        } else {
            endFill.call(current);
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

/**
 * A press on a column header starts its resize near its right edge, its move near its bottom edge, and selects the
 * column otherwise. A press on a nested header selects its columns.
 */
const pressHeader = function (this: Worksheet, e: TargetEvent<MouseEvent>) {
    const obj = this;

    const columnText = e.target.getAttribute('data-x');
    if (columnText) {
        const columnId = parseInt(columnText);
        // Update cursor
        const info = e.target.getBoundingClientRect();
        if (obj.options.columnResize != false && info.width - e.offsetX < 6) {
            // Resize helper
            obj.pointer.resizing = {
                mousePosition: e.pageX,
                column: columnId,
                width: info.width,
            };

            // Border indication
            obj.headers[columnId].classList.add('resizing');
            for (let j = 0; j < obj.records.length; j++) {
                if (obj.records[j][columnId]) {
                    obj.records[j][columnId].element.classList.add('resizing');
                }
            }
        } else if (obj.options.columnDrag != false && info.height - e.offsetY < 6) {
            if (isColMerged.call(obj, columnId).length) {
                console.error('Jspreadsheet: This column is part of a merged cell.');
            } else {
                // Reset selection
                obj.resetSelection();
                // Drag helper
                obj.pointer.dragging = {
                    element: e.target as HTMLTableCellElement,
                    column: columnId,
                    destination: columnId,
                };
                // Border indication
                obj.headers[columnId].classList.add('dragging');
                for (let j = 0; j < obj.records.length; j++) {
                    if (obj.records[j][columnId]) {
                        obj.records[j][columnId].element.classList.add('dragging');
                    }
                }
            }
        } else {
            let o, d;

            if (obj.selection.column != null && (e.shiftKey || e.ctrlKey)) {
                o = obj.selection.column;
                d = columnId;
            } else {
                // Press to rename
                if (obj.selection.column == columnId && obj.options.allowRenameColumn != false) {
                    // The active worksheet when the timer runs
                    libraryBase.jspreadsheet.timeControl = setTimeout(function () {
                        libraryBase.jspreadsheet.current!.setHeader(columnId);
                    }, 800);
                }

                // Keep track of which header was selected first
                obj.selection.column = columnId;

                // Update selection single column
                o = columnId;
                d = columnId;
            }

            // Update selection
            updateSelectionFromCoords.call(obj, o, 0, d, obj.options.data.length - 1, e);
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
                c2 = obj.options.columns.length - 1;
            }
            updateSelectionFromCoords.call(obj, c1, 0, c2, obj.options.data.length - 1, e);
        }
    }
};

/**
 * A press on a row number starts the resize of the row near its bottom edge, its move near its right edge, and selects
 * the row otherwise.
 */
const pressRowNumber = function (this: Worksheet, e: TargetEvent<MouseEvent>) {
    const obj = this;

    const rowId = parseInt(e.target.getAttribute('data-y'));
    const info = e.target.getBoundingClientRect();
    if (obj.options.rowResize != false && info.height - e.offsetY < 6) {
        // Resize helper
        obj.pointer.resizing = {
            element: e.target.parentNode as HTMLTableRowElement,
            mousePosition: e.pageY,
            row: rowId,
            height: info.height,
        };
        // Border indication
        (e.target.parentNode as HTMLElement).classList.add('resizing');
    } else if (obj.options.rowDrag != false && info.width - e.offsetX < 6) {
        if (isRowMerged.call(obj, rowId).length) {
            console.error('Jspreadsheet: This row is part of a merged cell');
        } else if (obj.options.search == true && obj.view.results) {
            console.error('Jspreadsheet: Please clear your search before perform this action');
        } else {
            // Reset selection
            obj.resetSelection();
            // Drag helper
            obj.pointer.dragging = {
                element: e.target.parentNode as HTMLTableRowElement,
                row: rowId,
                destination: rowId,
            };
            // Border indication
            (e.target.parentNode as HTMLElement).classList.add('dragging');
        }
    } else {
        let o, d;
        if (obj.selection.row != null && (e.shiftKey || e.ctrlKey)) {
            o = obj.selection.row;
            d = rowId;
        } else {
            // Keep track of which header was selected first
            obj.selection.row = rowId;

            // Update selection single column
            o = rowId;
            d = rowId;
        }

        // Update selection
        updateSelectionFromCoords.call(obj, null, o, null, d, e);
    }
};

/**
 * A press on a cell closes the editor of another cell and selects the cell, or extends the selection to it with Shift.
 * A press on the close button of an editor closes the editor.
 */
const pressCell = function (this: Worksheet, e: TargetEvent<MouseEvent>) {
    const obj = this;

    // Jclose
    if (e.target.classList.contains('jclose') && e.target.clientWidth - e.offsetX < 50 && e.offsetY < 50) {
        closeEditor.call(obj, obj.editor.cell!, true);
    } else {
        const position = getCellPosition(e.target);
        if (position) {
            const columnId = position[0];
            const rowId = position[1];
            // Close edition
            if (obj.editor.isOpen()) {
                if (obj.editor.x != columnId || obj.editor.y != rowId) {
                    closeEditor.call(obj, obj.editor.cell, true);
                }
            }

            if (!obj.editor.isOpen()) {
                // Update cell selection
                if (e.shiftKey) {
                    updateSelectionFromCoords.call(obj, obj.selection.range![0], obj.selection.range![1], columnId, rowId, e);
                } else {
                    updateSelectionFromCoords.call(obj, columnId, rowId, undefined, undefined, e);
                }
            }

            // No full row selected
            obj.selection.column = null;
            obj.selection.row = null;
        }
    }
};

/** A press on a page number of the pagination opens the page. */
const pressPage = function (this: Worksheet, e: TargetEvent<MouseEvent>) {
    const obj = this;

    if (e.target.textContent == '<') {
        obj.page(0);
    } else if (e.target.textContent == '>') {
        obj.page(parseInt(e.target.getAttribute('title')) - 1);
    } else {
        obj.page(parseInt(e.target.textContent) - 1);
    }
};

export const mouseDownControls = function (e: TargetEvent<MouseEvent>) {
    e = e || window.event;

    const mouseButton = getMouseButton(e);

    // Get elements
    const jssTable = getElement(e.target);

    // A press on a worksheet makes it the active worksheet. A press outside the spreadsheets leaves no active worksheet.
    const previous = libraryBase.jspreadsheet.current;
    if (jssTable[0]) {
        if (previous != jssTable[0].jssWorksheet) {
            if (previous) {
                if (previous.editor.isOpen()) {
                    closeEditor.call(previous, previous.editor.cell, true);
                }
                previous.resetSelection();
            }
            libraryBase.jspreadsheet.current = jssTable[0].jssWorksheet;
        }
    } else {
        if (previous) {
            if (previous.editor.isOpen()) {
                closeEditor.call(previous, previous.editor.cell, true);
            }

            if (!e.target.closest('.jss_object')) {
                resetSelection.call(previous, true);
                libraryBase.jspreadsheet.current = null;
            }
        }
    }

    const current = libraryBase.jspreadsheet.current;

    if (current && mouseButton == 1) {
        if (e.target.classList.contains('jss_selectall')) {
            selectAll.call(current);
        } else if (e.target.classList.contains('jss_corner')) {
            if (current.options.editable != false) {
                current.pointer.corner = true;
            }
        } else {
            // Header found
            if (jssTable[1] == 1) {
                pressHeader.call(current, e);
            } else {
                current.selection.column = null;
            }

            // Body found
            if (jssTable[1] == 2) {
                if (e.target.classList.contains('jss_row')) {
                    pressRowNumber.call(current, e);
                } else {
                    pressCell.call(current, e);
                }
            } else {
                current.selection.row = null;
            }

            // Pagination
            if (e.target.classList.contains('jss_page')) {
                pressPage.call(current, e);
            }
        }

        if (current.editor.isOpen()) {
            libraryBase.jspreadsheet.isMouseAction = false;
        } else {
            libraryBase.jspreadsheet.isMouseAction = true;
        }
    } else {
        libraryBase.jspreadsheet.isMouseAction = false;
    }
};

/**
 * The resize of a column or a row while the mouse moves: set the width of the col or the height of the row.
 */
const moveResize = function (this: Worksheet, e: TargetEvent<MouseEvent>) {
    const obj = this;
    const resizing = obj.pointer.resizing!;

    if (resizing.column != null) {
        const width = e.pageX - resizing.mousePosition;

        if (resizing.width + width > 0) {
            const tempWidth = resizing.width + width;
            obj.cols[resizing.column].colElement.setAttribute('width', tempWidth);

            updateCornerPosition.call(obj);
        }
    } else {
        const height = e.pageY - resizing.mousePosition;

        if (resizing.height + height > 0) {
            const tempHeight = resizing.height + height;
            obj.rows[resizing.row].element.setAttribute('height', tempHeight);

            updateCornerPosition.call(obj);
        }
    }
};

/**
 * The move of a column or a row while the mouse moves: mark the headers next to the destination of a column, or move
 * the row in the table.
 */
const moveDrag = function (this: Worksheet, e: TargetEvent<MouseEvent>) {
    const obj = this;
    const dragging = obj.pointer.dragging!;

    if (dragging.column != null) {
        const columnId = getIndex(e.target, 'data-x');
        if (columnId != null) {
            if (isColMerged.call(obj, columnId).length) {
                console.error('Jspreadsheet: This column is part of a merged cell.');
            } else {
                for (let i = 0; i < obj.headers.length; i++) {
                    obj.headers[i].classList.remove('dragging-left');
                    obj.headers[i].classList.remove('dragging-right');
                }

                if (dragging.column == columnId) {
                    dragging.destination = columnId;
                } else {
                    if (e.target.clientWidth / 2 > e.offsetX) {
                        if (dragging.column < columnId) {
                            dragging.destination = columnId - 1;
                        } else {
                            dragging.destination = columnId;
                        }
                        obj.headers[columnId].classList.add('dragging-left');
                    } else {
                        if (dragging.column < columnId) {
                            dragging.destination = columnId;
                        } else {
                            dragging.destination = columnId + 1;
                        }
                        obj.headers[columnId].classList.add('dragging-right');
                    }
                }
            }
        }
    } else {
        const rowId = getIndex(e.target, 'data-y');
        if (rowId != null) {
            if (isRowMerged.call(obj, rowId).length) {
                console.error('Jspreadsheet: This row is part of a merged cell.');
            } else {
                const target = e.target.clientHeight / 2 > e.offsetY ? e.target.parentNode!.nextSibling : e.target.parentNode;
                if (dragging.element != target) {
                    e.target.parentNode!.parentNode!.insertBefore(dragging.element, target);
                    dragging.destination = Array.prototype.indexOf.call(dragging.element.parentNode!.children, dragging.element);
                }
            }
        }
    }
};

/**
 * The cursor over the edges of the headers and the row numbers: col-resize or row-resize where a press starts a resize,
 * and move where it starts a move.
 */
const updateCursor = function (this: Worksheet, e: TargetEvent<MouseEvent>) {
    const obj = this;

    const x = e.target.getAttribute('data-x');
    const y = e.target.getAttribute('data-y');
    const rect = e.target.getBoundingClientRect();

    if (obj.pointer.cursor) {
        obj.pointer.cursor.style.cursor = '';
        obj.pointer.cursor = null;
    }

    if (e.target.parentNode!.parentNode && (e.target.parentNode!.parentNode as HTMLElement).className) {
        if ((e.target.parentNode!.parentNode as HTMLElement).classList.contains('resizable')) {
            if (e.target && x && !y && rect.width - (e.clientX - rect.left) < 6) {
                obj.pointer.cursor = e.target;
                obj.pointer.cursor.style.cursor = 'col-resize';
            } else if (e.target && !x && y && rect.height - (e.clientY - rect.top) < 6) {
                obj.pointer.cursor = e.target;
                obj.pointer.cursor.style.cursor = 'row-resize';
            }
        }

        if ((e.target.parentNode!.parentNode as HTMLElement).classList.contains('draggable')) {
            if (e.target && !x && y && rect.width - (e.clientX - rect.left) < 6) {
                obj.pointer.cursor = e.target;
                obj.pointer.cursor.style.cursor = 'move';
            } else if (e.target && x && !y && rect.height - (e.clientY - rect.top) < 6) {
                obj.pointer.cursor = e.target;
                obj.pointer.cursor.style.cursor = 'move';
            }
        }
    }
};

// Mouse move controls
export const mouseMoveControls = function (e: TargetEvent<MouseEvent>) {
    e = e || window.event;

    if (!getMouseButton(e)) {
        libraryBase.jspreadsheet.isMouseAction = false;
    }

    const current = libraryBase.jspreadsheet.current;

    if (current) {
        if (libraryBase.jspreadsheet.isMouseAction == true) {
            // Resizing is ongoing
            if (current.pointer.resizing) {
                moveResize.call(current, e);
            } else if (current.pointer.dragging) {
                moveDrag.call(current, e);
            }
        } else {
            updateCursor.call(current, e);
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

/**
 * The mouse over a header, a row number or a cell while the button is down after a press: extend the selection of
 * columns, rows or cells to it, or the area that the selection corner fills.
 */
const extendSelection = function (this: Worksheet, e: TargetEvent<MouseEvent>, section: number) {
    const obj = this;

    let columnId = getIndex(e.target, 'data-x');
    const rowId = getIndex(e.target, 'data-y');
    if (obj.pointer.resizing || obj.pointer.dragging) {
        // ignore
    } else {
        // Header found
        if (section == 1) {
            if (obj.selection.column != null) {
                columnId = getIndex(e.target, 'data-x');
                const o = obj.selection.column;
                const d = columnId;
                // Update selection
                updateSelectionFromCoords.call(obj, o, 0, d, obj.options.data.length - 1, e);
            }
        }

        // Body found
        if (section == 2) {
            if (e.target.classList.contains('jss_row')) {
                if (obj.selection.row != null) {
                    const o = obj.selection.row;
                    const d = rowId;
                    // Update selection
                    updateSelectionFromCoords.call(obj, 0, o, obj.options.data[0].length - 1, d, e);
                }
            } else {
                // Do not select edtion is in progress
                if (!obj.editor.isOpen()) {
                    if (columnId != null && rowId != null) {
                        if (obj.pointer.corner) {
                            updateCopySelection.call(obj, columnId, rowId);
                        } else {
                            if (obj.selection.range) {
                                updateSelectionFromCoords.call(obj, obj.selection.range[0], obj.selection.range[1], columnId, rowId, e);
                            }
                        }
                    }
                }
            }
        }
    }
};

export const mouseOverControls = function (e: TargetEvent<MouseEvent>) {
    e = e || window.event;

    if (!getMouseButton(e)) {
        libraryBase.jspreadsheet.isMouseAction = false;
    }

    const current = libraryBase.jspreadsheet.current;

    if (current && libraryBase.jspreadsheet.isMouseAction == true) {
        // Get elements
        const jssTable = getElement(e.target);

        if (jssTable[0]) {
            // Avoid cross reference
            if (current != jssTable[0].jssWorksheet) {
                return false;
            }

            extendSelection.call(current, e, jssTable[1]);
        }
    }

    // Clear any time control
    if (libraryBase.jspreadsheet.timeControl) {
        clearTimeout(libraryBase.jspreadsheet.timeControl);
        libraryBase.jspreadsheet.timeControl = null;
    }
};

/**
 * A double click on the selection corner fills the cells below the selection, down to the last row.
 */
const fillDown = function (this: Worksheet) {
    const obj = this;

    // Any selected cells
    if (obj.selection.highlighted.length > 0) {
        const last = obj.selection.highlighted[obj.selection.highlighted.length - 1].element;
        // Copy from this
        const x1 = parseInt(obj.selection.highlighted[0].element.getAttribute('data-x'));
        const y1 = parseInt(last.getAttribute('data-y')) + 1;
        // Until this
        const x2 = parseInt(last.getAttribute('data-x'));
        const y2 = obj.records.length - 1;
        // Execute copy
        copyData.call(obj, obj.records[y1][x1].element, obj.records[y2][x2].element);
    }
};

export const doubleClickControls = function (e: TargetEvent<MouseEvent>) {
    const current = libraryBase.jspreadsheet.current;

    // Jss is selected
    if (current) {
        // Corner action
        if (e.target.classList.contains('jss_corner')) {
            fillDown.call(current);
        } else if (e.target.classList.contains('jss_column_filter')) {
            // Column
            const columnId = parseInt(e.target.getAttribute('data-x'));
            // Open filter
            openFilter.call(current, columnId);
        } else {
            // Get table
            const jssTable = getElement(e.target);

            // Double click over header
            if (jssTable[1] == 1 && current.options.columnSorting != false) {
                // Check valid column header coords
                const columnId = getIndex(e.target, 'data-x');
                if (columnId != null) {
                    current.orderBy(columnId);
                }
            }

            // Double click over body
            if (jssTable[1] == 2 && current.options.editable != false) {
                if (!current.editor.isOpen()) {
                    const cell = getCell(e.target);
                    if (cell && cell.classList.contains('highlight')) {
                        openEditor.call(current, cell as HTMLTableCellElement, undefined, e);
                    }
                }
            }
        }
    }
};
