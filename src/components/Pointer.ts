import { getColumnNameFromId, getIdFromColumnName } from '../utils/internalHelpers.js';
import libraryBase from '../utils/libraryBase.js';
import { getIndex } from '../events/helpers.js';
import type { DragColumnInfo, DragRowInfo, ResizeColumnInfo, ResizeRowInfo, CellChange, CellValue } from '../jspreadsheet';
import type { Worksheet, TargetEvent } from '../types/internal';

function doubleDigitFormat(v: number | string) {
    v = `${v}`;
    if (v.length == 1) {
        v = `0${v}`;
    }
    return v;
}

/** The column and the row of the cell that contains the element, from its data-x and data-y attributes. */
function getCellPosition(element: HTMLElement): number[] | undefined {
    const x = element.getAttribute('data-x');
    const y = element.getAttribute('data-y');
    if (x && y) {
        return [parseInt(x), parseInt(y)];
    } else {
        if (element.parentNode) {
            return getCellPosition(element.parentNode as HTMLElement);
        }
    }
}

/**
 * The mouse gestures on a worksheet and their state: a press on a header, a row number, a cell or the pagination, the
 * resize or the move of a column or a row, the drag of the selection corner, and the element whose mouse cursor shows a
 * resize. It fills the cells when a drag of the corner ends.
 *
 * Each worksheet has one, as worksheet.pointer. The mouse handlers in src/events/mouse.ts find the worksheet and call it.
 */
export class Pointer {
    readonly worksheet: Worksheet;

    /**
     * The column or row that the mouse resizes, or null. Each member says that the other index is missing, so that a
     * test of column or row tells the type check which one it is.
     */
    resizing: null | (ResizeColumnInfo & { row?: undefined }) | (ResizeRowInfo & { column?: undefined }) = null;

    /** The column or row that the mouse moves, or null. */
    dragging: null | (DragColumnInfo & { row?: undefined }) | (DragRowInfo & { column?: undefined }) = null;

    /** The header or row element that shows a resize cursor, or null. */
    cursor: HTMLElement | null = null;

    /** True while the mouse drags the selection corner to fill cells. */
    corner = false;

    constructor(worksheet: Worksheet) {
        this.worksheet = worksheet;
    }

    /**
     * Helper function to copy data using the corner icon
     */
    fill(o: HTMLElement, d: HTMLElement) {
        const worksheet = this.worksheet;

        // Get data from all selected cells
        const data = worksheet.getData(true, false);

        // Selected cells. The corner is only there when cells are selected.
        const h = worksheet.selection.bounds!;

        // Cells
        const x1 = parseInt(o.getAttribute('data-x'));
        const y1 = parseInt(o.getAttribute('data-y'));
        const x2 = parseInt(d.getAttribute('data-x'));
        const y2 = parseInt(d.getAttribute('data-y'));

        // Records
        const records = [];
        let breakControl = false;

        let rowNumber, colNumber;

        if (h[0] == x1) {
            // Vertical copy
            if (y1 < h[1]) {
                rowNumber = y1 - h[1];
            } else {
                rowNumber = 1;
            }
            colNumber = 0;
        } else {
            if (x1 < h[0]) {
                colNumber = x1 - h[0];
            } else {
                colNumber = 1;
            }
            rowNumber = 0;
        }

        // Copy data procedure
        let posx = 0;
        let posy = 0;

        for (let j = y1; j <= y2; j++) {
            // Skip hidden rows
            if (worksheet.rows[j] && worksheet.rows[j].element.style.display == 'none') {
                continue;
            }

            // Controls
            if (data[posy] == undefined) {
                posy = 0;
            }
            posx = 0;

            // Data columns
            if (h[0] != x1) {
                if (x1 < h[0]) {
                    colNumber = x1 - h[0];
                } else {
                    colNumber = 1;
                }
            }
            // Data columns
            for (let i = x1; i <= x2; i++) {
                // Update non-readonly
                if (
                    worksheet.records[j][i] &&
                    !worksheet.records[j][i].element.classList.contains('readonly') &&
                    worksheet.records[j][i].element.style.display != 'none' &&
                    breakControl == false
                ) {
                    // Stop if contains value
                    if (!worksheet.selection.fill.length) {
                        if (worksheet.options.data[j][i] != '') {
                            breakControl = true;
                            continue;
                        }
                    }

                    // Column
                    if (data[posy] == undefined) {
                        posx = 0;
                    } else if (data[posy][posx] == undefined) {
                        posx = 0;
                    }

                    // Value
                    let value: CellValue = data[posy][posx];

                    if (value && !data[1] && worksheet.parent.config.autoIncrement != false) {
                        if (
                            worksheet.options.columns &&
                            worksheet.options.columns[i] &&
                            (!worksheet.options.columns[i].type ||
                                worksheet.options.columns[i].type == 'text' ||
                                (worksheet.options.columns[i].type as string) == 'number')
                        ) {
                            if (`${value}`.substr(0, 1) == '=') {
                                const tokens = (value as string).match(/([A-Z]+[0-9]+)/g);

                                if (tokens) {
                                    // An array used as a map, as before
                                    const affectedTokens = [] as unknown as Record<string, string>;
                                    for (const reference of tokens) {
                                        const position = getIdFromColumnName(reference, true);
                                        position[0] += colNumber;
                                        position[1] += rowNumber;
                                        if (position[1] < 0) {
                                            position[1] = 0;
                                        }
                                        const token = getColumnNameFromId([position[0], position[1]]);

                                        if (token != reference) {
                                            affectedTokens[reference] = token;
                                        }
                                    }
                                    // Update formula
                                    if (affectedTokens) {
                                        value = worksheet.formulas.updateFormula(value as string, affectedTokens);
                                    }
                                }
                            } else {
                                if (value == Number(value)) {
                                    value = Number(value) + rowNumber;
                                }
                            }
                        } else if (worksheet.options.columns && worksheet.options.columns[i] && worksheet.options.columns[i].type == 'calendar') {
                            // A calendar cell has a date as text
                            const date = new Date(value as string);
                            date.setDate(date.getDate() + rowNumber);
                            value = `${date.getFullYear()}-${doubleDigitFormat(parseInt(date.getMonth() + 1))}-${doubleDigitFormat(date.getDate())} 00:00:00`;
                        }
                    }

                    records.push(worksheet.cells.update(i, j, value));

                    // Update all formulas in the chain
                    worksheet.formulas.updateChain(i, j, records);
                }
                posx++;
                if (h[0] != x1) {
                    colNumber++;
                }
            }
            posy++;
            rowNumber++;
        }

        // Update history
        worksheet.history.add({
            action: 'setValue',
            records,
            selection: worksheet.selection.range && [...worksheet.selection.range],
        });

        // Update table with custom configuration if applicable
        worksheet.data.updateTable();

        // On after changes
        const onafterchangesRecords = records.map((record) => {
            return {
                x: record.x,
                y: record.y,
                value: record.value,
                oldValue: record.oldValue,
            };
        });

        // The coordinates can be text (issue #52), and a read-only cell has no value (issue #58)
        worksheet.dispatch('onafterchanges', worksheet, onafterchangesRecords as unknown as CellChange[]);
    }

    /**
     * The end of the resize of a column or a row: set the width or the height that the mouse gave, with the history.
     */
    endResize() {
        const worksheet = this.worksheet;
        const resizing = worksheet.pointer.resizing!;

        // Columns to be updated
        if (resizing.column != null) {
            // New width
            const newWidth = parseInt(worksheet.cols[resizing.column].colElement.getAttribute('width'));
            // Columns
            const columns = worksheet.getSelectedColumns();
            if (columns.length > 1) {
                const currentWidth = [];
                for (const columnIndex of columns) {
                    currentWidth.push(parseInt(worksheet.cols[columnIndex].colElement.getAttribute('width')));
                }
                // Previous width
                const index = columns.indexOf(parseInt(resizing.column));
                currentWidth[index] = resizing.width;
                worksheet.columns.setWidth(columns, newWidth, currentWidth);
            } else {
                worksheet.columns.setWidth(parseInt(resizing.column), newWidth, resizing.width);
            }
            // Remove border
            worksheet.headers[resizing.column].classList.remove('resizing');
            for (const record of worksheet.records) {
                if (record[resizing.column]) {
                    record[resizing.column].element.classList.remove('resizing');
                }
            }
        } else {
            // Remove Class
            worksheet.rows[resizing.row].element.children[0].classList.remove('resizing');
            // setHeight parses the height, and it ignores a missing attribute (NaN) as it ignored null
            const newHeight = parseInt(worksheet.rows[resizing.row].element.getAttribute('height'));
            worksheet.rowManager.setHeight(resizing.row, newHeight, resizing.height);
            // Remove border
            resizing.element.classList.remove('resizing');
        }
        // Reset resizing helper
        worksheet.pointer.resizing = null;
    }

    /**
     * The end of the move of a column or a row: move it to the position that the mouse gave.
     */
    endDrag(e: TargetEvent<MouseEvent>) {
        const worksheet = this.worksheet;
        const dragging = worksheet.pointer.dragging!;

        if (dragging.column != null) {
            // Target
            const columnId = getIndex(e.target, 'data-x');
            // Remove move style
            worksheet.headers[dragging.column].classList.remove('dragging');
            for (let j = 0; j < worksheet.rows.length; j++) {
                if (worksheet.records[j][dragging.column]) {
                    worksheet.records[j][dragging.column].element.classList.remove('dragging');
                }
            }
            for (const header of worksheet.headers) {
                header.classList.remove('dragging-left');
                header.classList.remove('dragging-right');
            }
            // Update position
            if (columnId != null) {
                if (dragging.column != dragging.destination) {
                    worksheet.moveColumn(dragging.column, dragging.destination);
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

                worksheet.rowManager.move(dragging.row, position, true);
            }
            dragging.element.classList.remove('dragging');
        }
        // Reset dragging helper
        worksheet.pointer.dragging = null;
    }

    /**
     * The end of a drag of the selection corner: fill the cells that the corner went over.
     */
    endFill() {
        const worksheet = this.worksheet;

        // Close any corner selection
        if (worksheet.pointer.corner) {
            worksheet.pointer.corner = false;

            // Data to be copied
            if (worksheet.selection.fill.length > 0) {
                // Copy data
                worksheet.pointer.fill(worksheet.selection.fill[0], worksheet.selection.fill[worksheet.selection.fill.length - 1]);

                // Remove selection
                worksheet.selection.clearFill();
            }
        }
    }

    /**
     * A press on a column header starts its resize near its right edge, its move near its bottom edge, and selects the
     * column otherwise. A press on a nested header selects its columns.
     */
    pressHeader(e: TargetEvent<MouseEvent>) {
        const worksheet = this.worksheet;

        const columnText = e.target.getAttribute('data-x');
        if (columnText) {
            const columnId = parseInt(columnText);
            // Update cursor
            const info = e.target.getBoundingClientRect();
            if (worksheet.options.columnResize != false && info.width - e.offsetX < 6) {
                // Resize helper
                worksheet.pointer.resizing = {
                    mousePosition: e.pageX,
                    column: columnId,
                    width: info.width,
                };

                // Border indication
                worksheet.headers[columnId].classList.add('resizing');
                for (const record of worksheet.records) {
                    if (record[columnId]) {
                        record[columnId].element.classList.add('resizing');
                    }
                }
            } else if (worksheet.options.columnDrag != false && info.height - e.offsetY < 6) {
                if (worksheet.merges.isColMerged(columnId).length) {
                    console.error('Jspreadsheet: This column is part of a merged cell.');
                } else {
                    // Reset selection
                    worksheet.resetSelection();
                    // Drag helper
                    worksheet.pointer.dragging = {
                        element: e.target as HTMLTableCellElement,
                        column: columnId,
                        destination: columnId,
                    };
                    // Border indication
                    worksheet.headers[columnId].classList.add('dragging');
                    for (const record of worksheet.records) {
                        if (record[columnId]) {
                            record[columnId].element.classList.add('dragging');
                        }
                    }
                }
            } else {
                let o, d;

                if (worksheet.selection.column != null && (e.shiftKey || e.ctrlKey)) {
                    o = worksheet.selection.column;
                    d = columnId;
                } else {
                    // Press to rename
                    if (worksheet.selection.column == columnId && worksheet.options.allowRenameColumn != false) {
                        // The active worksheet when the timer runs
                        libraryBase.jspreadsheet.timeControl = setTimeout(() => {
                            libraryBase.jspreadsheet.current!.setHeader(columnId);
                        }, 800);
                    }

                    // Keep track of which header was selected first
                    worksheet.selection.column = columnId;

                    // Update selection single column
                    o = columnId;
                    d = columnId;
                }

                // Update selection
                worksheet.selection.updateFromCoords(o, 0, d, worksheet.options.data.length - 1, e);
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
                    c2 = worksheet.options.columns.length - 1;
                }
                worksheet.selection.updateFromCoords(c1, 0, c2, worksheet.options.data.length - 1, e);
            }
        }
    }

    /**
     * A press on a row number starts the resize of the row near its bottom edge, its move near its right edge, and selects
     * the row otherwise.
     */
    pressRowNumber(e: TargetEvent<MouseEvent>) {
        const worksheet = this.worksheet;

        const rowId = parseInt(e.target.getAttribute('data-y'));
        const info = e.target.getBoundingClientRect();
        if (worksheet.options.rowResize != false && info.height - e.offsetY < 6) {
            // Resize helper
            worksheet.pointer.resizing = {
                element: e.target.parentNode as HTMLTableRowElement,
                mousePosition: e.pageY,
                row: rowId,
                height: info.height,
            };
            // Border indication
            (e.target.parentNode as HTMLElement).classList.add('resizing');
        } else if (worksheet.options.rowDrag != false && info.width - e.offsetX < 6) {
            if (worksheet.merges.isRowMerged(rowId).length) {
                console.error('Jspreadsheet: This row is part of a merged cell');
            } else if (worksheet.options.search == true && worksheet.view.results) {
                console.error('Jspreadsheet: Please clear your search before perform this action');
            } else {
                // Reset selection
                worksheet.resetSelection();
                // Drag helper
                worksheet.pointer.dragging = {
                    element: e.target.parentNode as HTMLTableRowElement,
                    row: rowId,
                    destination: rowId,
                };
                // Border indication
                (e.target.parentNode as HTMLElement).classList.add('dragging');
            }
        } else {
            let o, d;
            if (worksheet.selection.row != null && (e.shiftKey || e.ctrlKey)) {
                o = worksheet.selection.row;
                d = rowId;
            } else {
                // Keep track of which header was selected first
                worksheet.selection.row = rowId;

                // Update selection single column
                o = rowId;
                d = rowId;
            }

            // Update selection
            worksheet.selection.updateFromCoords(null, o, null, d, e);
        }
    }

    /**
     * A press on a cell closes the editor of another cell and selects the cell, or extends the selection to it with Shift.
     * A press on the close button of an editor closes the editor.
     */
    pressCell(e: TargetEvent<MouseEvent>) {
        const worksheet = this.worksheet;

        // Jclose
        if (e.target.classList.contains('jclose') && e.target.clientWidth - e.offsetX < 50 && e.offsetY < 50) {
            worksheet.editor.close(worksheet.editor.cell!, true);
        } else {
            const position = getCellPosition(e.target);
            if (position) {
                const columnId = position[0];
                const rowId = position[1];
                // Close edition
                if (worksheet.editor.isOpen()) {
                    if (worksheet.editor.x != columnId || worksheet.editor.y != rowId) {
                        worksheet.editor.close(worksheet.editor.cell, true);
                    }
                }

                if (!worksheet.editor.isOpen()) {
                    // Update cell selection
                    if (e.shiftKey) {
                        worksheet.selection.updateFromCoords(worksheet.selection.range![0], worksheet.selection.range![1], columnId, rowId, e);
                    } else {
                        worksheet.selection.updateFromCoords(columnId, rowId, undefined, undefined, e);
                    }
                }

                // No full row selected
                worksheet.selection.column = null;
                worksheet.selection.row = null;
            }
        }
    }

    /** A press on a page number of the pagination opens the page. */
    pressPage(e: TargetEvent<MouseEvent>) {
        const worksheet = this.worksheet;

        if (e.target.textContent == '<') {
            worksheet.page(0);
        } else if (e.target.textContent == '>') {
            worksheet.page(parseInt(e.target.getAttribute('title')) - 1);
        } else {
            worksheet.page(parseInt(e.target.textContent) - 1);
        }
    }

    /**
     * The resize of a column or a row while the mouse moves: set the width of the col or the height of the row.
     */
    moveResize(e: TargetEvent<MouseEvent>) {
        const worksheet = this.worksheet;
        const resizing = worksheet.pointer.resizing!;

        if (resizing.column != null) {
            const width = e.pageX - resizing.mousePosition;

            if (resizing.width + width > 0) {
                const tempWidth = resizing.width + width;
                worksheet.cols[resizing.column].colElement.setAttribute('width', tempWidth);

                worksheet.selection.updateCorner();
            }
        } else {
            const height = e.pageY - resizing.mousePosition;

            if (resizing.height + height > 0) {
                const tempHeight = resizing.height + height;
                worksheet.rows[resizing.row].element.setAttribute('height', tempHeight);

                worksheet.selection.updateCorner();
            }
        }
    }

    /**
     * The move of a column or a row while the mouse moves: mark the headers next to the destination of a column, or move
     * the row in the table.
     */
    moveDrag(e: TargetEvent<MouseEvent>) {
        const worksheet = this.worksheet;
        const dragging = worksheet.pointer.dragging!;

        if (dragging.column != null) {
            const columnId = getIndex(e.target, 'data-x');
            if (columnId != null) {
                if (worksheet.merges.isColMerged(columnId).length) {
                    console.error('Jspreadsheet: This column is part of a merged cell.');
                } else {
                    for (const header of worksheet.headers) {
                        header.classList.remove('dragging-left');
                        header.classList.remove('dragging-right');
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
                            worksheet.headers[columnId].classList.add('dragging-left');
                        } else {
                            if (dragging.column < columnId) {
                                dragging.destination = columnId;
                            } else {
                                dragging.destination = columnId + 1;
                            }
                            worksheet.headers[columnId].classList.add('dragging-right');
                        }
                    }
                }
            }
        } else {
            const rowId = getIndex(e.target, 'data-y');
            if (rowId != null) {
                if (worksheet.merges.isRowMerged(rowId).length) {
                    console.error('Jspreadsheet: This row is part of a merged cell.');
                } else {
                    const target = e.target.clientHeight / 2 > e.offsetY ? e.target.parentNode!.nextSibling : e.target.parentNode;
                    if (dragging.element != target) {
                        e.target.parentNode!.parentNode!.insertBefore(dragging.element, target);
                        dragging.destination = Array.from(dragging.element.parentNode!.children).indexOf(dragging.element);
                    }
                }
            }
        }
    }

    /**
     * The cursor over the edges of the headers and the row numbers: col-resize or row-resize where a press starts a resize,
     * and move where it starts a move.
     */
    updateCursor(e: TargetEvent<MouseEvent>) {
        const worksheet = this.worksheet;

        const x = e.target.getAttribute('data-x');
        const y = e.target.getAttribute('data-y');
        const rect = e.target.getBoundingClientRect();

        if (worksheet.pointer.cursor) {
            worksheet.pointer.cursor.style.cursor = '';
            worksheet.pointer.cursor = null;
        }

        if (e.target.parentNode!.parentNode && (e.target.parentNode!.parentNode as HTMLElement).className) {
            if ((e.target.parentNode!.parentNode as HTMLElement).classList.contains('resizable')) {
                if (e.target && x && !y && rect.width - (e.clientX - rect.left) < 6) {
                    worksheet.pointer.cursor = e.target;
                    worksheet.pointer.cursor.style.cursor = 'col-resize';
                } else if (e.target && !x && y && rect.height - (e.clientY - rect.top) < 6) {
                    worksheet.pointer.cursor = e.target;
                    worksheet.pointer.cursor.style.cursor = 'row-resize';
                }
            }

            if ((e.target.parentNode!.parentNode as HTMLElement).classList.contains('draggable')) {
                if (e.target && !x && y && rect.width - (e.clientX - rect.left) < 6) {
                    worksheet.pointer.cursor = e.target;
                    worksheet.pointer.cursor.style.cursor = 'move';
                } else if (e.target && x && !y && rect.height - (e.clientY - rect.top) < 6) {
                    worksheet.pointer.cursor = e.target;
                    worksheet.pointer.cursor.style.cursor = 'move';
                }
            }
        }
    }

    /**
     * Update copy selection
     *
     * @param int x, y
     * @return void
     */
    private updateCopySelection(x3: number, y3: number) {
        const worksheet = this.worksheet;

        // Remove selection
        worksheet.selection.clearFill();

        // Get elements first and last
        const x1 = worksheet.selection.bounds![0];
        const y1 = worksheet.selection.bounds![1];
        const x2 = worksheet.selection.bounds![2];
        const y2 = worksheet.selection.bounds![3];

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
                    if (
                        worksheet.records[j][i] &&
                        worksheet.rows[j].element.style.display != 'none' &&
                        worksheet.records[j][i].element.style.display != 'none'
                    ) {
                        worksheet.records[j][i].element.classList.add('selection');
                        worksheet.records[py][i].element.classList.add('selection-top');
                        worksheet.records[uy][i].element.classList.add('selection-bottom');
                        worksheet.records[j][px].element.classList.add('selection-left');
                        worksheet.records[j][ux].element.classList.add('selection-right');

                        // Persist selected elements
                        worksheet.selection.fill.push(worksheet.records[j][i].element);
                    }
                }
            }
        }
    }

    /**
     * The mouse over a header, a row number or a cell while the button is down after a press: extend the selection of
     * columns, rows or cells to it, or the area that the selection corner fills.
     */
    extendSelection(e: TargetEvent<MouseEvent>, section: number) {
        const worksheet = this.worksheet;

        let columnId = getIndex(e.target, 'data-x');
        const rowId = getIndex(e.target, 'data-y');
        if (worksheet.pointer.resizing || worksheet.pointer.dragging) {
            // ignore
        } else {
            // Header found
            if (section == 1) {
                if (worksheet.selection.column != null) {
                    columnId = getIndex(e.target, 'data-x');
                    const o = worksheet.selection.column;
                    const d = columnId;
                    // Update selection
                    worksheet.selection.updateFromCoords(o, 0, d, worksheet.options.data.length - 1, e);
                }
            }

            // Body found
            if (section == 2) {
                if (e.target.classList.contains('jss_row')) {
                    if (worksheet.selection.row != null) {
                        const o = worksheet.selection.row;
                        const d = rowId;
                        // Update selection
                        worksheet.selection.updateFromCoords(0, o, worksheet.options.data[0].length - 1, d, e);
                    }
                } else {
                    // Do not select edtion is in progress
                    if (!worksheet.editor.isOpen()) {
                        if (columnId != null && rowId != null) {
                            if (worksheet.pointer.corner) {
                                this.updateCopySelection(columnId, rowId);
                            } else {
                                if (worksheet.selection.range) {
                                    worksheet.selection.updateFromCoords(worksheet.selection.range[0], worksheet.selection.range[1], columnId, rowId, e);
                                }
                            }
                        }
                    }
                }
            }
        }
    }

    /**
     * A double click on the selection corner fills the cells below the selection, down to the last row.
     */
    fillDown() {
        const worksheet = this.worksheet;

        // Any selected cells
        if (worksheet.selection.highlighted.length > 0) {
            const last = worksheet.selection.highlighted[worksheet.selection.highlighted.length - 1].element;
            // Copy from this
            const x1 = parseInt(worksheet.selection.highlighted[0].element.getAttribute('data-x'));
            const y1 = parseInt(last.getAttribute('data-y')) + 1;
            // Until this
            const x2 = parseInt(last.getAttribute('data-x'));
            const y2 = worksheet.records.length - 1;
            // Execute copy
            worksheet.pointer.fill(worksheet.records[y1][x1].element, worksheet.records[y2][x2].element);
        }
    }
}
