import dispatch from './dispatch.js';
import { getFreezeWidth } from './freeze.js';
import { getCellNameFromCoords } from './helpers.js';
import { updateCell, updateFormula, updateFormulaChain, updateTable } from './internal.js';
import { getColumnNameFromId, getIdFromColumnName } from './internalHelpers.js';
import { updateToolbar } from './toolbar.js';
import { updateActiveCell } from './accessibility.js';
import type { Worksheet, WorksheetMethod } from '../types/internal';
import type { CellChange, CellValue } from '../jspreadsheet';

/**
 * Get the number of columns of the merged range that starts at x, y.
 * The colspan attribute counts only the columns that are not hidden.
 */
const getMergedColspan = function (this: Worksheet, element: HTMLElement, x: number, y: number) {
    const obj = this;

    const merged = obj.options.mergeCells && obj.options.mergeCells[getColumnNameFromId([x, y])];

    return merged ? parseInt(merged[0]) : parseInt(element.getAttribute('colspan'));
};

export const updateCornerPosition = function (this: Worksheet) {
    const obj = this;

    // If any selected cells
    if (!obj.highlighted || !obj.highlighted.length) {
        obj.corner.style.top = '-2000px';
        obj.corner.style.left = '-2000px';
    } else {
        // Get last cell
        const last = obj.highlighted[obj.highlighted.length - 1].element;
        const lastX = parseInt(last.getAttribute('data-x'));

        const contentRect = obj.content.getBoundingClientRect();
        const x1 = contentRect.left;
        const y1 = contentRect.top;

        const lastRect = last.getBoundingClientRect();
        const x2 = lastRect.left;
        const y2 = lastRect.top;
        const w2 = lastRect.width;
        const h2 = lastRect.height;

        const x = x2 - x1 + obj.content.scrollLeft + w2 - 4;
        const y = y2 - y1 + obj.content.scrollTop + h2 - 4;

        // Place the corner in the correct place
        obj.corner.style.top = y + 'px';
        obj.corner.style.left = x + 'px';

        if (obj.options.freezeColumns) {
            const width = getFreezeWidth.call(obj);
            // Only check if the last column is not part of the merged cells
            if (lastX > obj.options.freezeColumns - 1 && x2 - x1 + w2 < width) {
                obj.corner.style.display = 'none';
            } else {
                if (obj.options.selectionCopy != false) {
                    obj.corner.style.display = '';
                }
            }
        } else {
            if (obj.options.selectionCopy != false) {
                obj.corner.style.display = '';
            }
        }
    }

    updateToolbar(obj);
};

export const resetSelection = function (this: Worksheet, blur?: boolean) {
    const obj = this;

    let previousStatus;

    // Remove style
    if (!obj.highlighted || !obj.highlighted.length) {
        previousStatus = 0;
    } else {
        previousStatus = 1;

        for (let i = 0; i < obj.highlighted.length; i++) {
            obj.highlighted[i].element.classList.remove('highlight');
            obj.highlighted[i].element.classList.remove('highlight-left');
            obj.highlighted[i].element.classList.remove('highlight-right');
            obj.highlighted[i].element.classList.remove('highlight-top');
            obj.highlighted[i].element.classList.remove('highlight-bottom');
            obj.highlighted[i].element.classList.remove('highlight-selected');
            obj.highlighted[i].element.removeAttribute('aria-selected');

            const px = parseInt(obj.highlighted[i].element.getAttribute('data-x'));
            const py = parseInt(obj.highlighted[i].element.getAttribute('data-y'));

            // Check for merged cells
            let ux, uy;

            if (obj.highlighted[i].element.getAttribute('data-merged')) {
                const colspan = getMergedColspan.call(obj, obj.highlighted[i].element, px, py);
                const rowspan = parseInt(obj.highlighted[i].element.getAttribute('rowspan'));
                ux = colspan > 0 ? px + (colspan - 1) : px;
                uy = rowspan > 0 ? py + (rowspan - 1) : py;
            } else {
                ux = px;
                uy = py;
            }

            // Remove selected from headers
            for (let j = px; j <= ux; j++) {
                if (obj.headers[j]) {
                    obj.headers[j].classList.remove('selected');
                }
            }

            // Remove selected from rows
            for (let j = py; j <= uy; j++) {
                if (obj.rows[j]) {
                    obj.rows[j].element.classList.remove('selected');
                }
            }
        }
    }

    // Reset highlighted cells
    obj.highlighted = [];

    // The tab stop of the grid stays on the active cell, which is not selected now
    if (obj.activeCell) {
        obj.activeCell.setAttribute('aria-selected', 'false');
    }

    // Reset
    obj.selectedCell = null;

    // Hide corner
    obj.corner.style.top = '-2000px';
    obj.corner.style.left = '-2000px';

    if (blur == true && previousStatus == 1) {
        dispatch.call(obj, 'onblur', obj);
    }

    return previousStatus;
};

/**
 * Update selection based on two cells
 */
export const updateSelection = function (this: Worksheet, el1: HTMLElement, el2?: HTMLElement, origin?: Event) {
    const obj = this;

    // Both elements are cells, which have data-x and data-y
    const x1 = parseInt(el1.getAttribute('data-x'));
    const y1 = parseInt(el1.getAttribute('data-y'));

    let x2, y2;
    if (el2) {
        x2 = parseInt(el2.getAttribute('data-x'));
        y2 = parseInt(el2.getAttribute('data-y'));
    } else {
        x2 = x1;
        y2 = y1;
    }

    updateSelectionFromCoords.call(obj, x1, y1, x2, y2, origin);
};

export const removeCopyingSelection = function () {
    const copying = document.querySelectorAll('.jss_worksheet .copying');
    for (let i = 0; i < copying.length; i++) {
        copying[i].classList.remove('copying');
        copying[i].classList.remove('copying-left');
        copying[i].classList.remove('copying-right');
        copying[i].classList.remove('copying-top');
        copying[i].classList.remove('copying-bottom');
    }
};

// A null x1 selects the rows from y1 to y2, and a null y1 selects the columns from x1 to x2
export const updateSelectionFromCoords = function (
    this: Worksheet,
    x1: number | null,
    y1: number | null,
    x2?: number | null,
    y2?: number | null,
    origin?: Event
) {
    const obj = this;

    // select column
    if (y1 == null) {
        y1 = 0;
        y2 = obj.rows.length - 1;

        if (x1 == null) {
            return;
        }
    } else if (x1 == null) {
        // select row
        x1 = 0;
        x2 = obj.options.data[0].length - 1;
    }

    // Same element
    if (x2 == null) {
        x2 = x1;
    }
    if (y2 == null) {
        y2 = y1;
    }

    // The mouse events pass the text of data-x and data-y. Keep the selection as numbers.
    x1 = parseInt(x1);
    y1 = parseInt(y1);
    x2 = parseInt(x2);
    y2 = parseInt(y2);

    // Selection must be within the existing data
    if (x1 >= obj.headers.length) {
        x1 = obj.headers.length - 1;
    }
    if (y1 >= obj.rows.length) {
        y1 = obj.rows.length - 1;
    }
    if (x2 >= obj.headers.length) {
        x2 = obj.headers.length - 1;
    }
    if (y2 >= obj.rows.length) {
        y2 = obj.rows.length - 1;
    }

    // Limits
    let borderLeft = null;
    let borderRight = null;
    let borderTop = null;
    let borderBottom = null;

    // Origin & Destination
    let px: number, ux: number;

    if (parseInt(x1) < parseInt(x2)) {
        px = parseInt(x1);
        ux = parseInt(x2);
    } else {
        px = parseInt(x2);
        ux = parseInt(x1);
    }

    let py: number, uy: number;

    if (parseInt(y1) < parseInt(y2)) {
        py = parseInt(y1);
        uy = parseInt(y2);
    } else {
        py = parseInt(y2);
        uy = parseInt(y1);
    }

    // Verify merged columns
    for (let i = px; i <= ux; i++) {
        for (let j = py; j <= uy; j++) {
            if (obj.records[j][i] && obj.records[j][i].element.getAttribute('data-merged')) {
                const x = parseInt(obj.records[j][i].element.getAttribute('data-x'));
                const y = parseInt(obj.records[j][i].element.getAttribute('data-y'));
                const colspan = getMergedColspan.call(obj, obj.records[j][i].element, x, y);
                const rowspan = parseInt(obj.records[j][i].element.getAttribute('rowspan'));

                if (colspan > 1) {
                    if (x < px) {
                        px = x;
                    }
                    if (x + colspan > ux) {
                        ux = x + colspan - 1;
                    }
                }

                if (rowspan) {
                    if (y < py) {
                        py = y;
                    }
                    if (y + rowspan > uy) {
                        uy = y + rowspan - 1;
                    }
                }
            }
        }
    }

    // Vertical limits
    for (let j = py; j <= uy; j++) {
        if (obj.rows[j].element.style.display != 'none') {
            if (borderTop == null) {
                borderTop = j;
            }
            borderBottom = j;
        }
    }

    for (let i = px; i <= ux; i++) {
        for (let j = py; j <= uy; j++) {
            // Horizontal limits
            if (!obj.options.columns || !obj.options.columns[i] || obj.options.columns[i].type != 'hidden') {
                if (borderLeft == null) {
                    borderLeft = i;
                }
                borderRight = i;
            }
        }
    }

    // Create borders
    if (!borderLeft) {
        borderLeft = 0;
    }
    if (!borderRight) {
        borderRight = 0;
    }

    // The top and bottom are null when every selected row is hidden (issue #59)
    const ret = dispatch.call(obj, 'onbeforeselection', obj, borderLeft, borderTop!, borderRight, borderBottom!, origin);
    if (ret === false) {
        return false;
    }

    // Reset Selection
    const previousState = obj.resetSelection();

    // Keep selected cell
    obj.selectedCell = [x1, y1, x2, y2];

    // Add selected cell
    if (obj.records[y1][x1]) {
        obj.records[y1][x1].element.classList.add('highlight-selected');
    }

    // Redefining styles
    for (let i = px; i <= ux; i++) {
        for (let j = py; j <= uy; j++) {
            if (obj.rows[j].element.style.display != 'none' && obj.records[j][i].element.style.display != 'none') {
                obj.records[j][i].element.classList.add('highlight');
                obj.records[j][i].element.setAttribute('aria-selected', 'true');
                obj.highlighted.push(obj.records[j][i]);
            }
        }
    }

    for (let i = borderLeft; i <= borderRight; i++) {
        if (
            (!obj.options.columns || !obj.options.columns[i] || obj.options.columns[i].type != 'hidden') &&
            obj.cols[i].colElement.style &&
            obj.cols[i].colElement.style.display != 'none'
        ) {
            // Top border. The borders are null when each row of the range is hidden.
            if (obj.records[borderTop!] && obj.records[borderTop!][i]) {
                obj.records[borderTop!][i].element.classList.add('highlight-top');
            }
            // Bottom border
            if (obj.records[borderBottom!] && obj.records[borderBottom!][i]) {
                obj.records[borderBottom!][i].element.classList.add('highlight-bottom');
            }
            // Add selected from headers
            obj.headers[i].classList.add('selected');
        }
    }

    for (let j = borderTop!; j <= borderBottom!; j++) {
        if (obj.rows[j] && obj.rows[j].element.style.display != 'none') {
            // Left border
            obj.records[j][borderLeft].element.classList.add('highlight-left');
            // Right border
            obj.records[j][borderRight].element.classList.add('highlight-right');
            // Add selected from rows
            obj.rows[j].element.classList.add('selected');
        }
    }

    obj.selectedContainer = [borderLeft, borderTop!, borderRight, borderBottom!];

    // The selected cell is the tab stop of the grid
    updateActiveCell.call(obj, x1, y1, origin);

    // Handle events
    if (previousState == 0) {
        dispatch.call(obj, 'onfocus', obj);

        removeCopyingSelection();
    }

    // The top and bottom are null when every selected row is hidden (issue #59)
    dispatch.call(obj, 'onselection', obj, borderLeft, borderTop!, borderRight, borderBottom!, origin);

    // Find corner cell
    updateCornerPosition.call(obj);
};

/**
 * Get selected column numbers
 *
 * @return array
 */
export const getSelectedColumns: WorksheetMethod<'getSelectedColumns'> = function (visibleOnly) {
    const obj = this;

    if (!obj.selectedCell) {
        return [];
    }

    const result = [];

    for (let i = Math.min(obj.selectedCell[0], obj.selectedCell[2]); i <= Math.max(obj.selectedCell[0], obj.selectedCell[2]); i++) {
        if (!visibleOnly || obj.headers[i].style.display != 'none') {
            result.push(i);
        }
    }

    return result;
};

/**
 * Refresh current selection
 */
export const refreshSelection = function (this: Worksheet) {
    const obj = this;

    if (obj.selectedCell) {
        obj.updateSelectionFromCoords(obj.selectedCell[0], obj.selectedCell[1], obj.selectedCell[2], obj.selectedCell[3]);
    }
};

/**
 * Remove copy selection
 *
 * @return void
 */
export const removeCopySelection = function (this: Worksheet) {
    const obj = this;

    // Remove current selection
    for (let i = 0; i < obj.selection.length; i++) {
        obj.selection[i].classList.remove('selection');
        obj.selection[i].classList.remove('selection-left');
        obj.selection[i].classList.remove('selection-right');
        obj.selection[i].classList.remove('selection-top');
        obj.selection[i].classList.remove('selection-bottom');
    }

    obj.selection = [];
};

const doubleDigitFormat = function (v: number | string) {
    v = '' + v;
    if (v.length == 1) {
        v = '0' + v;
    }
    return v;
};

/**
 * Helper function to copy data using the corner icon
 */
export const copyData = function (this: Worksheet, o: HTMLElement, d: HTMLElement) {
    const obj = this;

    // Get data from all selected cells
    const data = obj.getData(true, false);

    // Selected cells. The corner is only there when cells are selected.
    const h = obj.selectedContainer!;

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
        if (obj.rows[j] && obj.rows[j].element.style.display == 'none') {
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
                obj.records[j][i] &&
                !obj.records[j][i].element.classList.contains('readonly') &&
                obj.records[j][i].element.style.display != 'none' &&
                breakControl == false
            ) {
                // Stop if contains value
                if (!obj.selection.length) {
                    if (obj.options.data[j][i] != '') {
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

                if (value && !data[1] && obj.parent.config.autoIncrement != false) {
                    if (
                        obj.options.columns &&
                        obj.options.columns[i] &&
                        (!obj.options.columns[i].type || obj.options.columns[i].type == 'text' || (obj.options.columns[i].type as string) == 'number')
                    ) {
                        if (('' + value).substr(0, 1) == '=') {
                            const tokens = (value as string).match(/([A-Z]+[0-9]+)/g);

                            if (tokens) {
                                // An array used as a map, as before
                                const affectedTokens = [] as unknown as Record<string, string>;
                                for (let index = 0; index < tokens.length; index++) {
                                    const position = getIdFromColumnName(tokens[index], true);
                                    position[0] += colNumber;
                                    position[1] += rowNumber;
                                    if (position[1] < 0) {
                                        position[1] = 0;
                                    }
                                    const token = getColumnNameFromId([position[0], position[1]]);

                                    if (token != tokens[index]) {
                                        affectedTokens[tokens[index]] = token;
                                    }
                                }
                                // Update formula
                                if (affectedTokens) {
                                    value = updateFormula(value as string, affectedTokens);
                                }
                            }
                        } else {
                            if (value == Number(value)) {
                                value = Number(value) + rowNumber;
                            }
                        }
                    } else if (obj.options.columns && obj.options.columns[i] && obj.options.columns[i].type == 'calendar') {
                        // A calendar cell has a date as text
                        const date = new Date(value as string);
                        date.setDate(date.getDate() + rowNumber);
                        value =
                            date.getFullYear() +
                            '-' +
                            doubleDigitFormat(parseInt(date.getMonth() + 1)) +
                            '-' +
                            doubleDigitFormat(date.getDate()) +
                            ' ' +
                            '00:00:00';
                    }
                }

                records.push(updateCell.call(obj, i, j, value));

                // Update all formulas in the chain
                updateFormulaChain.call(obj, i, j, records);
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
    obj.history.add({
        action: 'setValue',
        records: records,
        selection: obj.selectedCell,
    });

    // Update table with custom configuration if applicable
    updateTable.call(obj);

    // On after changes
    const onafterchangesRecords = records.map(function (record) {
        return {
            x: record.x,
            y: record.y,
            value: record.value,
            oldValue: record.oldValue,
        };
    });

    // The coordinates can be text (issue #52), and a read-only cell has no value (issue #58)
    dispatch.call(obj, 'onafterchanges', obj, onafterchangesRecords as unknown as CellChange[]);
};

export const hash = function (str: string) {
    let hash = 0,
        i,
        chr;

    if (!str || str.length === 0) {
        return hash;
    } else {
        for (i = 0; i < str.length; i++) {
            chr = str.charCodeAt(i);
            hash = (hash << 5) - hash + chr;
            hash |= 0;
        }
    }
    return hash;
};

/**
 * Move coords to A1 in case overlaps with an excluded cell
 */
// type is 1 for rows and 0 for columns
export const conditionalSelectionUpdate = function (this: Worksheet, type: 0 | 1, o: number, d: number) {
    const obj = this;

    if (type == 1) {
        if (obj.selectedCell && ((o >= obj.selectedCell[1] && o <= obj.selectedCell[3]) || (d >= obj.selectedCell[1] && d <= obj.selectedCell[3]))) {
            obj.resetSelection();
            return;
        }
    } else {
        if (obj.selectedCell && ((o >= obj.selectedCell[0] && o <= obj.selectedCell[2]) || (d >= obj.selectedCell[0] && d <= obj.selectedCell[2]))) {
            obj.resetSelection();
            return;
        }
    }
};

/**
 * Get selected rows numbers
 *
 * @return array
 */
export const getSelectedRows: WorksheetMethod<'getSelectedRows'> = function (visibleOnly) {
    const obj = this;

    if (!obj.selectedCell) {
        return [];
    }

    const result = [];

    for (let i = Math.min(obj.selectedCell[1], obj.selectedCell[3]); i <= Math.max(obj.selectedCell[1], obj.selectedCell[3]); i++) {
        if (!visibleOnly || obj.rows[i].element.style.display != 'none') {
            result.push(i);
        }
    }

    return result;
};

export const selectAll: WorksheetMethod<'selectAll'> = function () {
    const obj = this;

    if (!obj.selectedCell) {
        obj.selectedCell = [] as unknown as [number, number, number, number];
    }

    obj.selectedCell[0] = 0;
    obj.selectedCell[1] = 0;
    obj.selectedCell[2] = obj.headers.length - 1;
    obj.selectedCell[3] = obj.records.length - 1;

    obj.updateSelectionFromCoords(obj.selectedCell[0], obj.selectedCell[1], obj.selectedCell[2], obj.selectedCell[3]);
};

export const getSelection: WorksheetMethod<'getSelection'> = function () {
    const obj = this;

    if (!obj.selectedCell) {
        return null;
    }

    return [
        Math.min(obj.selectedCell[0], obj.selectedCell[2]),
        Math.min(obj.selectedCell[1], obj.selectedCell[3]),
        Math.max(obj.selectedCell[0], obj.selectedCell[2]),
        Math.max(obj.selectedCell[1], obj.selectedCell[3]),
    ];
};

// The names of the cells, or their records. The public typings have an overload for each.
export const getSelected = function (this: Worksheet, columnNameOnly?: boolean): (string | Worksheet['records'][number][number])[] {
    const obj = this;

    const selectedRange = getSelection.call(obj);

    if (!selectedRange) {
        return [];
    }

    const cells = [];

    for (let y = selectedRange[1]; y <= selectedRange[3]; y++) {
        for (let x = selectedRange[0]; x <= selectedRange[2]; x++) {
            if (columnNameOnly) {
                cells.push(getCellNameFromCoords(x, y));
            } else {
                cells.push(obj.records[y][x]);
            }
        }
    }

    return cells;
};

export const getRange: WorksheetMethod<'getRange'> = function () {
    const obj = this;

    const selectedRange = getSelection.call(obj);

    if (!selectedRange) {
        return '';
    }

    const start = getCellNameFromCoords(selectedRange[0], selectedRange[1]);
    const end = getCellNameFromCoords(selectedRange[2], selectedRange[3]);

    if (start === end) {
        return obj.options.worksheetName + '!' + start;
    }

    return obj.options.worksheetName + '!' + start + ':' + end;
};

export const isSelected: WorksheetMethod<'isSelected'> = function (x, y) {
    const obj = this;

    const selection = getSelection.call(obj);

    // No cell is selected
    if (!selection) {
        return false;
    }

    return x >= selection[0] && x <= selection[2] && y >= selection[1] && y <= selection[3];
};

export const getHighlighted: WorksheetMethod<'getHighlighted'> = function () {
    const obj = this;

    const selection = getSelection.call(obj);

    if (selection) {
        return [selection];
    }

    return [];
};
