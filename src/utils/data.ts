import { createRow } from './rows.js';
import { updateCell } from './render.js';
import { updateFormulaChain } from './references.js';
import { updateTable } from './table.js';
import { getIdFromColumnName } from './internalHelpers.js';
import dispatch from './dispatch.js';
import { updatePagination } from './pagination.js';
import { setMerge } from './merges.js';
import { getCoordsFromRange } from './helpers.js';
import { updateGridSize, updateTabStop } from './accessibility.js';
import type { Worksheet, WorksheetMethod } from '../types/internal';
import type { CellChange, CellValue } from '../jspreadsheet';

// data can be rows of objects. They are changed to arrays.
export const setData: WorksheetMethod<'setData'> = function (data) {
    const obj = this;

    // Update data. Rows of objects are changed to arrays below.
    if (data) {
        obj.options.data = data as CellValue[][];
    }

    // Data
    if (!obj.options.data) {
        obj.options.data = [];
    }

    // Prepare data
    if (obj.options.data && obj.options.data[0]) {
        if (!Array.isArray(obj.options.data[0])) {
            const rows: CellValue[][] = [];
            for (let j = 0; j < obj.options.data.length; j++) {
                const row: CellValue[] = [];
                for (let i = 0; i < obj.options.columns.length; i++) {
                    // The rows are objects here
                    row[i] = (obj.options.data[j] as unknown as Record<string, CellValue>)[obj.options.columns[i].name!];
                }
                rows.push(row);
            }

            obj.options.data = rows;
        }
    }

    // Adjust minimal dimensions
    let j = 0;
    let i = 0;
    const size_i = (obj.options.columns && obj.options.columns.length) || 0;
    const size_j = obj.options.data.length;
    const min_i = obj.options.minDimensions[0];
    const min_j = obj.options.minDimensions[1];
    const max_i = min_i > size_i ? min_i : size_i;
    const max_j = min_j > size_j ? min_j : size_j;

    for (j = 0; j < max_j; j++) {
        for (i = 0; i < max_i; i++) {
            if (obj.options.data[j] == undefined) {
                obj.options.data[j] = [];
            }

            if (obj.options.data[j][i] == undefined) {
                obj.options.data[j][i] = '';
            }
        }
    }

    // Reset containers
    obj.rows = [];
    obj.view.results = null;
    obj.records = [];
    obj.history.clear();

    // Reset data
    obj.tbody.innerHTML = '';

    let startNumber;
    let finalNumber;

    // Lazy loading
    if (obj.options.lazyLoading == true) {
        // Load only 100 records
        startNumber = 0;
        finalNumber = obj.options.data.length < 100 ? obj.options.data.length : 100;

        if (obj.options.pagination) {
            // pagination is declared as a number, and false turns it off as 0 does
            (obj.options as { pagination?: number | false }).pagination = false;
            console.error('Jspreadsheet: Pagination will be disable due the lazyLoading');
        }
    } else if (obj.options.pagination) {
        // Pagination
        if (!obj.view.pageNumber) {
            obj.view.pageNumber = 0;
        }
        startNumber = obj.options.pagination * obj.view.pageNumber;
        finalNumber = obj.options.pagination * obj.view.pageNumber + obj.options.pagination;

        if (obj.options.data.length < finalNumber) {
            finalNumber = obj.options.data.length;
        }
    } else {
        startNumber = 0;
        finalNumber = obj.options.data.length;
    }

    // Append nodes to the HTML
    for (j = 0; j < obj.options.data.length; j++) {
        // Create row
        const row = createRow.call(obj, j, obj.options.data[j]);
        // Append line to the table
        if (j >= startNumber && j < finalNumber) {
            obj.tbody.appendChild(row.element);
        }
    }

    if (obj.options.lazyLoading == true) {
        // Do not create pagination with lazyloading activated
    } else if (obj.options.pagination) {
        updatePagination.call(obj);
    }

    // Merge cells
    if (obj.options.mergeCells) {
        const keys = Object.keys(obj.options.mergeCells);
        for (let i = 0; i < keys.length; i++) {
            const num = obj.options.mergeCells[keys[i]];
            setMerge.call(obj, keys[i], num[0], num[1], 1);
        }
    }

    // Update the size of the grid
    updateGridSize.call(obj);

    // The rows are new, so the tab stop of the grid is new
    updateTabStop.call(obj);

    // Updata table with custom configurations if applicable
    updateTable.call(obj);
};

/**
 * Get the value from a cell
 *
 * @param object cell
 * @return string value
 */
export const getValue: WorksheetMethod<'getValue'> = function (cell, processedValue) {
    const obj = this;

    if (typeof cell !== 'string') {
        return null;
    }

    const coords = getIdFromColumnName(cell, true);
    const x = coords[0];
    const y = coords[1];

    let value = null;

    if (x != null && y != null) {
        if (obj.records[y] && obj.records[y][x] && processedValue) {
            value = obj.records[y][x].element.innerHTML;
        } else {
            if (obj.options.data[y] && obj.options.data[y][x] != 'undefined') {
                value = obj.options.data[y][x];
            }
        }
    }

    return value;
};

/**
 * Get the value from a coords
 *
 * @param int x
 * @param int y
 * @return string value
 */
export const getValueFromCoords: WorksheetMethod<'getValueFromCoords'> = function (x, y, processedValue) {
    const obj = this;

    let value = null;

    if (x != null && y != null) {
        if (obj.records[y] && obj.records[y][x] && processedValue) {
            value = obj.records[y][x].element.innerHTML;
        } else {
            if (obj.options.data[y] && obj.options.data[y][x] != 'undefined') {
                value = obj.options.data[y][x];
            }
        }
    }

    return value;
};

/**
 * Set a cell value
 *
 * @param mixed cell destination cell
 * @param string value value
 * @return void
 */
// any: a cell element gives the text of data-x and data-y to updateCell (issue #52)
export const setValue: WorksheetMethod<'setValue'> = function (cell: any, value, force) {
    const obj = this;

    const records = [];

    if (typeof cell == 'string') {
        const columnId = getIdFromColumnName(cell, true);
        const x = columnId[0];
        const y = columnId[1];

        // Update cell
        records.push(updateCell.call(obj, x, y, value, force));

        // Update all formulas in the chain
        updateFormulaChain.call(obj, x, y, records);
    } else {
        let x = null;
        let y = null;
        if (cell && cell.getAttribute) {
            x = cell.getAttribute('data-x');
            y = cell.getAttribute('data-y');
        }

        // Update cell
        if (x != null && y != null) {
            records.push(updateCell.call(obj, x, y, value, force));

            // Update all formulas in the chain
            updateFormulaChain.call(obj, x, y, records);
        } else {
            const keys = Object.keys(cell);
            if (keys.length > 0) {
                for (let i = 0; i < keys.length; i++) {
                    let x, y;
                    // An item without its own value gets the value argument
                    let cellValue = value;

                    if (typeof cell[i] == 'string') {
                        const columnId = getIdFromColumnName(cell[i], true);
                        x = columnId[0];
                        y = columnId[1];
                    } else {
                        if (cell[i].x != null && cell[i].y != null) {
                            x = cell[i].x;
                            y = cell[i].y;
                            // Flexible setup
                            if (cell[i].value != null) {
                                cellValue = cell[i].value;
                            }
                        } else {
                            x = cell[i].getAttribute('data-x');
                            y = cell[i].getAttribute('data-y');
                        }
                    }

                    // Update cell
                    if (x != null && y != null) {
                        records.push(updateCell.call(obj, x, y, cellValue, force));

                        // Update all formulas in the chain
                        updateFormulaChain.call(obj, x, y, records);
                    }
                }
            }
        }
    }

    // Update history
    obj.history.add({
        action: 'setValue',
        records: records,
        selection: obj.selection.range && [...obj.selection.range],
    });

    // Update table with custom configurations if applicable
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

/**
 * Set a cell value based on coordinates
 *
 * @param int x destination cell
 * @param int y destination cell
 * @param string value
 * @return void
 */
export const setValueFromCoords: WorksheetMethod<'setValueFromCoords'> = function (x, y, value, force) {
    const obj = this;

    const records = [];
    records.push(updateCell.call(obj, x, y, value, force));

    // Update all formulas in the chain
    updateFormulaChain.call(obj, x, y, records);

    // Update history
    obj.history.add({
        action: 'setValue',
        records: records,
        selection: obj.selection.range && [...obj.selection.range],
    });

    // Update table with custom configurations if applicable
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

/**
 * Get the whole table data
 *
 * @param bool get highlighted cells only
 * @return array data
 */
// The public typings have one overload for each kind of result: rows of values, text with a delimiter, or objects with asJson
export const getData = function (this: Worksheet, highlighted?: boolean, processed?: boolean, delimiter?: string, asJson?: boolean) {
    const obj = this;

    // Control vars
    const dataset: CellValue[][] = [];
    let px = 0;
    let py = 0;

    // Column and row length
    const x = Math.max(
        ...obj.options.data.map(function (row) {
            return row.length;
        })
    );
    const y = obj.options.data.length;

    // Go through the columns to get the data
    for (let j = 0; j < y; j++) {
        px = 0;
        for (let i = 0; i < x; i++) {
            // Cell selected or fullset
            if (!highlighted || obj.records[j][i].element.classList.contains('highlight')) {
                // Get value
                if (!dataset[py]) {
                    dataset[py] = [];
                }
                if (processed) {
                    dataset[py][px] = obj.records[j][i].element.innerHTML;
                } else {
                    dataset[py][px] = obj.options.data[j][i];
                }
                px++;
            }
        }
        if (px > 0) {
            py++;
        }
    }

    if (delimiter) {
        return (
            dataset
                .map(function (row) {
                    return row.join(delimiter);
                })
                .join('\r\n') + '\r\n'
        );
    }

    if (asJson) {
        return dataset.map(function (row) {
            const resultRow: Record<number, CellValue> = {};

            row.forEach(function (item, index) {
                resultRow[index] = item;
            });

            return resultRow;
        });
    }

    return dataset;
};

export const getDataFromRange: WorksheetMethod<'getDataFromRange'> = function (range, processed) {
    const obj = this;

    const coords = getCoordsFromRange(range);

    const dataset: CellValue[][] = [];

    for (let y = coords[1]; y <= coords[3]; y++) {
        dataset.push([]);

        for (let x = coords[0]; x <= coords[2]; x++) {
            if (processed) {
                dataset[dataset.length - 1].push(obj.records[y][x].element.innerHTML);
            } else {
                dataset[dataset.length - 1].push(obj.options.data[y][x]);
            }
        }
    }

    return dataset;
};
