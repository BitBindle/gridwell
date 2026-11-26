import { getIdFromColumnName } from '../utils/internalHelpers.js';
import { getCoordsFromRange } from '../utils/helpers.js';
import type { Worksheet } from '../types/internal';
import type { CellChange, CellValue } from '../jspreadsheet';

/**
 * The data of a worksheet: it gets and sets all the data, the data of a range, and the value of a cell by its name or
 * its coordinates, and updates the table after the cells change.
 *
 * Each worksheet has one, as worksheet.data. The data itself stays in worksheet.options.data.
 */
export class Data {
    readonly worksheet: Worksheet;

    constructor(worksheet: Worksheet) {
        this.worksheet = worksheet;
    }

    // data can be rows of objects. They are changed to arrays.
    set(data?: CellValue[][] | Record<string, CellValue>[]): void {
        const worksheet = this.worksheet;

        // Update data. Rows of objects are changed to arrays below.
        if (data) {
            worksheet.options.data = data as CellValue[][];
        }

        // Data
        if (!worksheet.options.data) {
            worksheet.options.data = [];
        }

        // Prepare data
        if (worksheet.options.data && worksheet.options.data[0]) {
            if (!Array.isArray(worksheet.options.data[0])) {
                const rows: CellValue[][] = [];
                for (const record of worksheet.options.data) {
                    const row: CellValue[] = [];
                    for (let i = 0; i < worksheet.options.columns.length; i++) {
                        // The rows are objects here
                        row[i] = (record as unknown as Record<string, CellValue>)[worksheet.options.columns[i].name!];
                    }
                    rows.push(row);
                }

                worksheet.options.data = rows;
            }
        }

        // Adjust minimal dimensions
        let j = 0;
        let i = 0;
        const size_i = (worksheet.options.columns && worksheet.options.columns.length) || 0;
        const size_j = worksheet.options.data.length;
        const min_i = worksheet.options.minDimensions[0];
        const min_j = worksheet.options.minDimensions[1];
        const max_i = min_i > size_i ? min_i : size_i;
        const max_j = min_j > size_j ? min_j : size_j;

        for (j = 0; j < max_j; j++) {
            for (i = 0; i < max_i; i++) {
                if (worksheet.options.data[j] == undefined) {
                    worksheet.options.data[j] = [];
                }

                if (worksheet.options.data[j][i] == undefined) {
                    worksheet.options.data[j][i] = '';
                }
            }
        }

        // Reset containers
        worksheet.rows = [];
        worksheet.view.results = null;
        worksheet.records = [];
        worksheet.history.clear();

        // Reset data
        worksheet.tbody.innerHTML = '';

        let startNumber;
        let finalNumber;

        // Lazy loading
        if (worksheet.options.lazyLoading == true) {
            // Load only 100 records
            startNumber = 0;
            finalNumber = worksheet.options.data.length < 100 ? worksheet.options.data.length : 100;

            if (worksheet.options.pagination) {
                // pagination is declared as a number, and false turns it off as 0 does
                (worksheet.options as { pagination?: number | false }).pagination = false;
                console.error('Jspreadsheet: Pagination will be disable due the lazyLoading');
            }
        } else if (worksheet.options.pagination) {
            // Pagination
            if (!worksheet.view.pageNumber) {
                worksheet.view.pageNumber = 0;
            }
            startNumber = worksheet.options.pagination * worksheet.view.pageNumber;
            finalNumber = worksheet.options.pagination * worksheet.view.pageNumber + worksheet.options.pagination;

            if (worksheet.options.data.length < finalNumber) {
                finalNumber = worksheet.options.data.length;
            }
        } else {
            startNumber = 0;
            finalNumber = worksheet.options.data.length;
        }

        // Append nodes to the HTML
        for (j = 0; j < worksheet.options.data.length; j++) {
            // Create row
            const row = worksheet.rowManager.create(j, worksheet.options.data[j]);
            // Append line to the table
            if (j >= startNumber && j < finalNumber) {
                worksheet.tbody.appendChild(row.element);
            }
        }

        if (worksheet.options.lazyLoading == true) {
            // Do not create pagination with lazyloading activated
        } else if (worksheet.options.pagination) {
            worksheet.view.updatePagination();
        }

        // Merge cells
        if (worksheet.options.mergeCells) {
            for (const key of Object.keys(worksheet.options.mergeCells)) {
                const num = worksheet.options.mergeCells[key];
                worksheet.merges.set(key, num[0], num[1], 1);
            }
        }

        // Update the size of the grid
        worksheet.accessibility.updateGridSize();

        // The rows are new, so the tab stop of the grid is new
        worksheet.selection.updateTabStop();

        // Updata table with custom configurations if applicable
        this.updateTable();
    }

    /**
     * Get the value from a cell
     *
     * @param object cell
     * @return string value
     */
    getValue(cell: string, processedValue?: boolean): CellValue | null {
        const worksheet = this.worksheet;

        if (typeof cell !== 'string') {
            return null;
        }

        const coords = getIdFromColumnName(cell, true);
        const x = coords[0];
        const y = coords[1];

        let value = null;

        if (x != null && y != null) {
            if (worksheet.records[y] && worksheet.records[y][x] && processedValue) {
                value = worksheet.records[y][x].element.innerHTML;
            } else {
                if (worksheet.options.data[y] && worksheet.options.data[y][x] != 'undefined') {
                    value = worksheet.options.data[y][x];
                }
            }
        }

        return value;
    }

    /**
     * Get the value from a coords
     *
     * @param int x
     * @param int y
     * @return string value
     */
    getValueFromCoords(x: number, y: number, processedValue?: boolean): CellValue | null {
        const worksheet = this.worksheet;

        let value = null;

        if (x != null && y != null) {
            if (worksheet.records[y] && worksheet.records[y][x] && processedValue) {
                value = worksheet.records[y][x].element.innerHTML;
            } else {
                if (worksheet.options.data[y] && worksheet.options.data[y][x] != 'undefined') {
                    value = worksheet.options.data[y][x];
                }
            }
        }

        return value;
    }

    /**
     * Set a cell value
     *
     * @param mixed cell destination cell
     * @param string value value
     * @return void
     */
    // any: a cell element gives the text of data-x and data-y to updateCell (issue #52)
    setValue(cell: any, value?: CellValue, force?: boolean): void {
        const worksheet = this.worksheet;

        const records = [];

        if (typeof cell == 'string') {
            const columnId = getIdFromColumnName(cell, true);
            const x = columnId[0];
            const y = columnId[1];

            // Update cell
            records.push(worksheet.cells.update(x, y, value, force));

            // Update all formulas in the chain
            worksheet.formulas.updateChain(x, y, records);
        } else {
            let x = null;
            let y = null;
            if (cell && cell.getAttribute) {
                x = cell.getAttribute('data-x');
                y = cell.getAttribute('data-y');
            }

            // Update cell
            if (x != null && y != null) {
                records.push(worksheet.cells.update(x, y, value, force));

                // Update all formulas in the chain
                worksheet.formulas.updateChain(x, y, records);
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
                            records.push(worksheet.cells.update(x, y, cellValue, force));

                            // Update all formulas in the chain
                            worksheet.formulas.updateChain(x, y, records);
                        }
                    }
                }
            }
        }

        // Update history
        worksheet.history.add({
            action: 'setValue',
            records,
            selection: worksheet.selection.range && [...worksheet.selection.range],
        });

        // Update table with custom configurations if applicable
        this.updateTable();

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
     * Set a cell value based on coordinates
     *
     * @param int x destination cell
     * @param int y destination cell
     * @param string value
     * @return void
     */
    setValueFromCoords(x: number, y: number, value: CellValue, force?: boolean): void {
        const worksheet = this.worksheet;

        const records = [];
        records.push(worksheet.cells.update(x, y, value, force));

        // Update all formulas in the chain
        worksheet.formulas.updateChain(x, y, records);

        // Update history
        worksheet.history.add({
            action: 'setValue',
            records,
            selection: worksheet.selection.range && [...worksheet.selection.range],
        });

        // Update table with custom configurations if applicable
        this.updateTable();

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
     * Get the whole table data
     *
     * @param bool get highlighted cells only
     * @return array data
     */
    // The public typings have one overload for each kind of result: rows of values, text with a delimiter, or objects with asJson
    get(highlighted?: boolean, processed?: boolean, delimiter?: string | null, asJson?: boolean) {
        const worksheet = this.worksheet;

        // Control vars
        const dataset: CellValue[][] = [];
        let px = 0;
        let py = 0;

        // Column and row length
        const x = Math.max(
            ...worksheet.options.data.map((row) => {
                return row.length;
            })
        );
        const y = worksheet.options.data.length;

        // Go through the columns to get the data
        for (let j = 0; j < y; j++) {
            px = 0;
            for (let i = 0; i < x; i++) {
                // Cell selected or fullset
                if (!highlighted || worksheet.records[j][i].element.classList.contains('highlight')) {
                    // Get value
                    if (!dataset[py]) {
                        dataset[py] = [];
                    }
                    if (processed) {
                        dataset[py][px] = worksheet.records[j][i].element.innerHTML;
                    } else {
                        dataset[py][px] = worksheet.options.data[j][i];
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
                    .map((row) => {
                        return row.join(delimiter);
                    })
                    .join('\r\n') + '\r\n'
            );
        }

        if (asJson) {
            return dataset.map((row) => {
                const resultRow: Record<number, CellValue> = {};

                row.forEach((item, index) => {
                    resultRow[index] = item;
                });

                return resultRow;
            });
        }

        return dataset;
    }

    getFromRange(range: string, processed: true): CellValue[][] {
        const worksheet = this.worksheet;

        const coords = getCoordsFromRange(range);

        const dataset: CellValue[][] = [];

        for (let y = coords[1]; y <= coords[3]; y++) {
            dataset.push([]);

            for (let x = coords[0]; x <= coords[2]; x++) {
                if (processed) {
                    dataset[dataset.length - 1].push(worksheet.records[y][x].element.innerHTML);
                } else {
                    dataset[dataset.length - 1].push(worksheet.options.data[y][x]);
                }
            }
        }

        return dataset;
    }

    /**
     * Run after the cells change: add the spare rows and columns that minSpareRows and minSpareCols ask for, update the
     * footers and place the selection corner again.
     */
    updateTable() {
        const worksheet = this.worksheet;

        // Check for spare
        if (worksheet.options.minSpareRows! > 0) {
            let numBlankRows = 0;
            for (let j = worksheet.rows.length - 1; j >= 0; j--) {
                let test = false;
                for (let i = 0; i < worksheet.headers.length; i++) {
                    if (worksheet.options.data[j][i]) {
                        test = true;
                    }
                }
                if (test) {
                    break;
                } else {
                    numBlankRows++;
                }
            }

            if (worksheet.options.minSpareRows! - numBlankRows > 0) {
                worksheet.insertRow(worksheet.options.minSpareRows! - numBlankRows);
            }
        }

        if (worksheet.options.minSpareCols! > 0) {
            let numBlankCols = 0;
            for (let i = worksheet.headers.length - 1; i >= 0; i--) {
                let test = false;
                for (let j = 0; j < worksheet.rows.length; j++) {
                    if (worksheet.options.data[j][i]) {
                        test = true;
                    }
                }
                if (test) {
                    break;
                } else {
                    numBlankCols++;
                }
            }

            if (worksheet.options.minSpareCols! - numBlankCols > 0) {
                worksheet.insertColumn(worksheet.options.minSpareCols! - numBlankCols);
            }
        }

        // Update footers
        if (worksheet.options.footers) {
            worksheet.formulas.setFooter();
        }

        if (worksheet.options.columns.length < worksheet.options.minDimensions[0]) {
            worksheet.options.minDimensions[0] = worksheet.options.columns.length;
        }

        // Update corner position
        setTimeout(() => {
            worksheet.selection.updateCorner();
        }, 0);
    }
}
