import { getIdFromColumnName } from '../utils/internalHelpers.js';
import { getCoordsFromRange } from '../utils/helpers.js';
import type { Worksheet } from '../types/internal';
import { RowSource, recordRows } from './RowSource.js';
import type { CellValue, WorksheetInstance } from '../jspreadsheet';

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

        // In the row source mode, the application gives the rows when the window shows them
        worksheet.rowSource?.stop();
        worksheet.rowSource = null;
        if (worksheet.options.rowSource) {
            worksheet.view.results = null;
            worksheet.history.clear();
            worksheet.rowSource = new RowSource(worksheet, worksheet.options.rowSource);
            worksheet.rowSource.start();
            worksheet.accessibility.updateGridSize();
            worksheet.selection.updateTabStop();
            worksheet.rowManager.updateFreeze();
            return;
        }

        // Update data. Rows of objects are changed to arrays below.
        if (data) {
            worksheet.options.data = data as CellValue[][];
        }

        // Data
        if (!worksheet.options.data) {
            worksheet.options.data = [];
        }

        // Prepare data
        if (worksheet.options.data?.[0]) {
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
        const size_i = worksheet.options.columns?.length || 0;
        const size_j = worksheet.options.data.length;
        const min_i = worksheet.options.minDimensions[0];
        const min_j = worksheet.options.minDimensions[1];
        const max_i = min_i > size_i ? min_i : size_i;
        const max_j = min_j > size_j ? min_j : size_j;

        for (let j = 0; j < max_j; j++) {
            for (let i = 0; i < max_i; i++) {
                if (worksheet.options.data[j] == null) {
                    worksheet.options.data[j] = [];
                }

                if (worksheet.options.data[j][i] == null) {
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
        if (worksheet.options.lazyLoading === true) {
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
        for (let j = 0; j < worksheet.options.data.length; j++) {
            // Create row
            const row = worksheet.rowManager.create(j, worksheet.options.data[j]);
            // Append line to the table
            if (j >= startNumber && j < finalNumber) {
                worksheet.tbody.appendChild(row.element);
            }
        }

        if (worksheet.options.lazyLoading === true) {
            // Do not create pagination with lazyloading activated
        } else if (worksheet.options.pagination) {
            worksheet.view.updatePagination();
        }

        // Merge cells
        if (worksheet.options.mergeCells) {
            for (const key of Object.keys(worksheet.options.mergeCells)) {
                const num = worksheet.options.mergeCells[key];
                worksheet.merges.set(key, num[0], num[1], 1, true);
            }
        }

        // Update the size of the grid
        worksheet.accessibility.updateGridSize();

        // The rows are new, so the tab stop of the grid is new
        worksheet.selection.updateTabStop();

        // The frozen rows are new
        worksheet.rowManager.updateFreeze();

        // Updata table with custom configurations if applicable
        this.updateTable();
    }

    /**
     * Get the value of a cell.
     * @param cell - The cell name, for example "B3".
     * @param processedValue - Get the value that the cell shows instead of the data.
     */
    getValue(cell: string, processedValue?: boolean): CellValue | null {
        const worksheet = this.worksheet;

        if (typeof cell !== 'string') {
            return null;
        }

        const coords = getIdFromColumnName(cell);
        const x = coords[0];
        const y = coords[1];

        let value = null;

        if (x != null && y != null) {
            if (worksheet.records[y]?.[x] && processedValue) {
                value = worksheet.records[y][x].element.innerHTML;
            } else {
                if (typeof worksheet.options.data[y]?.[x] !== 'undefined') {
                    value = worksheet.options.data[y][x];
                }
            }
        }

        return value;
    }

    /**
     * Get the value of a cell from its coordinates.
     * @param x - The column.
     * @param y - The row.
     * @param processedValue - Get the value that the cell shows instead of the data.
     */
    getValueFromCoords(x: number, y: number, processedValue?: boolean): CellValue | null {
        const worksheet = this.worksheet;

        let value = null;

        if (x != null && y != null) {
            if (worksheet.records[y]?.[x] && processedValue) {
                value = worksheet.records[y][x].element.innerHTML;
            } else {
                if (typeof worksheet.options.data[y]?.[x] !== 'undefined') {
                    value = worksheet.options.data[y][x];
                }
            }
        }

        return value;
    }

    /**
     * Set the value of cells.
     * @param cell - The cell name, the td of the cell, or a list of cell names, tds or { x, y, value } objects.
     * @param value - The new value.
     * @param force - Also change read-only cells.
     */
    setValue(cell: Parameters<WorksheetInstance['setValue']>[0], value?: CellValue, force?: boolean): void {
        const worksheet = this.worksheet;

        const records = [];

        // The cells to change, with their coordinates as numbers, as for a cell name. A list can also be a NodeList.
        const cells: [number, number, CellValue | undefined][] = [];
        for (const item of typeof cell === 'string' || 'getAttribute' in cell ? [cell] : Object.values(cell)) {
            if (typeof item === 'string') {
                const columnId = getIdFromColumnName(item);
                cells.push([columnId[0], columnId[1], value]);
            } else if ('getAttribute' in item) {
                const x = item.getAttribute('data-x');
                const y = item.getAttribute('data-y');
                if (x != null && y != null) {
                    cells.push([parseInt(x), parseInt(y), value]);
                }
            } else if (item.x != null && item.y != null) {
                // An item without its own value gets the value argument
                cells.push([item.x, item.y, item.value != null ? item.value : value]);
            }
        }

        for (const [x, y, cellValue] of cells) {
            const record = worksheet.cells.update(x, y, cellValue, force);
            // A cell that did not change has no record
            if (record) {
                records.push(record);

                // Update all formulas in the chain
                worksheet.formulas.updateChain(x, y, records);
            }
        }

        // Nothing changed, for example in a read-only cell
        if (!records.length) {
            return;
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

        worksheet.dispatch('onafterchanges', worksheet, onafterchangesRecords);
    }

    /**
     * Set the value of a cell from its coordinates.
     * @param x - The column.
     * @param y - The row.
     * @param value - The new value.
     * @param force - Also change a read-only cell.
     */
    setValueFromCoords(x: number, y: number, value: CellValue, force?: boolean): void {
        const worksheet = this.worksheet;

        const records = [];
        const record = worksheet.cells.update(x, y, value, force);
        // A cell that did not change has no record
        if (record) {
            records.push(record);

            // Update all formulas in the chain
            worksheet.formulas.updateChain(x, y, records);
        }

        // Nothing changed, for example in a read-only cell
        if (!records.length) {
            return;
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

        worksheet.dispatch('onafterchanges', worksheet, onafterchangesRecords);
    }

    /**
     * Get the data of the worksheet.
     * @param highlighted - Get only the highlighted cells.
     * @param processed - Get the values that the cells show instead of the data.
     * @param delimiter - Return text, with this delimiter between the columns.
     * @param asJson - Return the rows as objects, by column name.
     */
    // The public typings have one overload for each kind of result: rows of values, text with a delimiter, or objects with asJson
    get(highlighted?: boolean, processed?: boolean, delimiter?: string | null, asJson?: boolean) {
        const worksheet = this.worksheet;

        // Control vars
        const dataset: CellValue[][] = [];
        let py = 0;

        // The rows with a record: in the row source mode, the frozen rows and the window
        const rows = recordRows(worksheet);

        // Column length
        const x = Math.max(
            0,
            ...rows.map((j) => {
                return worksheet.options.data[j].length;
            })
        );

        // Go through the columns to get the data
        for (const j of rows) {
            let px = 0;
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
            return `${dataset
                .map((row) => {
                    return row.join(delimiter);
                })
                .join('\r\n')}\r\n`;
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

    getFromRange(range: string, processed?: boolean): CellValue[][] {
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
