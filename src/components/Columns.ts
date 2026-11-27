import { translate } from '../utils/dictionary.js';
import { getColumnName } from '../utils/helpers.js';
import { getColumnNameFromId, injectArray } from '../utils/internalHelpers.js';
import { setHeaderRole } from './Accessibility.js';
import type jspreadsheet from '../jspreadsheet';
import type { Worksheet } from '../types/internal';
import { toGrid, toScreen } from '../utils/zoom.js';
import { recordRows } from './RowSource.js';
import { askConfirm } from '../utils/dialogs.js';

/**
 * The columns of a worksheet: it counts them, creates their headers, inserts, moves, deletes, hides and shows them, gets
 * and sets their widths, their titles and their data, and places the frozen columns.
 *
 * Each worksheet has one, as worksheet.columns. insertColumn, moveColumn, deleteColumn, getWidth, setWidth, showColumn,
 * hideColumn, getColumnData, setColumnData, getHeader, getHeaders and setHeader call it.
 */
export class Columns {
    readonly worksheet: Worksheet;

    constructor(worksheet: Worksheet) {
        this.worksheet = worksheet;
    }

    count() {
        const worksheet = this.worksheet;

        let numberOfColumns = worksheet.options.columns?.length || 0;

        if (typeof worksheet.options.data?.[0] !== 'undefined') {
            // Data keys
            const keys = Object.keys(worksheet.options.data[0]);

            if (keys.length > numberOfColumns) {
                numberOfColumns = keys.length;
            }
        }

        if (worksheet.options.minDimensions && worksheet.options.minDimensions[0] > numberOfColumns) {
            numberOfColumns = worksheet.options.minDimensions[0];
        }

        return numberOfColumns;
    }

    createHeader(colNumber: number) {
        const worksheet = this.worksheet;

        // Create col global control
        const colWidth = worksheet.options.columns?.[colNumber]?.width || worksheet.options.defaultColWidth || 100;
        const colAlign = worksheet.options.columns?.[colNumber]?.align || worksheet.options.defaultColAlign || 'center';

        // Create header cell
        worksheet.headers[colNumber] = document.createElement('td');
        worksheet.headers[colNumber].textContent = worksheet.options.columns?.[colNumber]?.title || getColumnName(colNumber);

        worksheet.headers[colNumber].setAttribute('data-x', colNumber);
        setHeaderRole(worksheet.headers[colNumber], colNumber);
        worksheet.headers[colNumber].style.textAlign = colAlign;
        if (worksheet.options.columns?.[colNumber]?.title) {
            worksheet.headers[colNumber].setAttribute('title', worksheet.headers[colNumber].innerText);
        }
        // id is not in the public typings of a column
        if (worksheet.options.columns?.[colNumber]?.id) {
            worksheet.headers[colNumber].setAttribute('id', worksheet.options.columns[colNumber].id);
        }

        // Width control
        const colElement = document.createElement('col');
        colElement.setAttribute('width', colWidth);

        worksheet.cols[colNumber] = {
            colElement,
            x: colNumber,
        };

        // Hidden column
        if (worksheet.options.columns?.[colNumber]?.type === 'hidden') {
            worksheet.headers[colNumber].style.display = 'none';
            colElement.style.display = 'none';
        }
    }

    /**
     * The columns that insertColumn inserts, with their position, their options and their data, for the events.
     */
    private getInserted(
        mixed: number | jspreadsheet.CellValue[] | undefined,
        columnNumber: number,
        numOfColumns: number,
        insertBefore: boolean,
        properties: jspreadsheet.Column[]
    ) {
        const worksheet = this.worksheet;

        const columns = [];

        if (!Array.isArray(mixed)) {
            for (let i = 0; i < numOfColumns; i++) {
                const column = {
                    column: columnNumber + i + (insertBefore ? 0 : 1),
                    options: Object.assign({}, properties[i]),
                };

                columns.push(column);
            }
        } else {
            const data = [];

            for (let i = 0; i < worksheet.options.data.length; i++) {
                data.push(i < mixed.length ? mixed[i] : '');
            }

            const column = {
                column: columnNumber + (insertBefore ? 0 : 1),
                options: Object.assign({}, properties[0]),
                data,
            };

            columns.push(column);
        }

        return columns;
    }

    /**
     * Create the headers and the columns of the new columns, and put them in the head of the table.
     */
    private addHeaders(columnIndex: number, numOfColumns: number) {
        const worksheet = this.worksheet;

        const historyHeaders = [];
        const historyColgroup = [];

        for (let col = columnIndex; col < numOfColumns + columnIndex; col++) {
            this.createHeader(col);
            worksheet.headerContainer.insertBefore(worksheet.headers[col], worksheet.headerContainer.children[col + 1]);
            worksheet.colgroupContainer.insertBefore(worksheet.cols[col].colElement, worksheet.colgroupContainer.children[col + 1]);

            historyHeaders.push(worksheet.headers[col]);
            historyColgroup.push(worksheet.cols[col]);
        }

        return { historyHeaders, historyColgroup };
    }

    /**
     * Add an empty cell for each new column to the footers.
     */
    private addFooters(columnIndex: number, numOfColumns: number) {
        const worksheet = this.worksheet;

        const historyFooters: string[][] = [];

        if (worksheet.options.footers) {
            for (let j = 0; j < worksheet.options.footers.length; j++) {
                historyFooters[j] = [];
                for (let i = 0; i < numOfColumns; i++) {
                    historyFooters[j].push('');
                }
                worksheet.options.footers[j].splice(columnIndex, 0, ...historyFooters[j]);
            }
        }

        return historyFooters;
    }

    /**
     * Create the cells of the new columns in each row, with the data or empty.
     */
    private addCells(columnIndex: number, numOfColumns: number, data: jspreadsheet.CellValue[]) {
        const worksheet = this.worksheet;

        const historyRecords: { element: HTMLTableCellElement; x: number; y: number }[][] = [];
        const historyData: jspreadsheet.CellValue[][] = [];

        for (let row = 0; row < worksheet.options.data.length; row++) {
            // Keep the current data
            const currentData = worksheet.options.data[row].splice(columnIndex);
            const currentRecord = worksheet.records[row].splice(columnIndex);

            // History
            historyData[row] = [];
            historyRecords[row] = [];

            for (let col = columnIndex; col < numOfColumns + columnIndex; col++) {
                // New value
                const value = data[row] ? data[row] : '';
                worksheet.options.data[row][col] = value;
                // New cell
                const td = worksheet.cells.create(col, row, worksheet.options.data[row][col]);
                // updateTableReferences adds x
                worksheet.records[row][col] = {
                    element: td,
                    y: row,
                } as Worksheet['records'][number][number];
                // Add cell to the row
                if (worksheet.rows[row]) {
                    worksheet.rows[row].element.insertBefore(td, worksheet.rows[row].element.children[col + 1]);
                }

                if (worksheet.options.columns?.[col] && typeof worksheet.options.columns[col].render === 'function') {
                    // The test above makes it a function, which TypeScript does not narrow through the index
                    (worksheet.options.columns[col].render as NonNullable<jspreadsheet.BaseColumn['render']>)(
                        td,
                        value,
                        parseInt(col),
                        parseInt(row),
                        worksheet,
                        worksheet.options.columns[col]
                    );
                }

                // Record History
                historyData[row].push(value);
                historyRecords[row].push({ element: td, x: col, y: row });
            }

            // Copy the data back to the main data
            worksheet.options.data[row].push(...currentData);
            worksheet.records[row].push(...currentRecord);
        }

        return { historyRecords, historyData };
    }

    /**
     * Widen the last nested header of each row over the new columns.
     */
    private widenNestedHeaders(columnIndex: number, numOfColumns: number) {
        const worksheet = this.worksheet;

        if (worksheet.options.nestedHeaders && worksheet.options.nestedHeaders.length > 0 && worksheet.options.nestedHeaders[0]?.[0]) {
            for (let j = 0; j < worksheet.options.nestedHeaders.length; j++) {
                const colspan = parseInt(worksheet.options.nestedHeaders[j][worksheet.options.nestedHeaders[j].length - 1].colspan) + numOfColumns;
                worksheet.options.nestedHeaders[j][worksheet.options.nestedHeaders[j].length - 1].colspan = colspan;
                worksheet.thead.children[j].children[worksheet.thead.children[j].children.length - 1].setAttribute('colspan', colspan);
                const o: (string | number)[] = worksheet.thead.children[j].children[worksheet.thead.children[j].children.length - 1]
                    .getAttribute('data-column')!
                    .split(',');
                for (let col = columnIndex; col < numOfColumns + columnIndex; col++) {
                    o.push(col);
                }
                worksheet.thead.children[j].children[worksheet.thead.children[j].children.length - 1].setAttribute('data-column', o.join(','));
            }
        }
    }

    /**
     * Insert columns.
     * @param mixed - The number of columns, or the data of one new column.
     * @param columnNumber - The column next to which the columns go. The last column by default.
     * @param insertBefore - Put the columns before that column instead of after it.
     * @param properties - The options of the new columns.
     */
    insert(
        mixed?: number | jspreadsheet.CellValue[],
        columnNumber?: number,
        insertBefore?: boolean | number,
        properties?: jspreadsheet.Column[]
    ): false | undefined {
        const worksheet = this.worksheet;

        // The row source mode keeps the columns of the build
        if (worksheet.rowSource) {
            return false;
        }

        // Configuration
        if (worksheet.options.allowInsertColumn !== false) {
            // Data to be insert
            let data: jspreadsheet.CellValue[] = [];

            // The insert could be lead by number of rows or the array of data
            let numOfColumns;
            if (!Array.isArray(mixed)) {
                numOfColumns = typeof mixed === 'number' ? mixed : 1;
            } else {
                numOfColumns = 1;

                if (mixed) {
                    data = mixed;
                }
            }

            // Direction
            insertBefore = insertBefore ? true : false;

            // Current column number
            const currentNumOfColumns = Math.max(
                worksheet.options.columns.length,
                ...worksheet.options.data.map((row) => {
                    return row.length;
                })
            );

            const lastColumn = currentNumOfColumns - 1;

            // Confirm position
            if (columnNumber == null || columnNumber >= parseInt(lastColumn) || columnNumber < 0) {
                columnNumber = lastColumn;
            }

            // Create default properties
            if (!properties) {
                properties = [];
            }

            for (let i = 0; i < numOfColumns; i++) {
                if (!properties[i]) {
                    properties[i] = {};
                }
            }

            const columns = this.getInserted(mixed, columnNumber, numOfColumns, insertBefore, properties);

            // Onbeforeinsertcolumn
            if (worksheet.dispatch('onbeforeinsertcolumn', worksheet, columns) === false) {
                return false;
            }

            // Merged cells
            if (worksheet.options.mergeCells && Object.keys(worksheet.options.mergeCells).length > 0) {
                if (worksheet.merges.isColMerged(columnNumber, insertBefore).length) {
                    if (!askConfirm(worksheet.parent.config, translate('This action will destroy any existing merged cells. Are you sure?'))) {
                        return false;
                    } else {
                        worksheet.destroyMerge();
                    }
                }
            }

            // Insert before
            const columnIndex = !insertBefore ? columnNumber + 1 : columnNumber;
            worksheet.options.columns = injectArray(worksheet.options.columns, columnIndex, properties);

            // Open space in the containers
            const currentHeaders = worksheet.headers.splice(columnIndex);
            const currentColgroup = worksheet.cols.splice(columnIndex);

            // Add new headers
            const { historyHeaders, historyColgroup } = this.addHeaders(columnIndex, numOfColumns);

            // Add new footer cells
            const historyFooters = this.addFooters(columnIndex, numOfColumns);

            // Adding visual columns
            const { historyRecords, historyData } = this.addCells(columnIndex, numOfColumns, data);

            worksheet.headers.push(...currentHeaders);
            worksheet.cols.push(...currentColgroup);

            for (let i = columnIndex; i < worksheet.cols.length; i++) {
                worksheet.cols[i].x = i;
            }

            for (const record of worksheet.records) {
                for (let i = 0; i < record.length; i++) {
                    record[i].x = i;
                }
            }

            // Adjust nested headers
            this.widenNestedHeaders(columnIndex, numOfColumns);

            // Keep history
            worksheet.history.add({
                action: 'insertColumn',
                columnNumber,
                numOfColumns,
                insertBefore,
                columns: properties,
                headers: historyHeaders,
                cols: historyColgroup,
                records: historyRecords,
                footers: historyFooters,
                data: historyData,
            });

            // Remove table references
            worksheet.formulas.updateTableReferences({ column: Number(columnNumber) });

            // The frozen columns can be other columns now
            this.updateFrozenColumns();

            // Events
            worksheet.dispatch('oninsertcolumn', worksheet, columns);
        }
    }

    /**
     * Move a column.
     * @param o - The column to move.
     * @param d - The position it goes to.
     */
    move(o: number, d: number): false | undefined {
        const worksheet = this.worksheet;

        // The row source mode keeps the columns of the build
        if (worksheet.rowSource) {
            return false;
        }

        if (worksheet.options.mergeCells && Object.keys(worksheet.options.mergeCells).length > 0) {
            let insertBefore: 0 | 1;
            if (o > d) {
                insertBefore = 1;
            } else {
                insertBefore = 0;
            }

            if (worksheet.merges.isColMerged(o).length || worksheet.merges.isColMerged(d, insertBefore).length) {
                if (!askConfirm(worksheet.parent.config, translate('This action will destroy any existing merged cells. Are you sure?'))) {
                    return false;
                } else {
                    worksheet.destroyMerge();
                }
            }
        }

        o = parseInt(o);
        d = parseInt(d);

        if (o > d) {
            worksheet.headerContainer.insertBefore(worksheet.headers[o], worksheet.headers[d]);
            worksheet.colgroupContainer.insertBefore(worksheet.cols[o].colElement, worksheet.cols[d].colElement);

            for (let j = 0; j < worksheet.rows.length; j++) {
                worksheet.rows[j].element.insertBefore(worksheet.records[j][o].element, worksheet.records[j][d].element);
            }
        } else {
            worksheet.headerContainer.insertBefore(worksheet.headers[o], worksheet.headers[d].nextSibling);
            worksheet.colgroupContainer.insertBefore(worksheet.cols[o].colElement, worksheet.cols[d].colElement.nextSibling);

            for (let j = 0; j < worksheet.rows.length; j++) {
                worksheet.rows[j].element.insertBefore(worksheet.records[j][o].element, worksheet.records[j][d].element.nextSibling);
            }
        }

        worksheet.options.columns.splice(d, 0, worksheet.options.columns.splice(o, 1)[0]);
        worksheet.headers.splice(d, 0, worksheet.headers.splice(o, 1)[0]);
        worksheet.cols.splice(d, 0, worksheet.cols.splice(o, 1)[0]);

        const firstAffectedIndex = Math.min(o, d);
        const lastAffectedIndex = Math.max(o, d);

        for (let j = 0; j < worksheet.rows.length; j++) {
            worksheet.options.data[j].splice(d, 0, worksheet.options.data[j].splice(o, 1)[0]);
            worksheet.records[j].splice(d, 0, worksheet.records[j].splice(o, 1)[0]);
        }

        for (let i = firstAffectedIndex; i <= lastAffectedIndex; i++) {
            worksheet.cols[i].x = i;
        }

        for (const record of worksheet.records) {
            for (let i = firstAffectedIndex; i <= lastAffectedIndex; i++) {
                record[i].x = i;
            }
        }

        // Update footers position
        if (worksheet.options.footers) {
            for (const footer of worksheet.options.footers) {
                footer.splice(d, 0, footer.splice(o, 1)[0]);
            }
        }

        // Keeping history of changes
        worksheet.history.add({
            action: 'moveColumn',
            oldValue: o,
            newValue: d,
        });

        // Update table references
        worksheet.formulas.updateTableReferences({ column: Math.min(Number(o), Number(d)) });

        // The frozen columns can be other columns now
        this.updateFrozenColumns();

        // Events
        worksheet.dispatch('onmovecolumn', worksheet, o, d, 1);
    }

    /**
     * Delete columns.
     * @param columnNumber - The first column to delete. Without it, the selected columns, or the last column.
     * @param numOfColumns - The number of columns to delete.
     */
    delete(columnNumber?: number, numOfColumns?: number): false | undefined {
        const worksheet = this.worksheet;

        // The row source mode keeps the columns of the build
        if (worksheet.rowSource) {
            return false;
        }

        // Global Configuration
        if (worksheet.options.allowDeleteColumn !== false) {
            if (worksheet.headers.length > 1) {
                // Delete column definitions
                if (columnNumber == null) {
                    const number = worksheet.getSelectedColumns(true);

                    if (!number.length) {
                        // Remove last column
                        columnNumber = worksheet.headers.length - 1;
                        numOfColumns = 1;
                    } else {
                        // Remove selected
                        columnNumber = parseInt(number[0]);
                        numOfColumns = parseInt(number.length);
                    }
                }

                // Lasat column
                const lastColumn = worksheet.options.data[0].length - 1;

                if (columnNumber == null || columnNumber > lastColumn || columnNumber < 0) {
                    columnNumber = lastColumn;
                }

                // Minimum of columns to be delete is 1
                if (!numOfColumns) {
                    numOfColumns = 1;
                }

                // Can't delete more than the limit of the table
                if (numOfColumns > worksheet.options.data[0].length - columnNumber) {
                    numOfColumns = worksheet.options.data[0].length - columnNumber;
                }

                const removedColumns = [];
                for (let i = 0; i < numOfColumns; i++) {
                    removedColumns.push(i + columnNumber);
                }

                // onbeforedeletecolumn
                if (worksheet.dispatch('onbeforedeletecolumn', worksheet, removedColumns) === false) {
                    return false;
                }

                // Can't remove the last column
                if (parseInt(columnNumber) > -1) {
                    // Merged cells
                    let mergeExists = false;
                    if (worksheet.options.mergeCells && Object.keys(worksheet.options.mergeCells).length > 0) {
                        for (let col = columnNumber; col < columnNumber + numOfColumns; col++) {
                            if (worksheet.merges.isColMerged(col, null).length) {
                                mergeExists = true;
                            }
                        }
                    }
                    if (mergeExists) {
                        if (!askConfirm(worksheet.parent.config, translate('This action will destroy any existing merged cells. Are you sure?'))) {
                            return false;
                        } else {
                            worksheet.destroyMerge();
                        }
                    }

                    // Delete the column properties
                    const columns = worksheet.options.columns ? worksheet.options.columns.splice(columnNumber, numOfColumns) : undefined;

                    for (let col = columnNumber; col < columnNumber + numOfColumns; col++) {
                        worksheet.cols[col].colElement.className = '';
                        worksheet.headers[col].className = '';
                        worksheet.cols[col].colElement.parentNode!.removeChild(worksheet.cols[col].colElement);
                        worksheet.headers[col].parentNode!.removeChild(worksheet.headers[col]);
                    }

                    const historyHeaders = worksheet.headers.splice(columnNumber, numOfColumns);
                    const historyColgroup = worksheet.cols.splice(columnNumber, numOfColumns);
                    const historyRecords = [];
                    const historyData = [];
                    const historyFooters = [];

                    for (let row = 0; row < worksheet.options.data.length; row++) {
                        for (let col = columnNumber; col < columnNumber + numOfColumns; col++) {
                            worksheet.records[row][col].element.className = '';
                            worksheet.records[row][col].element.parentNode!.removeChild(worksheet.records[row][col].element);
                        }
                    }

                    // Delete headers
                    for (let row = 0; row < worksheet.options.data.length; row++) {
                        // History
                        historyData[row] = worksheet.options.data[row].splice(columnNumber, numOfColumns);
                        historyRecords[row] = worksheet.records[row].splice(columnNumber, numOfColumns);
                    }

                    for (let i = columnNumber; i < worksheet.cols.length; i++) {
                        worksheet.cols[i].x = i;
                    }

                    for (const record of worksheet.records) {
                        for (let i = columnNumber; i < record.length; i++) {
                            record[i].x = i;
                        }
                    }

                    // Delete footers
                    if (worksheet.options.footers) {
                        for (let row = 0; row < worksheet.options.footers.length; row++) {
                            historyFooters[row] = worksheet.options.footers[row].splice(columnNumber, numOfColumns);
                        }
                    }

                    // Remove selection
                    worksheet.selection.conditionalUpdate(0, columnNumber, columnNumber + numOfColumns - 1);

                    // Adjust nested headers
                    if (worksheet.options.nestedHeaders && worksheet.options.nestedHeaders.length > 0 && worksheet.options.nestedHeaders[0]?.[0]) {
                        for (let j = 0; j < worksheet.options.nestedHeaders.length; j++) {
                            const colspan = parseInt(worksheet.options.nestedHeaders[j][worksheet.options.nestedHeaders[j].length - 1].colspan) - numOfColumns;
                            worksheet.options.nestedHeaders[j][worksheet.options.nestedHeaders[j].length - 1].colspan = colspan;
                            worksheet.thead.children[j].children[worksheet.thead.children[j].children.length - 1].setAttribute('colspan', colspan);
                        }
                    }

                    // Keeping history of changes
                    worksheet.history.add({
                        action: 'deleteColumn',
                        columnNumber,
                        numOfColumns,
                        insertBefore: 1,
                        columns,
                        headers: historyHeaders,
                        cols: historyColgroup,
                        records: historyRecords,
                        footers: historyFooters,
                        data: historyData,
                    });

                    // Update table references
                    worksheet.formulas.updateTableReferences({ column: Number(columnNumber) });

                    // The frozen columns can be other columns now
                    this.updateFrozenColumns();

                    // Delete
                    worksheet.dispatch('ondeletecolumn', worksheet, removedColumns);
                }
            } else {
                console.error('Jspreadsheet: It is not possible to delete the last column');
            }
        }
    }

    /**
     * Get the width of a column.
     * @param column - The column, from 0. Without it, the widths of all columns.
     */
    getWidth(column?: number): number | (number | string)[] {
        const worksheet = this.worksheet;

        let data;

        if (typeof column === 'undefined') {
            // Get all headers
            data = [];
            for (let i = 0; i < worksheet.headers.length; i++) {
                data.push(worksheet.options.columns?.[i]?.width || worksheet.options.defaultColWidth || 100);
            }
        } else {
            data = parseInt(worksheet.cols[column].colElement.getAttribute('width'));
        }

        return data;
    }

    /**
     * Set the width of columns.
     * @param column - The column, from 0, or the columns.
     * @param width - The new width, or the widths of the columns.
     * @param oldWidth - The width before the change, for the history. Read from the column by default.
     */
    setWidth(column: number | number[], width: number | number[], oldWidth?: number | number[]) {
        const worksheet = this.worksheet;

        if (width) {
            if (Array.isArray(column)) {
                // Oldwidth
                if (!oldWidth) {
                    oldWidth = [];
                }
                // Set width. An array of columns has an array of old widths.
                for (let i = 0; i < column.length; i++) {
                    if (!(oldWidth as number[])[i]) {
                        (oldWidth as number[])[i] = parseInt(worksheet.cols[column[i]].colElement.getAttribute('width'));
                    }
                    // A column without a width in the array of widths keeps its width
                    const w = Array.isArray(width) ? width[i] : width;
                    if (!w) {
                        continue;
                    }
                    worksheet.cols[column[i]].colElement.setAttribute('width', w);

                    if (!worksheet.options.columns) {
                        worksheet.options.columns = [];
                    }

                    if (!worksheet.options.columns[column[i]]) {
                        worksheet.options.columns[column[i]] = {};
                    }

                    worksheet.options.columns[column[i]].width = w;
                }
            } else {
                // Oldwidth
                if (!oldWidth) {
                    oldWidth = parseInt(worksheet.cols[column].colElement.getAttribute('width'));
                }
                // Set width. One column has one width.
                worksheet.cols[column].colElement.setAttribute('width', width as number);

                if (!worksheet.options.columns) {
                    worksheet.options.columns = [];
                }

                if (!worksheet.options.columns[column]) {
                    worksheet.options.columns[column] = {};
                }

                worksheet.options.columns[column].width = width as number;
            }

            // Keeping history of changes
            worksheet.history.add({
                action: 'setWidth',
                column,
                oldValue: oldWidth,
                newValue: width,
            });

            // The column window reads the widths again
            worksheet.rowSource?.columns?.refresh();

            // On resize column
            worksheet.dispatch('onresizecolumn', worksheet, column, width, oldWidth);

            // Update corner position
            worksheet.selection.updateCorner();
        }
    }

    /**
     * Show column
     */
    show(colNumber: number | number[]) {
        const worksheet = this.worksheet;

        if (!Array.isArray(colNumber)) {
            colNumber = [colNumber];
        }

        for (const columnIndex of colNumber) {
            worksheet.headers[columnIndex].style.display = '';
            worksheet.cols[columnIndex].colElement.style.display = '';
            if (worksheet.filter && worksheet.filter.children.length > columnIndex + 1) {
                (worksheet.filter.children[columnIndex + 1] as HTMLElement).style.display = '';
            }
            for (const j of this.rowIndexes()) {
                // A column outside the column window has no cells
                const record = worksheet.records[j][columnIndex];
                if (record) {
                    record.element.style.display = '';
                }
            }
        }

        worksheet.merges.updateColumns(colNumber);

        // The column window reads the widths again
        worksheet.rowSource?.columns?.refresh();

        // Update footers
        if (worksheet.options.footers) {
            worksheet.formulas.setFooter();
        }

        // The line of the frozen pane moves when a frozen column is shown or hidden
        if (colNumber.some((columnIndex) => columnIndex < worksheet.options.freezeColumns!)) {
            this.updateFreezePosition();
        }

        worksheet.resetSelection();
    }

    /**
     * Hide column
     */
    hide(colNumber: number | number[]) {
        const worksheet = this.worksheet;

        if (!Array.isArray(colNumber)) {
            colNumber = [colNumber];
        }

        for (const columnIndex of colNumber) {
            worksheet.headers[columnIndex].style.display = 'none';
            worksheet.cols[columnIndex].colElement.style.display = 'none';
            if (worksheet.filter && worksheet.filter.children.length > columnIndex + 1) {
                (worksheet.filter.children[columnIndex + 1] as HTMLElement).style.display = 'none';
            }
            for (const j of this.rowIndexes()) {
                // A column outside the column window has no cells
                const record = worksheet.records[j][columnIndex];
                if (record) {
                    record.element.style.display = 'none';
                }
            }
        }

        worksheet.merges.updateColumns(colNumber);

        // The column window reads the widths again
        worksheet.rowSource?.columns?.refresh();

        // Update footers
        if (worksheet.options.footers) {
            worksheet.formulas.setFooter();
        }

        // The line of the frozen pane moves when a frozen column is shown or hidden
        if (colNumber.some((columnIndex) => columnIndex < worksheet.options.freezeColumns!)) {
            this.updateFreezePosition();
        }

        worksheet.resetSelection();

        // The tab stop of the grid can be in a hidden column
        worksheet.selection.updateTabStop();
    }

    /**
     * Get a column data by columnNumber
     */
    getData(columnNumber: number, processed?: boolean): jspreadsheet.CellValue[] {
        const worksheet = this.worksheet;

        const dataset = [];
        // Go through the rows to get the data
        for (const j of this.rowIndexes()) {
            // A column outside the column window has no cells, and is empty
            if (!worksheet.records[j][columnNumber]) {
                dataset.push('');
            } else if (processed) {
                dataset.push(worksheet.records[j][columnNumber].element.innerHTML);
            } else {
                dataset.push(worksheet.options.data[j][columnNumber]);
            }
        }
        return dataset;
    }

    /**
     * Set a column data by colNumber
     */
    setData(colNumber: number, data: (jspreadsheet.CellValue | null)[], force?: boolean) {
        const worksheet = this.worksheet;

        for (let j = 0; j < worksheet.rows.length; j++) {
            // Update cell
            const columnName = getColumnNameFromId([colNumber, j]);
            // Set value
            if (data[j] != null) {
                worksheet.setValue(columnName, data[j]!, force);
            }
        }
    }

    /**
     * Get the column title
     * @param column - column number (first column is: 0)
     */
    getHeader(column: number): string {
        const worksheet = this.worksheet;

        return worksheet.headers[column].textContent;
    }

    /**
     * Get the headers
     * @param asArray - Return an array instead of a text with the CSV delimiter.
     */
    getHeaders(asArray?: boolean): string | string[] {
        const worksheet = this.worksheet;

        const title = [];

        for (let i = 0; i < worksheet.headers.length; i++) {
            title.push(worksheet.getHeader(i));
        }

        return asArray ? title : title.join(worksheet.options.csvDelimiter);
    }

    /**
     * Set the column title
     * @param column - column number (first column is: 0)
     * @param newValue - new column title. Without it, the title is the column name, for example "B".
     */
    setHeader(column: number, newValue?: string) {
        const worksheet = this.worksheet;

        if (worksheet.headers[column]) {
            const oldValue = worksheet.headers[column].textContent;
            const onchangeheaderOldValue = worksheet.options.columns?.[column]?.title || '';

            if (!newValue) {
                newValue = getColumnName(column);
            }

            worksheet.headers[column].textContent = newValue;
            // Keep the title property
            worksheet.headers[column].setAttribute('title', newValue);
            // Update title
            if (!worksheet.options.columns) {
                worksheet.options.columns = [];
            }
            if (!worksheet.options.columns[column]) {
                worksheet.options.columns[column] = {};
            }
            worksheet.options.columns[column].title = newValue;

            worksheet.history.add({
                action: 'setHeader',
                column,
                oldValue,
                newValue,
            });

            // On onchange header
            worksheet.dispatch('onchangeheader', worksheet, parseInt(column), newValue, onchangeheaderOldValue);
        }
    }

    /**
     * Ask the user for the title of a column, and set it. A cancelled prompt keeps the title. An empty title is the
     * column name, as with setHeader.
     * @param column - The column, from 0.
     */
    rename(column: number) {
        const worksheet = this.worksheet;

        const newValue = prompt(translate('Column name'), worksheet.getHeader(column));

        if (newValue !== null) {
            worksheet.setHeader(column, newValue);
        }
    }

    /**
     * Create a nested header object
     */
    // The row of nested headers keeps its element
    createNestedHeader(nestedInformation: jspreadsheet.NestedHeaderCell[] & { element?: HTMLTableRowElement }) {
        const worksheet = this.worksheet;

        const tr = document.createElement('tr');
        tr.classList.add('jss_nested');
        const td = document.createElement('td');
        td.classList.add('jss_selectall');

        tr.appendChild(td);
        // Element
        nestedInformation.element = tr;

        let headerIndex = 0;
        for (const header of nestedInformation) {
            // Default values
            if (!header.colspan) {
                header.colspan = 1;
            }
            if (!header.title) {
                header.title = '';
            }
            if (!header.id) {
                header.id = '';
            }

            // Number of columns. colspan, title and id have default values above.
            let numberOfColumns = header.colspan!;

            // Classes container
            const column = [];
            // Header classes for this cell
            for (let x = 0; x < numberOfColumns; x++) {
                if (worksheet.options.columns[headerIndex]?.type === 'hidden') {
                    numberOfColumns++;
                }
                column.push(headerIndex);
                headerIndex++;
            }

            // Created the nested cell
            const td = document.createElement('td');
            td.setAttribute('data-column', column.join(','));
            td.setAttribute('colspan', header.colspan!);
            td.setAttribute('align', header.align || 'center');
            td.setAttribute('id', header.id!);
            td.textContent = header.title!;
            tr.appendChild(td);
        }

        return tr;
    }

    /** The rows with a record: all rows, or the frozen rows and the window of the row source mode. */
    private rowIndexes() {
        return recordRows(this.worksheet);
    }

    // Get the width of the index column. The frozen columns stay at the left of the content when scrollLeft is more than this width.
    getIndexWidth() {
        return this.worksheet.table.classList.contains('jss_hidden_index') ? 0 : 50;
    }

    // Get width of all freezed cells together
    getFreezeWidth() {
        const worksheet = this.worksheet;

        let width = 0;
        if (worksheet.options.freezeColumns! > 0) {
            const freezeColumns = Math.min(worksheet.options.freezeColumns!, worksheet.headers.length);
            for (let i = 0; i < freezeColumns; i++) {
                let columnWidth;
                if (worksheet.options.columns?.[i]?.width !== undefined) {
                    columnWidth = parseInt(worksheet.options.columns[i].width);
                } else {
                    columnWidth = worksheet.options.defaultColWidth !== undefined ? parseInt(worksheet.options.defaultColWidth) : 100;
                }

                width += columnWidth;
            }
        }
        return width;
    }

    /**
     * Place the frozen columns at the left of the content after a scroll or a change of the columns.
     */
    updateFreezePosition() {
        const worksheet = this.worksheet;

        const scrollLeft = worksheet.content.scrollLeft;
        const indexWidth = this.getIndexWidth();
        // The worksheet can have fewer columns than freezeColumns, for example after deleteColumn
        const freezeColumns = Math.min(worksheet.options.freezeColumns!, worksheet.headers.length);
        // The rows with a record, once for all frozen columns
        const rows = freezeColumns > 0 ? this.rowIndexes() : [];
        let width = 0;
        // The scroll position is in screen pixels, and the widths and the left of a cell in grid pixels
        if (scrollLeft > toScreen(worksheet, indexWidth)) {
            // The last frozen column that is not hidden shows the line of the frozen pane
            let last = freezeColumns - 1;
            while (last > 0 && worksheet.headers[last].style.display === 'none') {
                last--;
            }

            for (let i = 0; i < freezeColumns; i++) {
                if (i > 0) {
                    // Must check if the previous column is hidden or not to determin whether the width shoule be added or not!
                    if (worksheet.options.columns?.[i - 1]?.type !== 'hidden') {
                        let columnWidth;
                        if (worksheet.options.columns?.[i - 1]?.width !== undefined) {
                            columnWidth = parseInt(worksheet.options.columns[i - 1].width);
                        } else {
                            columnWidth = worksheet.options.defaultColWidth !== undefined ? parseInt(worksheet.options.defaultColWidth) : 100;
                        }

                        width += parseInt(columnWidth);
                    }
                }
                // Only the cells of the last column, or of a column that was the last before, get a change of the line
                const isLast = i === last;
                const wasLast = worksheet.headers[i].classList.contains('jss_freezed_last');
                worksheet.headers[i].classList.add('jss_freezed');
                worksheet.headers[i].classList.toggle('jss_freezed_last', isLast);
                worksheet.headers[i].style.left = `${width}px`;
                // The cells are relative, so all the frozen cells move by the same distance as the scroll. A cell of a
                // frozen row is sticky, as a header is, so it gets the same distance from the left as the header.
                const shifted = `${toGrid(worksheet, scrollLeft) - indexWidth - 1}px`;
                const frozenRows = worksheet.rowManager.getFreezeCount();
                for (const j of rows) {
                    if (worksheet.rows[j] && worksheet.records[j][i]) {
                        worksheet.records[j][i].element.classList.add('jss_freezed');
                        if (isLast || wasLast) {
                            worksheet.records[j][i].element.classList.toggle('jss_freezed_last', isLast);
                        }
                        worksheet.records[j][i].element.style.left = j < frozenRows ? `${width}px` : shifted;
                    }
                }
            }
        } else {
            for (let i = 0; i < freezeColumns; i++) {
                this.unfreeze(i);
            }
        }

        // A column that an insert, a move or a delete took out of the frozen columns is not frozen any more. Its header
        // and its cells moved together, so the header tells.
        for (let i = freezeColumns; i < worksheet.headers.length; i++) {
            if (worksheet.headers[i].classList.contains('jss_freezed')) {
                this.unfreeze(i);
            }
        }

        // Place the corner in the correct place
        worksheet.selection.updateCorner();
    }

    /**
     * Place a column back in the table, where the frozen columns are not.
     * @param i - The column.
     */
    private unfreeze(i: number) {
        const worksheet = this.worksheet;

        worksheet.headers[i].classList.remove('jss_freezed', 'jss_freezed_last');
        worksheet.headers[i].style.left = '';
        for (const j of this.rowIndexes()) {
            if (worksheet.records[j][i]) {
                worksheet.records[j][i].element.classList.remove('jss_freezed', 'jss_freezed_last');
                worksheet.records[j][i].element.style.left = '';
            }
        }
    }

    /**
     * Place the frozen columns again after the columns changed, when the worksheet has frozen columns.
     */
    updateFrozenColumns() {
        if (this.worksheet.options.freezeColumns! > 0) {
            this.updateFreezePosition();
        }
    }
}
