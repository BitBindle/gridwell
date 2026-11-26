import jSuites from 'jsuites';
import dispatch from '../utils/dispatch.js';
import { getColumnName } from '../utils/helpers.js';
import { getColumnNameFromId, injectArray } from '../utils/internalHelpers.js';
import { setHeaderRole } from '../utils/accessibility.js';
import type jspreadsheet from '../jspreadsheet';
import type { Worksheet } from '../types/internal';

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
        const obj = this.worksheet;

        let numberOfColumns = (obj.options.columns && obj.options.columns.length) || 0;

        if (obj.options.data && typeof obj.options.data[0] !== 'undefined') {
            // Data keys
            const keys = Object.keys(obj.options.data[0]);

            if (keys.length > numberOfColumns) {
                numberOfColumns = keys.length;
            }
        }

        if (obj.options.minDimensions && obj.options.minDimensions[0] > numberOfColumns) {
            numberOfColumns = obj.options.minDimensions[0];
        }

        return numberOfColumns;
    }

    createHeader(colNumber: number) {
        const obj = this.worksheet;

        // Create col global control
        const colWidth = (obj.options.columns && obj.options.columns[colNumber] && obj.options.columns[colNumber].width) || obj.options.defaultColWidth || 100;
        const colAlign =
            (obj.options.columns && obj.options.columns[colNumber] && obj.options.columns[colNumber].align) || obj.options.defaultColAlign || 'center';

        // Create header cell
        obj.headers[colNumber] = document.createElement('td');
        obj.headers[colNumber].textContent =
            (obj.options.columns && obj.options.columns[colNumber] && obj.options.columns[colNumber].title) || getColumnName(colNumber);

        obj.headers[colNumber].setAttribute('data-x', colNumber);
        setHeaderRole(obj.headers[colNumber], colNumber);
        obj.headers[colNumber].style.textAlign = colAlign;
        if (obj.options.columns && obj.options.columns[colNumber] && obj.options.columns[colNumber].title) {
            obj.headers[colNumber].setAttribute('title', obj.headers[colNumber].innerText);
        }
        // id is not in the public typings of a column
        if (obj.options.columns && obj.options.columns[colNumber] && obj.options.columns[colNumber].id) {
            obj.headers[colNumber].setAttribute('id', obj.options.columns[colNumber].id);
        }

        // Width control
        const colElement = document.createElement('col');
        colElement.setAttribute('width', colWidth);

        obj.cols[colNumber] = {
            colElement,
            x: colNumber,
        };

        // Hidden column
        if (obj.options.columns && obj.options.columns[colNumber] && obj.options.columns[colNumber].type == 'hidden') {
            obj.headers[colNumber].style.display = 'none';
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
        const obj = this.worksheet;

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

            for (let i = 0; i < obj.options.data.length; i++) {
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
        const obj = this.worksheet;

        const historyHeaders = [];
        const historyColgroup = [];

        for (let col = columnIndex; col < numOfColumns + columnIndex; col++) {
            this.createHeader(col);
            obj.headerContainer.insertBefore(obj.headers[col], obj.headerContainer.children[col + 1]);
            obj.colgroupContainer.insertBefore(obj.cols[col].colElement, obj.colgroupContainer.children[col + 1]);

            historyHeaders.push(obj.headers[col]);
            historyColgroup.push(obj.cols[col]);
        }

        return { historyHeaders, historyColgroup };
    }

    /**
     * Add an empty cell for each new column to the footers.
     */
    private addFooters(columnIndex: number, numOfColumns: number) {
        const obj = this.worksheet;

        const historyFooters: string[][] = [];

        if (obj.options.footers) {
            for (let j = 0; j < obj.options.footers.length; j++) {
                historyFooters[j] = [];
                for (let i = 0; i < numOfColumns; i++) {
                    historyFooters[j].push('');
                }
                obj.options.footers[j].splice(columnIndex, 0, ...historyFooters[j]);
            }
        }

        return historyFooters;
    }

    /**
     * Create the cells of the new columns in each row, with the data or empty.
     */
    private addCells(columnIndex: number, numOfColumns: number, data: jspreadsheet.CellValue[]) {
        const obj = this.worksheet;

        const historyRecords: { element: HTMLTableCellElement; x: number; y: number }[][] = [];
        const historyData: jspreadsheet.CellValue[][] = [];

        for (let row = 0; row < obj.options.data.length; row++) {
            // Keep the current data
            const currentData = obj.options.data[row].splice(columnIndex);
            const currentRecord = obj.records[row].splice(columnIndex);

            // History
            historyData[row] = [];
            historyRecords[row] = [];

            for (let col = columnIndex; col < numOfColumns + columnIndex; col++) {
                // New value
                const value = data[row] ? data[row] : '';
                obj.options.data[row][col] = value;
                // New cell
                const td = obj.cells.create(col, row, obj.options.data[row][col]);
                // updateTableReferences adds x
                obj.records[row][col] = {
                    element: td,
                    y: row,
                } as Worksheet['records'][number][number];
                // Add cell to the row
                if (obj.rows[row]) {
                    obj.rows[row].element.insertBefore(td, obj.rows[row].element.children[col + 1]);
                }

                if (obj.options.columns && obj.options.columns[col] && typeof obj.options.columns[col].render === 'function') {
                    // The test above makes it a function, which TypeScript does not narrow through the index
                    (obj.options.columns[col].render as NonNullable<jspreadsheet.BaseColumn['render']>)(
                        td,
                        value,
                        parseInt(col),
                        parseInt(row),
                        obj,
                        obj.options.columns[col]
                    );
                }

                // Record History
                historyData[row].push(value);
                historyRecords[row].push({ element: td, x: col, y: row });
            }

            // Copy the data back to the main data
            Array.prototype.push.apply(obj.options.data[row], currentData);
            Array.prototype.push.apply(obj.records[row], currentRecord);
        }

        return { historyRecords, historyData };
    }

    /**
     * Widen the last nested header of each row over the new columns.
     */
    private widenNestedHeaders(columnIndex: number, numOfColumns: number) {
        const obj = this.worksheet;

        if (obj.options.nestedHeaders && obj.options.nestedHeaders.length > 0 && obj.options.nestedHeaders[0] && obj.options.nestedHeaders[0][0]) {
            for (let j = 0; j < obj.options.nestedHeaders.length; j++) {
                const colspan = parseInt(obj.options.nestedHeaders[j][obj.options.nestedHeaders[j].length - 1].colspan) + numOfColumns;
                obj.options.nestedHeaders[j][obj.options.nestedHeaders[j].length - 1].colspan = colspan;
                obj.thead.children[j].children[obj.thead.children[j].children.length - 1].setAttribute('colspan', colspan);
                const o: (string | number)[] = obj.thead.children[j].children[obj.thead.children[j].children.length - 1]
                    .getAttribute('data-column')!
                    .split(',');
                for (let col = columnIndex; col < numOfColumns + columnIndex; col++) {
                    o.push(col);
                }
                obj.thead.children[j].children[obj.thead.children[j].children.length - 1].setAttribute('data-column', o.join(','));
            }
        }
    }

    /**
     * Insert a new column
     *
     * @param mixed - num of columns to be added or data to be added in one single column
     * @param int columnNumber - number of columns to be created
     * @param bool insertBefore
     * @param object properties - column properties
     * @return void
     */
    insert(
        mixed?: number | jspreadsheet.CellValue[],
        columnNumber?: number,
        insertBefore?: boolean | number,
        properties?: jspreadsheet.Column[]
    ): false | undefined {
        const obj = this.worksheet;

        // Configuration
        if (obj.options.allowInsertColumn != false) {
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
                obj.options.columns.length,
                ...obj.options.data.map(function (row) {
                    return row.length;
                })
            );

            const lastColumn = currentNumOfColumns - 1;

            // Confirm position
            if (columnNumber == undefined || columnNumber >= parseInt(lastColumn) || columnNumber < 0) {
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
            if (dispatch.call(obj, 'onbeforeinsertcolumn', obj, columns) === false) {
                return false;
            }

            // Merged cells
            if (obj.options.mergeCells && Object.keys(obj.options.mergeCells).length > 0) {
                if (obj.merges.isColMerged(columnNumber, insertBefore).length) {
                    if (!confirm(jSuites.translate('This action will destroy any existing merged cells. Are you sure?'))) {
                        return false;
                    } else {
                        obj.destroyMerge();
                    }
                }
            }

            // Insert before
            const columnIndex = !insertBefore ? columnNumber + 1 : columnNumber;
            obj.options.columns = injectArray(obj.options.columns, columnIndex, properties);

            // Open space in the containers
            const currentHeaders = obj.headers.splice(columnIndex);
            const currentColgroup = obj.cols.splice(columnIndex);

            // Add new headers
            const { historyHeaders, historyColgroup } = this.addHeaders(columnIndex, numOfColumns);

            // Add new footer cells
            const historyFooters = this.addFooters(columnIndex, numOfColumns);

            // Adding visual columns
            const { historyRecords, historyData } = this.addCells(columnIndex, numOfColumns, data);

            Array.prototype.push.apply(obj.headers, currentHeaders);
            Array.prototype.push.apply(obj.cols, currentColgroup);

            for (let i = columnIndex; i < obj.cols.length; i++) {
                obj.cols[i].x = i;
            }

            for (let j = 0; j < obj.records.length; j++) {
                for (let i = 0; i < obj.records[j].length; i++) {
                    obj.records[j][i].x = i;
                }
            }

            // Adjust nested headers
            this.widenNestedHeaders(columnIndex, numOfColumns);

            // Keep history
            obj.history.add({
                action: 'insertColumn',
                columnNumber: columnNumber,
                numOfColumns: numOfColumns,
                insertBefore: insertBefore,
                columns: properties,
                headers: historyHeaders,
                cols: historyColgroup,
                records: historyRecords,
                footers: historyFooters,
                data: historyData,
            });

            // Remove table references
            obj.formulas.updateTableReferences();

            // Events
            dispatch.call(obj, 'oninsertcolumn', obj, columns);
        }
    }

    /**
     * Move column
     *
     * @return void
     */
    move(o: number, d: number): false | undefined {
        const obj = this.worksheet;

        if (obj.options.mergeCells && Object.keys(obj.options.mergeCells).length > 0) {
            let insertBefore: 0 | 1;
            if (o > d) {
                insertBefore = 1;
            } else {
                insertBefore = 0;
            }

            if (obj.merges.isColMerged(o).length || obj.merges.isColMerged(d, insertBefore).length) {
                if (!confirm(jSuites.translate('This action will destroy any existing merged cells. Are you sure?'))) {
                    return false;
                } else {
                    obj.destroyMerge();
                }
            }
        }

        o = parseInt(o);
        d = parseInt(d);

        if (o > d) {
            obj.headerContainer.insertBefore(obj.headers[o], obj.headers[d]);
            obj.colgroupContainer.insertBefore(obj.cols[o].colElement, obj.cols[d].colElement);

            for (let j = 0; j < obj.rows.length; j++) {
                obj.rows[j].element.insertBefore(obj.records[j][o].element, obj.records[j][d].element);
            }
        } else {
            obj.headerContainer.insertBefore(obj.headers[o], obj.headers[d].nextSibling);
            obj.colgroupContainer.insertBefore(obj.cols[o].colElement, obj.cols[d].colElement.nextSibling);

            for (let j = 0; j < obj.rows.length; j++) {
                obj.rows[j].element.insertBefore(obj.records[j][o].element, obj.records[j][d].element.nextSibling);
            }
        }

        obj.options.columns.splice(d, 0, obj.options.columns.splice(o, 1)[0]);
        obj.headers.splice(d, 0, obj.headers.splice(o, 1)[0]);
        obj.cols.splice(d, 0, obj.cols.splice(o, 1)[0]);

        const firstAffectedIndex = Math.min(o, d);
        const lastAffectedIndex = Math.max(o, d);

        for (let j = 0; j < obj.rows.length; j++) {
            obj.options.data[j].splice(d, 0, obj.options.data[j].splice(o, 1)[0]);
            obj.records[j].splice(d, 0, obj.records[j].splice(o, 1)[0]);
        }

        for (let i = firstAffectedIndex; i <= lastAffectedIndex; i++) {
            obj.cols[i].x = i;
        }

        for (let j = 0; j < obj.records.length; j++) {
            for (let i = firstAffectedIndex; i <= lastAffectedIndex; i++) {
                obj.records[j][i].x = i;
            }
        }

        // Update footers position
        if (obj.options.footers) {
            for (let j = 0; j < obj.options.footers.length; j++) {
                obj.options.footers[j].splice(d, 0, obj.options.footers[j].splice(o, 1)[0]);
            }
        }

        // Keeping history of changes
        obj.history.add({
            action: 'moveColumn',
            oldValue: o,
            newValue: d,
        });

        // Update table references
        obj.formulas.updateTableReferences();

        // Events
        dispatch.call(obj, 'onmovecolumn', obj, o, d, 1);
    }

    /**
     * Delete a column by number
     *
     * @param integer columnNumber - reference column to be excluded
     * @param integer numOfColumns - number of columns to be excluded from the reference column
     * @return void
     */
    delete(columnNumber?: number, numOfColumns?: number): false | undefined {
        const obj = this.worksheet;

        // Global Configuration
        if (obj.options.allowDeleteColumn != false) {
            if (obj.headers.length > 1) {
                // Delete column definitions
                if (columnNumber == undefined) {
                    const number = obj.getSelectedColumns(true);

                    if (!number.length) {
                        // Remove last column
                        columnNumber = obj.headers.length - 1;
                        numOfColumns = 1;
                    } else {
                        // Remove selected
                        columnNumber = parseInt(number[0]);
                        numOfColumns = parseInt(number.length);
                    }
                }

                // Lasat column
                const lastColumn = obj.options.data[0].length - 1;

                if (columnNumber == undefined || columnNumber > lastColumn || columnNumber < 0) {
                    columnNumber = lastColumn;
                }

                // Minimum of columns to be delete is 1
                if (!numOfColumns) {
                    numOfColumns = 1;
                }

                // Can't delete more than the limit of the table
                if (numOfColumns > obj.options.data[0].length - columnNumber) {
                    numOfColumns = obj.options.data[0].length - columnNumber;
                }

                const removedColumns = [];
                for (let i = 0; i < numOfColumns; i++) {
                    removedColumns.push(i + columnNumber);
                }

                // onbeforedeletecolumn
                if (dispatch.call(obj, 'onbeforedeletecolumn', obj, removedColumns) === false) {
                    return false;
                }

                // Can't remove the last column
                if (parseInt(columnNumber) > -1) {
                    // Merged cells
                    let mergeExists = false;
                    if (obj.options.mergeCells && Object.keys(obj.options.mergeCells).length > 0) {
                        for (let col = columnNumber; col < columnNumber + numOfColumns; col++) {
                            if (obj.merges.isColMerged(col, null).length) {
                                mergeExists = true;
                            }
                        }
                    }
                    if (mergeExists) {
                        if (!confirm(jSuites.translate('This action will destroy any existing merged cells. Are you sure?'))) {
                            return false;
                        } else {
                            obj.destroyMerge();
                        }
                    }

                    // Delete the column properties
                    const columns = obj.options.columns ? obj.options.columns.splice(columnNumber, numOfColumns) : undefined;

                    for (let col = columnNumber; col < columnNumber + numOfColumns; col++) {
                        obj.cols[col].colElement.className = '';
                        obj.headers[col].className = '';
                        obj.cols[col].colElement.parentNode!.removeChild(obj.cols[col].colElement);
                        obj.headers[col].parentNode!.removeChild(obj.headers[col]);
                    }

                    const historyHeaders = obj.headers.splice(columnNumber, numOfColumns);
                    const historyColgroup = obj.cols.splice(columnNumber, numOfColumns);
                    const historyRecords = [];
                    const historyData = [];
                    const historyFooters = [];

                    for (let row = 0; row < obj.options.data.length; row++) {
                        for (let col = columnNumber; col < columnNumber + numOfColumns; col++) {
                            obj.records[row][col].element.className = '';
                            obj.records[row][col].element.parentNode!.removeChild(obj.records[row][col].element);
                        }
                    }

                    // Delete headers
                    for (let row = 0; row < obj.options.data.length; row++) {
                        // History
                        historyData[row] = obj.options.data[row].splice(columnNumber, numOfColumns);
                        historyRecords[row] = obj.records[row].splice(columnNumber, numOfColumns);
                    }

                    for (let i = columnNumber; i < obj.cols.length; i++) {
                        obj.cols[i].x = i;
                    }

                    for (let j = 0; j < obj.records.length; j++) {
                        for (let i = columnNumber; i < obj.records[j].length; i++) {
                            obj.records[j][i].x = i;
                        }
                    }

                    // Delete footers
                    if (obj.options.footers) {
                        for (let row = 0; row < obj.options.footers.length; row++) {
                            historyFooters[row] = obj.options.footers[row].splice(columnNumber, numOfColumns);
                        }
                    }

                    // Remove selection
                    obj.selection.conditionalUpdate(0, columnNumber, columnNumber + numOfColumns - 1);

                    // Adjust nested headers
                    if (obj.options.nestedHeaders && obj.options.nestedHeaders.length > 0 && obj.options.nestedHeaders[0] && obj.options.nestedHeaders[0][0]) {
                        for (let j = 0; j < obj.options.nestedHeaders.length; j++) {
                            const colspan = parseInt(obj.options.nestedHeaders[j][obj.options.nestedHeaders[j].length - 1].colspan) - numOfColumns;
                            obj.options.nestedHeaders[j][obj.options.nestedHeaders[j].length - 1].colspan = colspan;
                            obj.thead.children[j].children[obj.thead.children[j].children.length - 1].setAttribute('colspan', colspan);
                        }
                    }

                    // Keeping history of changes
                    obj.history.add({
                        action: 'deleteColumn',
                        columnNumber: columnNumber,
                        numOfColumns: numOfColumns,
                        insertBefore: 1,
                        columns: columns,
                        headers: historyHeaders,
                        cols: historyColgroup,
                        records: historyRecords,
                        footers: historyFooters,
                        data: historyData,
                    });

                    // Update table references
                    obj.formulas.updateTableReferences();

                    // Delete
                    dispatch.call(obj, 'ondeletecolumn', obj, removedColumns);
                }
            } else {
                console.error('Jspreadsheet: It is not possible to delete the last column');
            }
        }
    }

    /**
     * Get the column width
     *
     * @param int column column number (first column is: 0)
     * @return int current width
     */
    getWidth(column?: number): number | (number | string)[] {
        const obj = this.worksheet;

        let data;

        if (typeof column === 'undefined') {
            // Get all headers
            data = [];
            for (let i = 0; i < obj.headers.length; i++) {
                data.push((obj.options.columns && obj.options.columns[i] && obj.options.columns[i].width) || obj.options.defaultColWidth || 100);
            }
        } else {
            data = parseInt(obj.cols[column].colElement.getAttribute('width'));
        }

        return data;
    }

    /**
     * Set the column width
     *
     * @param int column number (first column is: 0)
     * @param int new column width
     * @param int old column width
     */
    setWidth(column: number | number[], width: number | number[], oldWidth?: number | number[]) {
        const obj = this.worksheet;

        if (width) {
            if (Array.isArray(column)) {
                // Oldwidth
                if (!oldWidth) {
                    oldWidth = [];
                }
                // Set width. An array of columns has an array of old widths.
                for (let i = 0; i < column.length; i++) {
                    if (!(oldWidth as number[])[i]) {
                        (oldWidth as number[])[i] = parseInt(obj.cols[column[i]].colElement.getAttribute('width'));
                    }
                    // w is the whole array when the array of widths has no value for this column (issue #47)
                    const w = (Array.isArray(width) && width[i] ? width[i] : width) as number;
                    obj.cols[column[i]].colElement.setAttribute('width', w);

                    if (!obj.options.columns) {
                        obj.options.columns = [];
                    }

                    if (!obj.options.columns[column[i]]) {
                        obj.options.columns[column[i]] = {};
                    }

                    obj.options.columns[column[i]].width = w;
                }
            } else {
                // Oldwidth
                if (!oldWidth) {
                    oldWidth = parseInt(obj.cols[column].colElement.getAttribute('width'));
                }
                // Set width. One column has one width.
                obj.cols[column].colElement.setAttribute('width', width as number);

                if (!obj.options.columns) {
                    obj.options.columns = [];
                }

                if (!obj.options.columns[column]) {
                    obj.options.columns[column] = {};
                }

                obj.options.columns[column].width = width as number;
            }

            // Keeping history of changes
            obj.history.add({
                action: 'setWidth',
                column: column,
                oldValue: oldWidth,
                newValue: width,
            });

            // On resize column
            dispatch.call(obj, 'onresizecolumn', obj, column, width, oldWidth);

            // Update corner position
            obj.selection.updateCorner();
        }
    }

    /**
     * Show column
     */
    show(colNumber: number | number[]) {
        const obj = this.worksheet;

        if (!Array.isArray(colNumber)) {
            colNumber = [colNumber];
        }

        for (let i = 0; i < colNumber.length; i++) {
            const columnIndex = colNumber[i];

            obj.headers[columnIndex].style.display = '';
            obj.cols[columnIndex].colElement.style.display = '';
            if (obj.filter && obj.filter.children.length > columnIndex + 1) {
                (obj.filter.children[columnIndex + 1] as HTMLElement).style.display = '';
            }
            for (let j = 0; j < obj.options.data.length; j++) {
                obj.records[j][columnIndex].element.style.display = '';
            }
        }

        obj.merges.updateColumns(colNumber);

        // Update footers
        if (obj.options.footers) {
            obj.formulas.setFooter();
        }

        // The line of the frozen pane moves when a frozen column is shown or hidden
        if (colNumber.some((columnIndex) => columnIndex < obj.options.freezeColumns!)) {
            this.updateFreezePosition();
        }

        obj.resetSelection();
    }

    /**
     * Hide column
     */
    hide(colNumber: number | number[]) {
        const obj = this.worksheet;

        if (!Array.isArray(colNumber)) {
            colNumber = [colNumber];
        }

        for (let i = 0; i < colNumber.length; i++) {
            const columnIndex = colNumber[i];

            obj.headers[columnIndex].style.display = 'none';
            obj.cols[columnIndex].colElement.style.display = 'none';
            if (obj.filter && obj.filter.children.length > columnIndex + 1) {
                (obj.filter.children[columnIndex + 1] as HTMLElement).style.display = 'none';
            }
            for (let j = 0; j < obj.options.data.length; j++) {
                obj.records[j][columnIndex].element.style.display = 'none';
            }
        }

        obj.merges.updateColumns(colNumber);

        // Update footers
        if (obj.options.footers) {
            obj.formulas.setFooter();
        }

        // The line of the frozen pane moves when a frozen column is shown or hidden
        if (colNumber.some((columnIndex) => columnIndex < obj.options.freezeColumns!)) {
            this.updateFreezePosition();
        }

        obj.resetSelection();

        // The tab stop of the grid can be in a hidden column
        obj.selection.updateTabStop();
    }

    /**
     * Get a column data by columnNumber
     */
    getData(columnNumber: number, processed?: boolean): jspreadsheet.CellValue[] {
        const obj = this.worksheet;

        const dataset = [];
        // Go through the rows to get the data
        for (let j = 0; j < obj.options.data.length; j++) {
            if (processed) {
                dataset.push(obj.records[j][columnNumber].element.innerHTML);
            } else {
                dataset.push(obj.options.data[j][columnNumber]);
            }
        }
        return dataset;
    }

    /**
     * Set a column data by colNumber
     */
    setData(colNumber: number, data: (jspreadsheet.CellValue | null)[], force?: boolean) {
        const obj = this.worksheet;

        for (let j = 0; j < obj.rows.length; j++) {
            // Update cell
            const columnName = getColumnNameFromId([colNumber, j]);
            // Set value
            if (data[j] != null) {
                obj.setValue(columnName, data[j]!, force);
            }
        }
    }

    /**
     * Get the column title
     * @param column - column number (first column is: 0)
     */
    getHeader(column: number): string {
        const obj = this.worksheet;

        return obj.headers[column].textContent;
    }

    /**
     * Get the headers
     * @param asArray - Return an array instead of a text with the CSV delimiter.
     */
    getHeaders(asArray?: boolean): string | string[] {
        const obj = this.worksheet;

        const title = [];

        for (let i = 0; i < obj.headers.length; i++) {
            title.push(obj.getHeader(i));
        }

        return asArray ? title : title.join(obj.options.csvDelimiter);
    }

    /**
     * Set the column title
     * @param column - column number (first column is: 0)
     * @param newValue - new column title. Without it, the title is the column name, for example "B".
     */
    setHeader(column: number, newValue?: string) {
        const obj = this.worksheet;

        if (obj.headers[column]) {
            const oldValue = obj.headers[column].textContent;
            const onchangeheaderOldValue = (obj.options.columns && obj.options.columns[column] && obj.options.columns[column].title) || '';

            if (!newValue) {
                newValue = getColumnName(column);
            }

            obj.headers[column].textContent = newValue;
            // Keep the title property
            obj.headers[column].setAttribute('title', newValue);
            // Update title
            if (!obj.options.columns) {
                obj.options.columns = [];
            }
            if (!obj.options.columns[column]) {
                obj.options.columns[column] = {};
            }
            obj.options.columns[column].title = newValue;

            obj.history.add({
                action: 'setHeader',
                column: column,
                oldValue: oldValue,
                newValue: newValue,
            });

            // On onchange header
            dispatch.call(obj, 'onchangeheader', obj, parseInt(column), newValue, onchangeheaderOldValue);
        }
    }

    /**
     * Create a nested header object
     */
    // The row of nested headers keeps its element
    createNestedHeader(nestedInformation: jspreadsheet.NestedHeaderCell[] & { element?: HTMLTableRowElement }) {
        const obj = this.worksheet;

        const tr = document.createElement('tr');
        tr.classList.add('jss_nested');
        const td = document.createElement('td');
        td.classList.add('jss_selectall');

        tr.appendChild(td);
        // Element
        nestedInformation.element = tr;

        let headerIndex = 0;
        for (let i = 0; i < nestedInformation.length; i++) {
            // Default values
            if (!nestedInformation[i].colspan) {
                nestedInformation[i].colspan = 1;
            }
            if (!nestedInformation[i].title) {
                nestedInformation[i].title = '';
            }
            if (!nestedInformation[i].id) {
                nestedInformation[i].id = '';
            }

            // Number of columns. colspan, title and id have default values above.
            let numberOfColumns = nestedInformation[i].colspan!;

            // Classes container
            const column = [];
            // Header classes for this cell
            for (let x = 0; x < numberOfColumns; x++) {
                if (obj.options.columns[headerIndex] && obj.options.columns[headerIndex].type == 'hidden') {
                    numberOfColumns++;
                }
                column.push(headerIndex);
                headerIndex++;
            }

            // Created the nested cell
            const td = document.createElement('td');
            td.setAttribute('data-column', column.join(','));
            td.setAttribute('colspan', nestedInformation[i].colspan!);
            td.setAttribute('align', nestedInformation[i].align || 'center');
            td.setAttribute('id', nestedInformation[i].id!);
            td.textContent = nestedInformation[i].title!;
            tr.appendChild(td);
        }

        return tr;
    }

    // Get the width of the index column. The frozen columns stay at the left of the content when scrollLeft is more than this width.
    getIndexWidth() {
        return this.worksheet.table.classList.contains('jss_hidden_index') ? 0 : 50;
    }

    // Get width of all freezed cells together
    getFreezeWidth() {
        const obj = this.worksheet;

        let width = 0;
        if (obj.options.freezeColumns! > 0) {
            for (let i = 0; i < obj.options.freezeColumns!; i++) {
                let columnWidth;
                if (obj.options.columns && obj.options.columns[i] && obj.options.columns[i].width !== undefined) {
                    columnWidth = parseInt(obj.options.columns[i].width);
                } else {
                    columnWidth = obj.options.defaultColWidth !== undefined ? parseInt(obj.options.defaultColWidth) : 100;
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
        const obj = this.worksheet;

        const scrollLeft = obj.content.scrollLeft;
        const indexWidth = this.getIndexWidth();
        // The worksheet can have fewer columns than freezeColumns, for example after deleteColumn
        const freezeColumns = Math.min(obj.options.freezeColumns!, obj.headers.length);
        let width = 0;
        if (scrollLeft > indexWidth) {
            // The last frozen column that is not hidden shows the line of the frozen pane
            let last = freezeColumns - 1;
            while (last > 0 && obj.headers[last].style.display === 'none') {
                last--;
            }

            for (let i = 0; i < freezeColumns; i++) {
                if (i > 0) {
                    // Must check if the previous column is hidden or not to determin whether the width shoule be added or not!
                    if (!obj.options.columns || !obj.options.columns[i - 1] || obj.options.columns[i - 1].type !== 'hidden') {
                        let columnWidth;
                        if (obj.options.columns && obj.options.columns[i - 1] && obj.options.columns[i - 1].width !== undefined) {
                            columnWidth = parseInt(obj.options.columns[i - 1].width);
                        } else {
                            columnWidth = obj.options.defaultColWidth !== undefined ? parseInt(obj.options.defaultColWidth) : 100;
                        }

                        width += parseInt(columnWidth);
                    }
                }
                // Only the cells of the last column, or of a column that was the last before, get a change of the line
                const isLast = i === last;
                const wasLast = obj.headers[i].classList.contains('jss_freezed_last');
                obj.headers[i].classList.add('jss_freezed');
                obj.headers[i].classList.toggle('jss_freezed_last', isLast);
                obj.headers[i].style.left = width + 'px';
                for (let j = 0; j < obj.rows.length; j++) {
                    if (obj.rows[j] && obj.records[j][i]) {
                        // The cells are relative, so all the frozen cells move by the same distance as the scroll
                        const shifted = scrollLeft - indexWidth - 1 + 'px';
                        obj.records[j][i].element.classList.add('jss_freezed');
                        if (isLast || wasLast) {
                            obj.records[j][i].element.classList.toggle('jss_freezed_last', isLast);
                        }
                        obj.records[j][i].element.style.left = shifted;
                    }
                }
            }
        } else {
            for (let i = 0; i < freezeColumns; i++) {
                obj.headers[i].classList.remove('jss_freezed', 'jss_freezed_last');
                obj.headers[i].style.left = '';
                for (let j = 0; j < obj.rows.length; j++) {
                    if (obj.records[j][i]) {
                        obj.records[j][i].element.classList.remove('jss_freezed', 'jss_freezed_last');
                        obj.records[j][i].element.style.left = '';
                    }
                }
            }
        }

        // Place the corner in the correct place
        obj.selection.updateCorner();
    }
}
