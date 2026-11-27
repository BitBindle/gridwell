import { injectArray } from '../utils/internalHelpers.js';
import type { WorksheetHistoryRecord, Worksheet } from '../types/internal';

/** 1 removes the rows or columns of the record, 0 inserts them again. */
type HistoryProcess = 0 | 1;

/**
 * The undo history of a worksheet: the records of the changes, the position of undo and redo in them, and the flag that
 * stops new records while a change must not be recorded. It undoes and redoes the records.
 *
 * Each worksheet has one, as worksheet.history.
 */
export class History {
    readonly worksheet: Worksheet;

    /** The records, oldest first. */
    records: WorksheetHistoryRecord[] = [];

    /** The index of the record that undo reverts next, or -1 when there is nothing to undo. */
    index = -1;

    /** While true, add keeps no records. Undo and redo set it while they apply a record. */
    ignore = false;

    constructor(worksheet: Worksheet) {
        this.worksheet = worksheet;
    }

    /**
     * Keep a record of a change, unless the history ignores changes. The records that were undone are dropped, so they
     * can no longer be redone.
     */
    add(record: WorksheetHistoryRecord) {
        if (this.ignore !== true) {
            // Increment and get the current history index
            const index = ++this.index;

            // Slice the array to discard undone changes
            this.records = this.records.slice(0, index + 1);

            // Keep history
            this.records[index] = record;
        }
    }

    /** Drop all the records. */
    clear() {
        this.records = [];
        this.index = -1;
    }

    /** Whether there is a record to undo. */
    canUndo() {
        return this.index >= 0;
    }

    /** Whether there is a record to redo. */
    canRedo() {
        return this.index < this.records.length - 1;
    }

    /** The record to undo, or undefined when there is none. The next record to undo is the one before it. */
    back(): WorksheetHistoryRecord | undefined {
        if (this.canUndo()) {
            return this.records[this.index--];
        }
        return undefined;
    }

    /** The record to redo, or undefined when there is none. The next record to redo is the one after it. */
    forward(): WorksheetHistoryRecord | undefined {
        if (this.canRedo()) {
            return this.records[++this.index];
        }
        return undefined;
    }

    /**
     * Process row
     */
    private processRow(type: HistoryProcess, historyRecord: Extract<WorksheetHistoryRecord, { action: 'insertRow' | 'deleteRow' }>) {
        const worksheet = this.worksheet;

        const rowIndex = !historyRecord.insertBefore ? historyRecord.rowNumber + 1 : +historyRecord.rowNumber;

        if (worksheet.options.search === true) {
            if (worksheet.view.results && worksheet.view.results.length !== worksheet.rows.length) {
                worksheet.resetSearch();
            }
        }

        // Remove row
        if (type === 1) {
            const numOfRows = historyRecord.numOfRows;
            // An editor in a removed row closes without a save
            worksheet.editor.closeIn(1, rowIndex, rowIndex + numOfRows - 1);
            // Remove nodes
            for (let j = rowIndex; j < numOfRows + rowIndex; j++) {
                worksheet.rows[j].element.parentNode!.removeChild(worksheet.rows[j].element);
            }
            // Remove references
            worksheet.records.splice(rowIndex, numOfRows);
            worksheet.options.data.splice(rowIndex, numOfRows);
            worksheet.rows.splice(rowIndex, numOfRows);
        } else {
            // Insert data
            const records = historyRecord.rowRecords.map((row) => {
                return [...row];
            });
            worksheet.records = injectArray(worksheet.records, rowIndex, records);

            const data = historyRecord.rowData.map((row) => {
                return [...row];
            });
            worksheet.options.data = injectArray(worksheet.options.data, rowIndex, data);

            worksheet.rows = injectArray(worksheet.rows, rowIndex, historyRecord.rowNode);
            // Insert nodes
            let index = 0;
            for (let j = rowIndex; j < historyRecord.numOfRows + rowIndex; j++) {
                worksheet.tbody.insertBefore(historyRecord.rowNode[index].element, worksheet.tbody.children[j]);
                index++;
            }
        }

        for (let j = rowIndex; j < worksheet.rows.length; j++) {
            worksheet.rows[j].y = j;
        }

        for (let j = rowIndex; j < worksheet.records.length; j++) {
            for (const cell of worksheet.records[j]) {
                cell.y = j;
            }
        }

        // Respect pagination
        if (worksheet.options.pagination! > 0) {
            worksheet.page(worksheet.view.pageNumber);
        }

        // The selection and the editor stay on their cells. A selection with a removed cell is removed.
        const count = type === 1 ? -historyRecord.numOfRows : historyRecord.numOfRows;
        worksheet.editor.shift(1, rowIndex, count);
        worksheet.selection.shift(1, rowIndex, count);

        // The rows that the undo or the redo adds keep their formulas
        worksheet.formulas.updateTableReferences(type === 1 ? {} : { inserted: { axis: 1, index: rowIndex, count: historyRecord.numOfRows } });
    }

    /**
     * Process column
     */
    private processColumn(type: HistoryProcess, historyRecord: Extract<WorksheetHistoryRecord, { action: 'insertColumn' | 'deleteColumn' }>) {
        const worksheet = this.worksheet;

        const columnIndex = !historyRecord.insertBefore ? historyRecord.columnNumber + 1 : historyRecord.columnNumber;

        // Remove column
        if (type === 1) {
            const numOfColumns = historyRecord.numOfColumns;
            // An editor in a removed column closes without a save
            worksheet.editor.closeIn(0, columnIndex, columnIndex + numOfColumns - 1);

            worksheet.options.columns.splice(columnIndex, numOfColumns);
            for (let i = columnIndex; i < numOfColumns + columnIndex; i++) {
                worksheet.headers[i].parentNode!.removeChild(worksheet.headers[i]);
                worksheet.cols[i].colElement.parentNode!.removeChild(worksheet.cols[i].colElement);
            }
            worksheet.headers.splice(columnIndex, numOfColumns);
            worksheet.cols.splice(columnIndex, numOfColumns);
            for (let j = 0; j < historyRecord.data.length; j++) {
                for (let i = columnIndex; i < numOfColumns + columnIndex; i++) {
                    worksheet.records[j][i].element.parentNode!.removeChild(worksheet.records[j][i].element);
                }
                worksheet.records[j].splice(columnIndex, numOfColumns);
                worksheet.options.data[j].splice(columnIndex, numOfColumns);
            }
            // Process footers
            if (worksheet.options.footers) {
                for (const footer of worksheet.options.footers) {
                    footer.splice(columnIndex, numOfColumns);
                }
            }
        } else {
            // Insert data
            // The worksheet always has column options, so deleteColumn stored them
            worksheet.options.columns = injectArray(worksheet.options.columns, columnIndex, historyRecord.columns!);
            worksheet.headers = injectArray(worksheet.headers, columnIndex, historyRecord.headers);
            worksheet.cols = injectArray(worksheet.cols, columnIndex, historyRecord.cols);

            let index = 0;
            for (let i = columnIndex; i < historyRecord.numOfColumns + columnIndex; i++) {
                worksheet.headerContainer.insertBefore(historyRecord.headers[index], worksheet.headerContainer.children[i + 1]);
                worksheet.colgroupContainer.insertBefore(historyRecord.cols[index].colElement, worksheet.colgroupContainer.children[i + 1]);
                index++;
            }

            for (let j = 0; j < historyRecord.data.length; j++) {
                worksheet.options.data[j] = injectArray(worksheet.options.data[j], columnIndex, historyRecord.data[j]);
                worksheet.records[j] = injectArray(worksheet.records[j], columnIndex, historyRecord.records[j]);
                let index = 0;
                for (let i = columnIndex; i < historyRecord.numOfColumns + columnIndex; i++) {
                    worksheet.rows[j].element.insertBefore(historyRecord.records[j][index].element, worksheet.rows[j].element.children[i + 1]);
                    index++;
                }
            }
            // Process footers
            if (worksheet.options.footers) {
                for (let j = 0; j < worksheet.options.footers.length; j++) {
                    worksheet.options.footers[j] = injectArray(worksheet.options.footers[j], columnIndex, historyRecord.footers[j]);
                }
            }
        }

        for (let i = columnIndex; i < worksheet.cols.length; i++) {
            worksheet.cols[i].x = i;
        }

        for (const record of worksheet.records) {
            for (let i = columnIndex; i < record.length; i++) {
                record[i].x = i;
            }
        }

        // Adjust nested headers
        if (worksheet.options.nestedHeaders && worksheet.options.nestedHeaders.length > 0 && worksheet.options.nestedHeaders[0]?.[0]) {
            for (let j = 0; j < worksheet.options.nestedHeaders.length; j++) {
                let colspan;

                if (type === 1) {
                    colspan = parseInt(worksheet.options.nestedHeaders[j][worksheet.options.nestedHeaders[j].length - 1].colspan) - historyRecord.numOfColumns;
                } else {
                    colspan = parseInt(worksheet.options.nestedHeaders[j][worksheet.options.nestedHeaders[j].length - 1].colspan) + historyRecord.numOfColumns;
                }
                worksheet.options.nestedHeaders[j][worksheet.options.nestedHeaders[j].length - 1].colspan = colspan;
                worksheet.thead.children[j].children[worksheet.thead.children[j].children.length - 1].setAttribute('colspan', colspan);
            }
        }

        // The frozen columns can be other columns now
        worksheet.columns.updateFrozenColumns();

        // The selection and the editor stay on their cells. A selection with a removed cell is removed.
        const count = type === 1 ? -historyRecord.numOfColumns : historyRecord.numOfColumns;
        worksheet.editor.shift(0, columnIndex, count);
        worksheet.selection.shift(0, columnIndex, count);

        // The columns that the undo or the redo adds keep their formulas
        worksheet.formulas.updateTableReferences(type === 1 ? {} : { inserted: { axis: 0, index: columnIndex, count: historyRecord.numOfColumns } });
    }

    /**
     * Undo last action
     */
    undo() {
        const worksheet = this.worksheet;

        // Ignore events and history
        const ignoreEvents = worksheet.parent.ignoreEvents ? true : false;
        const ignoreHistory = worksheet.history.ignore ? true : false;

        worksheet.parent.ignoreEvents = true;
        worksheet.history.ignore = true;

        // Records
        const records = [];

        // Update cells
        const historyRecord = worksheet.history.back();

        if (historyRecord) {
            if (historyRecord.action === 'insertRow') {
                this.processRow(1, historyRecord);
            } else if (historyRecord.action === 'deleteRow') {
                this.processRow(0, historyRecord);
            } else if (historyRecord.action === 'insertColumn') {
                this.processColumn(1, historyRecord);
            } else if (historyRecord.action === 'deleteColumn') {
                this.processColumn(0, historyRecord);
            } else if (historyRecord.action === 'moveRow') {
                worksheet.moveRow(historyRecord.newValue, historyRecord.oldValue);
            } else if (historyRecord.action === 'moveColumn') {
                worksheet.moveColumn(historyRecord.newValue, historyRecord.oldValue);
            } else if (historyRecord.action === 'setMerge') {
                worksheet.removeMerge(historyRecord.column, historyRecord.data);
            } else if (historyRecord.action === 'setStyle') {
                worksheet.setStyle(historyRecord.oldValue, null, null, true); // true means force
            } else if (historyRecord.action === 'setWidth') {
                // The public overloads take one column or a list, not either
                (worksheet.setWidth as (column: number | number[], width: number | number[]) => void)(historyRecord.column, historyRecord.oldValue);
            } else if (historyRecord.action === 'setHeight') {
                worksheet.setHeight(historyRecord.row, historyRecord.oldValue);
            } else if (historyRecord.action === 'setHeader') {
                worksheet.setHeader(historyRecord.column, historyRecord.oldValue);
            } else if (historyRecord.action === 'setComments') {
                worksheet.setComments(historyRecord.oldValue);
            } else if (historyRecord.action === 'orderBy') {
                const rows = [];
                for (let j = 0; j < historyRecord.rows.length; j++) {
                    rows[historyRecord.rows[j]] = j;
                }
                worksheet.view.updateOrderArrow(historyRecord.column, historyRecord.order ? 0 : 1);
                worksheet.view.updateOrder(rows);
            } else if (historyRecord.action === 'setValue') {
                // Redo for changes in cells
                for (const record of historyRecord.records) {
                    records.push({
                        x: record.x,
                        y: record.y,
                        value: record.oldValue,
                    });

                    if (historyRecord.oldStyle) {
                        worksheet.resetStyle(historyRecord.oldStyle);
                    }
                }
                // Update records
                worksheet.setValue(records);

                // Update selection
                if (historyRecord.selection) {
                    worksheet.updateSelectionFromCoords(
                        historyRecord.selection[0],
                        historyRecord.selection[1],
                        historyRecord.selection[2],
                        historyRecord.selection[3]
                    );
                }
            }
        }
        worksheet.parent.ignoreEvents = ignoreEvents;
        worksheet.history.ignore = ignoreHistory;

        // Events
        worksheet.dispatch('onundo', worksheet, historyRecord);
    }

    /**
     * Redo previously undone action
     */
    redo() {
        const worksheet = this.worksheet;

        // Ignore events and history
        const ignoreEvents = worksheet.parent.ignoreEvents ? true : false;
        const ignoreHistory = worksheet.history.ignore ? true : false;

        worksheet.parent.ignoreEvents = true;
        worksheet.history.ignore = true;

        // Update cells
        const historyRecord = worksheet.history.forward();

        if (historyRecord) {
            if (historyRecord.action === 'insertRow') {
                this.processRow(0, historyRecord);
            } else if (historyRecord.action === 'deleteRow') {
                this.processRow(1, historyRecord);
            } else if (historyRecord.action === 'insertColumn') {
                this.processColumn(0, historyRecord);
            } else if (historyRecord.action === 'deleteColumn') {
                this.processColumn(1, historyRecord);
            } else if (historyRecord.action === 'moveRow') {
                worksheet.moveRow(historyRecord.oldValue, historyRecord.newValue);
            } else if (historyRecord.action === 'moveColumn') {
                worksheet.moveColumn(historyRecord.oldValue, historyRecord.newValue);
            } else if (historyRecord.action === 'setMerge') {
                worksheet.merges.set(historyRecord.column, historyRecord.colspan, historyRecord.rowspan, 1);
            } else if (historyRecord.action === 'setStyle') {
                worksheet.setStyle(historyRecord.newValue, null, null, true); // true means force
            } else if (historyRecord.action === 'setWidth') {
                (worksheet.setWidth as (column: number | number[], width: number | number[]) => void)(historyRecord.column, historyRecord.newValue);
            } else if (historyRecord.action === 'setHeight') {
                worksheet.setHeight(historyRecord.row, historyRecord.newValue);
            } else if (historyRecord.action === 'setHeader') {
                worksheet.setHeader(historyRecord.column, historyRecord.newValue);
            } else if (historyRecord.action === 'setComments') {
                worksheet.setComments(historyRecord.newValue);
            } else if (historyRecord.action === 'orderBy') {
                worksheet.view.updateOrderArrow(historyRecord.column, historyRecord.order);
                worksheet.view.updateOrder(historyRecord.rows);
            } else if (historyRecord.action === 'setValue') {
                worksheet.setValue(historyRecord.records);
                // Redo for changes in cells
                for (const record of historyRecord.records) {
                    if (historyRecord.oldStyle) {
                        worksheet.resetStyle(historyRecord.newStyle!);
                    }
                }
                // Update selection
                if (historyRecord.selection) {
                    worksheet.updateSelectionFromCoords(
                        historyRecord.selection[0],
                        historyRecord.selection[1],
                        historyRecord.selection[2],
                        historyRecord.selection[3]
                    );
                }
            }
        }
        worksheet.parent.ignoreEvents = ignoreEvents;
        worksheet.history.ignore = ignoreHistory;

        // Events
        worksheet.dispatch('onredo', worksheet, historyRecord);
    }
}
