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
        if (this.ignore != true) {
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
        const obj = this.worksheet;

        const rowIndex = !historyRecord.insertBefore ? historyRecord.rowNumber + 1 : +historyRecord.rowNumber;

        if (obj.options.search == true) {
            if (obj.view.results && obj.view.results.length != obj.rows.length) {
                obj.resetSearch();
            }
        }

        // Remove row
        if (type == 1) {
            const numOfRows = historyRecord.numOfRows;
            // Remove nodes
            for (let j = rowIndex; j < numOfRows + rowIndex; j++) {
                obj.rows[j].element.parentNode!.removeChild(obj.rows[j].element);
            }
            // Remove references
            obj.records.splice(rowIndex, numOfRows);
            obj.options.data.splice(rowIndex, numOfRows);
            obj.rows.splice(rowIndex, numOfRows);

            obj.selection.conditionalUpdate(1, rowIndex, numOfRows + rowIndex - 1);
        } else {
            // Insert data
            const records = historyRecord.rowRecords.map((row) => {
                return [...row];
            });
            obj.records = injectArray(obj.records, rowIndex, records);

            const data = historyRecord.rowData.map((row) => {
                return [...row];
            });
            obj.options.data = injectArray(obj.options.data, rowIndex, data);

            obj.rows = injectArray(obj.rows, rowIndex, historyRecord.rowNode);
            // Insert nodes
            let index = 0;
            for (let j = rowIndex; j < historyRecord.numOfRows + rowIndex; j++) {
                obj.tbody.insertBefore(historyRecord.rowNode[index].element, obj.tbody.children[j]);
                index++;
            }
        }

        for (let j = rowIndex; j < obj.rows.length; j++) {
            obj.rows[j].y = j;
        }

        for (let j = rowIndex; j < obj.records.length; j++) {
            for (let i = 0; i < obj.records[j].length; i++) {
                obj.records[j][i].y = j;
            }
        }

        // Respect pagination
        if (obj.options.pagination! > 0) {
            obj.page(obj.view.pageNumber);
        }

        obj.formulas.updateTableReferences();
    }

    /**
     * Process column
     */
    private processColumn(type: HistoryProcess, historyRecord: Extract<WorksheetHistoryRecord, { action: 'insertColumn' | 'deleteColumn' }>) {
        const obj = this.worksheet;

        const columnIndex = !historyRecord.insertBefore ? historyRecord.columnNumber + 1 : historyRecord.columnNumber;

        // Remove column
        if (type == 1) {
            const numOfColumns = historyRecord.numOfColumns;

            obj.options.columns.splice(columnIndex, numOfColumns);
            for (let i = columnIndex; i < numOfColumns + columnIndex; i++) {
                obj.headers[i].parentNode!.removeChild(obj.headers[i]);
                obj.cols[i].colElement.parentNode!.removeChild(obj.cols[i].colElement);
            }
            obj.headers.splice(columnIndex, numOfColumns);
            obj.cols.splice(columnIndex, numOfColumns);
            for (let j = 0; j < historyRecord.data.length; j++) {
                for (let i = columnIndex; i < numOfColumns + columnIndex; i++) {
                    obj.records[j][i].element.parentNode!.removeChild(obj.records[j][i].element);
                }
                obj.records[j].splice(columnIndex, numOfColumns);
                obj.options.data[j].splice(columnIndex, numOfColumns);
            }
            // Process footers
            if (obj.options.footers) {
                for (let j = 0; j < obj.options.footers.length; j++) {
                    obj.options.footers[j].splice(columnIndex, numOfColumns);
                }
            }
        } else {
            // Insert data
            // The worksheet always has column options, so deleteColumn stored them
            obj.options.columns = injectArray(obj.options.columns, columnIndex, historyRecord.columns!);
            obj.headers = injectArray(obj.headers, columnIndex, historyRecord.headers);
            obj.cols = injectArray(obj.cols, columnIndex, historyRecord.cols);

            let index = 0;
            for (let i = columnIndex; i < historyRecord.numOfColumns + columnIndex; i++) {
                obj.headerContainer.insertBefore(historyRecord.headers[index], obj.headerContainer.children[i + 1]);
                obj.colgroupContainer.insertBefore(historyRecord.cols[index].colElement, obj.colgroupContainer.children[i + 1]);
                index++;
            }

            for (let j = 0; j < historyRecord.data.length; j++) {
                obj.options.data[j] = injectArray(obj.options.data[j], columnIndex, historyRecord.data[j]);
                obj.records[j] = injectArray(obj.records[j], columnIndex, historyRecord.records[j]);
                let index = 0;
                for (let i = columnIndex; i < historyRecord.numOfColumns + columnIndex; i++) {
                    obj.rows[j].element.insertBefore(historyRecord.records[j][index].element, obj.rows[j].element.children[i + 1]);
                    index++;
                }
            }
            // Process footers
            if (obj.options.footers) {
                for (let j = 0; j < obj.options.footers.length; j++) {
                    obj.options.footers[j] = injectArray(obj.options.footers[j], columnIndex, historyRecord.footers[j]);
                }
            }
        }

        for (let i = columnIndex; i < obj.cols.length; i++) {
            obj.cols[i].x = i;
        }

        for (let j = 0; j < obj.records.length; j++) {
            for (let i = columnIndex; i < obj.records[j].length; i++) {
                obj.records[j][i].x = i;
            }
        }

        // Adjust nested headers
        if (obj.options.nestedHeaders && obj.options.nestedHeaders.length > 0 && obj.options.nestedHeaders[0] && obj.options.nestedHeaders[0][0]) {
            for (let j = 0; j < obj.options.nestedHeaders.length; j++) {
                let colspan;

                if (type == 1) {
                    colspan = parseInt(obj.options.nestedHeaders[j][obj.options.nestedHeaders[j].length - 1].colspan) - historyRecord.numOfColumns;
                } else {
                    colspan = parseInt(obj.options.nestedHeaders[j][obj.options.nestedHeaders[j].length - 1].colspan) + historyRecord.numOfColumns;
                }
                obj.options.nestedHeaders[j][obj.options.nestedHeaders[j].length - 1].colspan = colspan;
                obj.thead.children[j].children[obj.thead.children[j].children.length - 1].setAttribute('colspan', colspan);
            }
        }

        obj.formulas.updateTableReferences();
    }

    /**
     * Undo last action
     */
    undo() {
        const obj = this.worksheet;

        // Ignore events and history
        const ignoreEvents = obj.parent.ignoreEvents ? true : false;
        const ignoreHistory = obj.history.ignore ? true : false;

        obj.parent.ignoreEvents = true;
        obj.history.ignore = true;

        // Records
        const records = [];

        // Update cells
        const historyRecord = obj.history.back();

        if (historyRecord) {
            if (historyRecord.action == 'insertRow') {
                this.processRow(1, historyRecord);
            } else if (historyRecord.action == 'deleteRow') {
                this.processRow(0, historyRecord);
            } else if (historyRecord.action == 'insertColumn') {
                this.processColumn(1, historyRecord);
            } else if (historyRecord.action == 'deleteColumn') {
                this.processColumn(0, historyRecord);
            } else if (historyRecord.action == 'moveRow') {
                obj.moveRow(historyRecord.newValue, historyRecord.oldValue);
            } else if (historyRecord.action == 'moveColumn') {
                obj.moveColumn(historyRecord.newValue, historyRecord.oldValue);
            } else if (historyRecord.action == 'setMerge') {
                obj.removeMerge(historyRecord.column, historyRecord.data);
            } else if (historyRecord.action == 'setStyle') {
                obj.setStyle(historyRecord.oldValue, null, null, true); // true means force
            } else if (historyRecord.action == 'setWidth') {
                // The public overloads take one column or a list, not either
                (obj.setWidth as (column: number | number[], width: number | number[]) => void)(historyRecord.column, historyRecord.oldValue);
            } else if (historyRecord.action == 'setHeight') {
                obj.setHeight(historyRecord.row, historyRecord.oldValue);
            } else if (historyRecord.action == 'setHeader') {
                obj.setHeader(historyRecord.column, historyRecord.oldValue);
            } else if (historyRecord.action == 'setComments') {
                obj.setComments(historyRecord.oldValue);
            } else if (historyRecord.action == 'orderBy') {
                const rows = [];
                for (let j = 0; j < historyRecord.rows.length; j++) {
                    rows[historyRecord.rows[j]] = j;
                }
                obj.view.updateOrderArrow(historyRecord.column, historyRecord.order ? 0 : 1);
                obj.view.updateOrder(rows);
            } else if (historyRecord.action == 'setValue') {
                // Redo for changes in cells
                for (let i = 0; i < historyRecord.records.length; i++) {
                    records.push({
                        x: historyRecord.records[i].x,
                        y: historyRecord.records[i].y,
                        value: historyRecord.records[i].oldValue,
                    });

                    if (historyRecord.oldStyle) {
                        obj.resetStyle(historyRecord.oldStyle);
                    }
                }
                // Update records
                obj.setValue(records);

                // Update selection
                if (historyRecord.selection) {
                    obj.updateSelectionFromCoords(
                        historyRecord.selection[0],
                        historyRecord.selection[1],
                        historyRecord.selection[2],
                        historyRecord.selection[3]
                    );
                }
            }
        }
        obj.parent.ignoreEvents = ignoreEvents;
        obj.history.ignore = ignoreHistory;

        // Events
        obj.dispatch('onundo', obj, historyRecord);
    }

    /**
     * Redo previously undone action
     */
    redo() {
        const obj = this.worksheet;

        // Ignore events and history
        const ignoreEvents = obj.parent.ignoreEvents ? true : false;
        const ignoreHistory = obj.history.ignore ? true : false;

        obj.parent.ignoreEvents = true;
        obj.history.ignore = true;

        // Update cells
        const historyRecord = obj.history.forward();

        if (historyRecord) {
            if (historyRecord.action == 'insertRow') {
                this.processRow(0, historyRecord);
            } else if (historyRecord.action == 'deleteRow') {
                this.processRow(1, historyRecord);
            } else if (historyRecord.action == 'insertColumn') {
                this.processColumn(0, historyRecord);
            } else if (historyRecord.action == 'deleteColumn') {
                this.processColumn(1, historyRecord);
            } else if (historyRecord.action == 'moveRow') {
                obj.moveRow(historyRecord.oldValue, historyRecord.newValue);
            } else if (historyRecord.action == 'moveColumn') {
                obj.moveColumn(historyRecord.oldValue, historyRecord.newValue);
            } else if (historyRecord.action == 'setMerge') {
                obj.merges.set(historyRecord.column, historyRecord.colspan, historyRecord.rowspan, 1);
            } else if (historyRecord.action == 'setStyle') {
                obj.setStyle(historyRecord.newValue, null, null, true); // true means force
            } else if (historyRecord.action == 'setWidth') {
                (obj.setWidth as (column: number | number[], width: number | number[]) => void)(historyRecord.column, historyRecord.newValue);
            } else if (historyRecord.action == 'setHeight') {
                obj.setHeight(historyRecord.row, historyRecord.newValue);
            } else if (historyRecord.action == 'setHeader') {
                obj.setHeader(historyRecord.column, historyRecord.newValue);
            } else if (historyRecord.action == 'setComments') {
                obj.setComments(historyRecord.newValue);
            } else if (historyRecord.action == 'orderBy') {
                obj.view.updateOrderArrow(historyRecord.column, historyRecord.order);
                obj.view.updateOrder(historyRecord.rows);
            } else if (historyRecord.action == 'setValue') {
                obj.setValue(historyRecord.records);
                // Redo for changes in cells
                for (let i = 0; i < historyRecord.records.length; i++) {
                    if (historyRecord.oldStyle) {
                        obj.resetStyle(historyRecord.newStyle!);
                    }
                }
                // Update selection
                if (historyRecord.selection) {
                    obj.updateSelectionFromCoords(
                        historyRecord.selection[0],
                        historyRecord.selection[1],
                        historyRecord.selection[2],
                        historyRecord.selection[3]
                    );
                }
            }
        }
        obj.parent.ignoreEvents = ignoreEvents;
        obj.history.ignore = ignoreHistory;

        // Events
        obj.dispatch('onredo', obj, historyRecord);
    }
}
