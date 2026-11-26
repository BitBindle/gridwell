import { cutControls } from '../events/clipboard.js';
import type jspreadsheet from '../jspreadsheet';
import type { Worksheet } from '../types/internal';

/**
 * The public methods of a worksheet that call the method of a component. src/Worksheet.ts puts them on the prototype.
 */

// never[]: each method has its own parameters
export const worksheetPublicMethods: [keyof jspreadsheet.WorksheetInstance, (this: Worksheet, ...args: never[]) => unknown][] = [
    [
        'selectAll',
        function (this: Worksheet) {
            return this.selection.selectAll();
        },
    ],
    [
        'updateSelectionFromCoords',
        function (this: Worksheet, x1: number, y1: number, x2: number, y2: number) {
            return this.selection.updateFromCoords(x1, y1, x2, y2);
        },
    ],
    [
        'resetSelection',
        function (this: Worksheet) {
            return this.selection.reset();
        },
    ],
    [
        'getSelection',
        function (this: Worksheet) {
            return this.selection.get();
        },
    ],
    [
        'getSelected',
        function (this: Worksheet, columnNameOnly?: boolean) {
            return this.selection.getSelected(columnNameOnly);
        },
    ],
    [
        'getSelectedColumns',
        function (this: Worksheet, visibleOnly?: boolean) {
            return this.selection.getSelectedColumns(visibleOnly);
        },
    ],
    [
        'getSelectedRows',
        function (this: Worksheet, visibleOnly?: boolean) {
            return this.selection.getSelectedRows(visibleOnly);
        },
    ],
    [
        'getData',
        function (this: Worksheet, highlighted?: boolean, processed?: boolean, delimiter?: string, asJson?: boolean) {
            return this.data.get(highlighted, processed, delimiter, asJson);
        },
    ],
    [
        'setData',
        function (this: Worksheet, data?: jspreadsheet.CellValue[][] | Record<string, jspreadsheet.CellValue>[]) {
            return this.data.set(data);
        },
    ],
    [
        'getValue',
        function (this: Worksheet, cell: string, processedValue?: boolean) {
            return this.data.getValue(cell, processedValue);
        },
    ],
    [
        'getValueFromCoords',
        function (this: Worksheet, x: number, y: number, processedValue?: boolean) {
            return this.data.getValueFromCoords(x, y, processedValue);
        },
    ],
    [
        'setValue',
        function (this: Worksheet, cell: Parameters<jspreadsheet.WorksheetInstance['setValue']>[0], value?: jspreadsheet.CellValue, force?: boolean) {
            return this.data.setValue(cell, value, force);
        },
    ],
    [
        'setValueFromCoords',
        function (this: Worksheet, x: number, y: number, value: jspreadsheet.CellValue, force?: boolean) {
            return this.data.setValueFromCoords(x, y, value, force);
        },
    ],
    [
        'getWidth',
        function (this: Worksheet, column?: number) {
            return this.columns.getWidth(column);
        },
    ],
    [
        'setWidth',
        function (this: Worksheet, column: number | number[], width: number | number[]) {
            return this.columns.setWidth(column, width);
        },
    ],
    [
        'insertRow',
        function (this: Worksheet, mixed?: number | jspreadsheet.CellValue[], rowNumber?: number, insertBefore?: boolean | number) {
            return this.rowManager.insert(mixed, rowNumber, insertBefore);
        },
    ],
    [
        'moveRow',
        function (this: Worksheet, rowNumber: number, newPositionNumber: number) {
            return this.rowManager.move(rowNumber, newPositionNumber);
        },
    ],
    [
        'deleteRow',
        function (this: Worksheet, rowNumber?: number, numOfRows?: number) {
            return this.rowManager.delete(rowNumber, numOfRows);
        },
    ],
    [
        'hideRow',
        function (this: Worksheet, rowNumber: number | number[]) {
            return this.rowManager.hide(rowNumber);
        },
    ],
    [
        'showRow',
        function (this: Worksheet, rowNumber: number | number[]) {
            return this.rowManager.show(rowNumber);
        },
    ],
    [
        'getRowData',
        function (this: Worksheet, rowNumber: number, processed?: boolean) {
            return this.rowManager.getData(rowNumber, processed);
        },
    ],
    [
        'setRowData',
        function (this: Worksheet, rowNumber: number, data: (jspreadsheet.CellValue | null)[], force?: boolean) {
            return this.rowManager.setData(rowNumber, data, force);
        },
    ],
    [
        'getHeight',
        function (this: Worksheet, row?: number) {
            return this.rowManager.getHeight(row);
        },
    ],
    [
        'setHeight',
        function (this: Worksheet, row: number, height: number) {
            return this.rowManager.setHeight(row, height);
        },
    ],
    [
        'getMerge',
        function (this: Worksheet, cellName?: string) {
            return this.merges.get(cellName);
        },
    ],
    [
        'setMerge',
        function (this: Worksheet, cellName: string, colspan: number, rowspan: number) {
            return this.merges.set(cellName, colspan, rowspan);
        },
    ],
    [
        'destroyMerge',
        function (this: Worksheet) {
            return this.merges.destroy();
        },
    ],
    [
        'removeMerge',
        function (this: Worksheet, cellName: string, data?: jspreadsheet.CellValue[]) {
            return this.merges.remove(cellName, data);
        },
    ],
    [
        'search',
        function (this: Worksheet, query: string) {
            return this.view.search(query);
        },
    ],
    [
        'resetSearch',
        function (this: Worksheet) {
            return this.view.resetSearch();
        },
    ],
    [
        'getHeader',
        function (this: Worksheet, column: number) {
            return this.columns.getHeader(column);
        },
    ],
    [
        'getHeaders',
        function (this: Worksheet, asArray?: boolean) {
            return this.columns.getHeaders(asArray);
        },
    ],
    [
        'setHeader',
        function (this: Worksheet, column: number, newValue?: string) {
            return this.columns.setHeader(column, newValue);
        },
    ],
    [
        'getStyle',
        function (this: Worksheet, cell?: string | number[], key?: string) {
            return this.styles.get(cell, key);
        },
    ],
    [
        'setStyle',
        function (
            this: Worksheet,
            cell: string | Record<string, string | string[]>,
            property: string | null | undefined,
            value: string | null | undefined,
            forceOverwrite: boolean,
            ignoreHistoryAndEvents?: boolean
        ) {
            return this.styles.set(cell, property, value, forceOverwrite, ignoreHistoryAndEvents);
        },
    ],
    [
        'resetStyle',
        function (this: Worksheet, o: Record<string, string | string[]>, ignoreHistoryAndEvents?: boolean) {
            return this.styles.reset(o, ignoreHistoryAndEvents);
        },
    ],
    [
        'insertColumn',
        function (
            this: Worksheet,
            mixed?: number | jspreadsheet.CellValue[],
            columnNumber?: number,
            insertBefore?: boolean | number,
            properties?: jspreadsheet.Column[]
        ) {
            return this.columns.insert(mixed, columnNumber, insertBefore, properties);
        },
    ],
    [
        'moveColumn',
        function (this: Worksheet, o: number, d: number) {
            return this.columns.move(o, d);
        },
    ],
    [
        'deleteColumn',
        function (this: Worksheet, columnNumber?: number, numOfColumns?: number) {
            return this.columns.delete(columnNumber, numOfColumns);
        },
    ],
    [
        'getColumnData',
        function (this: Worksheet, columnNumber: number, processed?: boolean) {
            return this.columns.getData(columnNumber, processed);
        },
    ],
    [
        'setColumnData',
        function (this: Worksheet, colNumber: number, data: (jspreadsheet.CellValue | null)[], force?: boolean) {
            return this.columns.setData(colNumber, data, force);
        },
    ],
    [
        'whichPage',
        function (this: Worksheet, row: number) {
            return this.view.whichPage(row);
        },
    ],
    [
        'page',
        function (this: Worksheet, pageNumber: number) {
            return this.view.page(pageNumber);
        },
    ],
    [
        'download',
        function (this: Worksheet, includeHeaders?: boolean, processed?: boolean) {
            return this.clipboard.download(includeHeaders, processed);
        },
    ],
    [
        'getComments',
        function (this: Worksheet, cell?: string | number[]) {
            return this.comments.get(cell);
        },
    ],
    [
        'setComments',
        function (this: Worksheet, cellId: string | Record<string, string | null | undefined>, comments?: string | null) {
            return this.comments.set(cellId, comments);
        },
    ],
    [
        'orderBy',
        function (this: Worksheet, column: number, order?: 0 | 1) {
            return this.view.orderBy(column, order);
        },
    ],
    [
        'undo',
        function (this: Worksheet) {
            return this.history.undo();
        },
    ],
    [
        'redo',
        function (this: Worksheet) {
            return this.history.redo();
        },
    ],
    [
        'getCell',
        function (this: Worksheet, x: number, y: number) {
            return this.cells.get(x, y);
        },
    ],
    [
        'getCellFromCoords',
        function (this: Worksheet, x: number, y: number) {
            return this.cells.getFromCoords(x, y);
        },
    ],
    [
        'getLabel',
        function (this: Worksheet, x: number, y: number) {
            return this.cells.getLabel(x, y);
        },
    ],
    [
        'getMeta',
        function (this: Worksheet, cell: string) {
            return this.meta.get(cell);
        },
    ],
    [
        'setMeta',
        function (this: Worksheet, o: string | Record<string, jspreadsheet.MetaInformation>, k?: string, v?: string) {
            return this.meta.set(o, k, v);
        },
    ],
    [
        'showColumn',
        function (this: Worksheet, colNumber: number | number[]) {
            return this.columns.show(colNumber);
        },
    ],
    [
        'hideColumn',
        function (this: Worksheet, colNumber: number | number[]) {
            return this.columns.hide(colNumber);
        },
    ],
    [
        'showIndex',
        function (this: Worksheet) {
            return this.rowManager.showIndex();
        },
    ],
    [
        'hideIndex',
        function (this: Worksheet) {
            return this.rowManager.hideIndex();
        },
    ],
    [
        'openEditor',
        function (this: Worksheet, cell: HTMLTableCellElement, empty?: boolean, e?: KeyboardEvent | MouseEvent | TouchEvent) {
            return this.editor.open(cell, empty, e);
        },
    ],
    [
        'closeEditor',
        function (this: Worksheet, cell: HTMLTableCellElement, save: boolean) {
            return this.editor.close(cell, save);
        },
    ],
    [
        'copy',
        function (this: Worksheet, cut: boolean) {
            if (cut) {
                cutControls();
            } else {
                this.clipboard.copy(true);
            }
        },
    ],
    [
        'paste',
        function (this: Worksheet, x: number, y: number, data: string) {
            return this.clipboard.paste(x, y, data);
        },
    ],
    [
        'executeFormula',
        function (this: Worksheet, expression: string, x?: number, y?: number) {
            return this.formulas.execute(expression, x, y);
        },
    ],
    [
        'getDataFromRange',
        function (this: Worksheet, range: string, processed: true) {
            return this.data.getFromRange(range, processed);
        },
    ],
    [
        'quantiyOfPages',
        function (this: Worksheet) {
            return this.view.countPages();
        },
    ],
    [
        'getRange',
        function (this: Worksheet) {
            return this.selection.getRange();
        },
    ],
    [
        'isSelected',
        function (this: Worksheet, x: number, y: number) {
            return this.selection.isSelected(x, y);
        },
    ],
    [
        'setReadOnly',
        function (this: Worksheet, cell: string | HTMLElement, state: boolean) {
            return this.cells.setReadOnly(cell, state);
        },
    ],
    [
        'isReadOnly',
        function (this: Worksheet, x: number | string, y?: number) {
            return this.cells.isReadOnly(x, y);
        },
    ],
    [
        'getHighlighted',
        function (this: Worksheet) {
            return this.selection.getHighlighted();
        },
    ],
    [
        'down',
        function (this: Worksheet, shiftKey?: boolean, ctrlKey?: boolean) {
            return this.selection.down(shiftKey, ctrlKey);
        },
    ],
    [
        'first',
        function (this: Worksheet, shiftKey?: boolean, ctrlKey?: boolean) {
            return this.selection.first(shiftKey, ctrlKey);
        },
    ],
    [
        'last',
        function (this: Worksheet, shiftKey?: boolean, ctrlKey?: boolean) {
            return this.selection.last(shiftKey, ctrlKey);
        },
    ],
    [
        'left',
        function (this: Worksheet, shiftKey?: boolean, ctrlKey?: boolean) {
            return this.selection.left(shiftKey, ctrlKey);
        },
    ],
    [
        'right',
        function (this: Worksheet, shiftKey?: boolean, ctrlKey?: boolean) {
            return this.selection.right(shiftKey, ctrlKey);
        },
    ],
    [
        'up',
        function (this: Worksheet, shiftKey?: boolean, ctrlKey?: boolean) {
            return this.selection.up(shiftKey, ctrlKey);
        },
    ],
    [
        'openFilter',
        function (this: Worksheet, columnId: number) {
            return this.view.openFilter(columnId);
        },
    ],
    [
        'resetFilters',
        function (this: Worksheet) {
            return this.view.resetFilters();
        },
    ],
];
