import jSuites from 'jsuites';
import { setHistory } from './history.js';
import dispatch from './dispatch.js';
import { updateTableReferences } from './internal.js';
import { loadPage } from './lazyLoading.js';
import { closeFilter } from './filter.js';
import { updateTabStop } from './accessibility.js';
import type jspreadsheet from '../jspreadsheet';
import type { Worksheet, WorksheetMethod } from '../types/internal';

/**
 * Update order arrow
 */
export const updateOrderArrow = function (this: Worksheet, column: number, order: boolean | number) {
    const obj = this;

    // Remove order
    for (let i = 0; i < obj.headers.length; i++) {
        obj.headers[i].classList.remove('arrow-up');
        obj.headers[i].classList.remove('arrow-down');
    }

    // No order specified then toggle order
    if (order) {
        obj.headers[column].classList.add('arrow-up');
    } else {
        obj.headers[column].classList.add('arrow-down');
    }
};

/**
 * Update rows position
 */
export const updateOrder = function (this: Worksheet, rows: number[]) {
    const obj = this;

    // History
    const data: Worksheet['options']['data'] = [];
    for (let j = 0; j < rows.length; j++) {
        data[j] = obj.options.data[rows[j]];
    }
    obj.options.data = data;

    const records: Worksheet['records'] = [];
    for (let j = 0; j < rows.length; j++) {
        records[j] = obj.records[rows[j]];

        for (let i = 0; i < records[j].length; i++) {
            records[j][i].y = j;
        }
    }
    obj.records = records;

    const rowObjects: Worksheet['rows'] = [];
    for (let j = 0; j < rows.length; j++) {
        rowObjects[j] = obj.rows[rows[j]];
        rowObjects[j].y = j;
    }
    obj.rows = rowObjects;

    // Update references
    updateTableReferences.call(obj);

    // Redo search
    if (obj.results && obj.results.length) {
        if (obj.searchInput.value) {
            obj.search(obj.searchInput.value);
        } else {
            closeFilter.call(obj);
        }
    } else {
        // Create page
        obj.results = null;
        obj.pageNumber = 0;

        if (obj.options.pagination! > 0) {
            obj.page(0);
        } else if (obj.options.lazyLoading == true) {
            loadPage.call(obj, 0);
        } else {
            for (let j = 0; j < obj.rows.length; j++) {
                obj.tbody.appendChild(obj.rows[j].element);
            }
            // A moved row loses the focus
            updateTabStop.call(obj);
        }
    }
};

/**
 * Sort data and reload table
 */
export const orderBy: WorksheetMethod<'orderBy'> = function (column, order) {
    const obj = this;

    if (column >= 0) {
        // Merged cells
        if (obj.options.mergeCells && Object.keys(obj.options.mergeCells).length > 0) {
            if (!confirm(jSuites.translate('This action will destroy any existing merged cells. Are you sure?'))) {
                return false;
            } else {
                // Remove merged cells
                obj.destroyMerge();
            }
        }

        // Direction
        if (order == null) {
            order = obj.headers[column].classList.contains('arrow-down') ? 1 : 0;
        } else {
            order = order ? 1 : 0;
        }

        // Test order
        // The row and the value to sort by
        let temp: jspreadsheet.SortingItem[] = [];
        if (
            obj.options.columns &&
            obj.options.columns[column] &&
            ((obj.options.columns[column].type as string) == 'number' ||
                obj.options.columns[column].type == 'numeric' ||
                (obj.options.columns[column].type as string) == 'percentage' ||
                (obj.options.columns[column].type as string) == 'autonumber' ||
                obj.options.columns[column].type == 'color')
        ) {
            for (let j = 0; j < obj.options.data.length; j++) {
                temp[j] = [j, Number(obj.options.data[j][column])];
            }
        } else if (
            obj.options.columns &&
            obj.options.columns[column] &&
            (obj.options.columns[column].type == 'calendar' || obj.options.columns[column].type == 'checkbox' || obj.options.columns[column].type == 'radio')
        ) {
            for (let j = 0; j < obj.options.data.length; j++) {
                temp[j] = [j, obj.options.data[j][column]];
            }
        } else {
            for (let j = 0; j < obj.options.data.length; j++) {
                temp[j] = [j, obj.records[j][column].element.textContent.toLowerCase()];
            }
        }

        // Default sorting method
        if (typeof obj.parent.config.sorting !== 'function') {
            obj.parent.config.sorting = function (direction) {
                return function (a, b) {
                    const valueA = a[1];
                    const valueB = b[1];

                    if (!direction) {
                        return valueA === '' && valueB !== '' ? 1 : valueA !== '' && valueB === '' ? -1 : valueA > valueB ? 1 : valueA < valueB ? -1 : 0;
                    } else {
                        return valueA === '' && valueB !== '' ? 1 : valueA !== '' && valueB === '' ? -1 : valueA > valueB ? -1 : valueA < valueB ? 1 : 0;
                    }
                };
            };
        }

        temp = temp.sort(obj.parent.config.sorting(order));

        // Save history
        const newValue = [];
        for (let j = 0; j < temp.length; j++) {
            newValue[j] = temp[j][0];
        }

        // Save history
        setHistory.call(obj, {
            action: 'orderBy',
            rows: newValue,
            column: column,
            order: order,
        });

        // Update order
        updateOrderArrow.call(obj, column, order);
        updateOrder.call(obj, newValue);

        // On sort event
        dispatch.call(
            obj,
            'onsort',
            obj,
            column,
            order,
            newValue.map((row) => row)
        );

        return true;
    }
};
