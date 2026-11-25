import dispatch from './dispatch.js';
import { getColumnName } from './helpers.js';
import type jspreadsheet from '../jspreadsheet';
import type { Worksheet } from '../types/internal';

/**
 * Get the column title
 * @param column - column number (first column is: 0)
 */
export const getHeader = function (this: Worksheet, column: number): string {
    const obj = this;

    return obj.headers[column].textContent;
};

/**
 * Get the headers
 * @param asArray - Return an array instead of a text with the CSV delimiter.
 */
export const getHeaders = function (this: Worksheet, asArray?: boolean): string | string[] {
    const obj = this;

    const title = [];

    for (let i = 0; i < obj.headers.length; i++) {
        title.push(obj.getHeader(i));
    }

    return asArray ? title : title.join(obj.options.csvDelimiter);
};

/**
 * Set the column title
 * @param column - column number (first column is: 0)
 * @param newValue - new column title. Without it, the title is the column name, for example "B".
 */
export const setHeader = function (this: Worksheet, column: number, newValue?: string) {
    const obj = this;

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
};

/**
 * Create a nested header object
 */
// The row of nested headers keeps its element
export const createNestedHeader = function (this: Worksheet, nestedInformation: jspreadsheet.NestedHeaderCell[] & { element?: HTMLTableRowElement }) {
    const obj = this;

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
};
