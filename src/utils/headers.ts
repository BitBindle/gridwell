import dispatch from './dispatch.js';
import { getColumnName } from './helpers.js';
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
