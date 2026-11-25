// @ts-check
import { getCoordsFromCellName } from './helpers.js';
/** @import { Worksheet } from '../types/internal' */

/**
 * Make a cell read-only, or editable again.
 * @this {Worksheet}
 * @param {string | HTMLElement} cell - The cell name, for example "B3", or the td of the cell.
 * @param {boolean} state - Whether the cell is read-only.
 */
export const setReadOnly = function (cell, state) {
    const obj = this;

    let record;

    if (typeof cell === 'string') {
        const coords = getCoordsFromCellName(cell);

        record = obj.records[coords[1]][coords[0]];
    } else {
        const x = parseInt(cell.getAttribute('data-x'));
        const y = parseInt(cell.getAttribute('data-y'));

        record = obj.records[y][x];
    }

    if (state) {
        record.element.classList.add('readonly');
    } else {
        record.element.classList.remove('readonly');
    }
};

/**
 * Check if a cell is read-only.
 * @this {Worksheet}
 * @param {number | string} x - The column, or the cell name, for example "B3".
 * @param {number} [y] - The row.
 * @returns {boolean}
 */
export const isReadOnly = function (x, y) {
    const obj = this;

    if (typeof x === 'string' && typeof y === 'undefined') {
        const coords = getCoordsFromCellName(x);

        [x, y] = coords;
    }

    return obj.records[y][/** @type {number} */ (x)].element.classList.contains('readonly');
};
