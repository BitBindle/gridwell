// @ts-check
import { getColumnName } from './helpers.js';

/**
 * Make a copy of the array with the items of arr inserted at idx. When idx is past the end, the gap is filled with undefined.
 * @template T
 * @param {T[]} o - The array.
 * @param {number} idx - The position.
 * @param {T[]} arr - The items to insert.
 * @returns {(T | undefined)[]}
 */
export const injectArray = function (o, idx, arr) {
    if (idx <= o.length) {
        return o.slice(0, idx).concat(arr).concat(o.slice(idx));
    }

    const array = o.slice(0, o.length);

    while (idx > array.length) {
        array.push(undefined);
    }

    return array.concat(arr);
};

/**
 * Convert excel like column to jss id
 *
 * With arr true, it returns the coordinates, for example [1, 2] for "B3". Otherwise it returns the id, for example "1-2".
 * When id does not start with letters, it returns id.
 *
 * @type {{ (id: string, arr: true): number[]; (id: string, arr?: boolean): string }}
 */
export const getIdFromColumnName = /** @param {any} id @param {boolean} [arr] @returns {any} */ function (id, arr) {
    // Get the letters
    const t = /^[a-zA-Z]+/.exec(id);

    if (t) {
        // Base 26 calculation
        let code = 0;
        for (let i = 0; i < t[0].length; i++) {
            code += parseInt(t[0].charCodeAt(i) - 64) * Math.pow(26, t[0].length - 1 - i);
        }
        code--;
        // Make sure jss starts on zero
        if (code < 0) {
            code = 0;
        }

        // Number
        // parseInt converts the match to its text, or null to NaN
        let number = parseInt(/** @type {any} */ (/[0-9]+$/.exec(id)));
        if (number > 0) {
            number--;
        }

        if (arr == true) {
            id = [code, number];
        } else {
            id = code + '-' + number;
        }
    }

    return id;
};

/**
 * Convert jss id to excel like column name
 *
 * @param {string | (string | number)[]} cellId - An id, for example "1-2", or the coordinates, for example [1, 2].
 * @returns {string} The cell name, for example "B3".
 */
export const getColumnNameFromId = function (cellId) {
    if (!Array.isArray(cellId)) {
        cellId = cellId.split('-');
    }

    return getColumnName(parseInt(cellId[0])) + (parseInt(cellId[1]) + 1);
};
