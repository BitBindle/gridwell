import { getColumnName } from './helpers.js';

/**
 * Make a copy of the array with the items of arr inserted at idx. When idx is past the end, the gap is filled with undefined.
 * @param o - The array.
 * @param idx - The position.
 * @param arr - The items to insert.
 */
export function injectArray<T>(o: T[], idx: number, arr: T[]): T[] {
    if (idx <= o.length) {
        return o.slice(0, idx).concat(arr).concat(o.slice(idx));
    }

    const array: (T | undefined)[] = o.slice(0, o.length);

    while (idx > array.length) {
        array.push(undefined);
    }

    // A position after the end leaves undefined items between. The callers insert into the array or at its end.
    return array.concat(arr) as T[];
}

/**
 * Convert excel like column to jss id
 *
 * With arr true, it returns the coordinates, for example [1, 2] for "B3". Otherwise it returns the id, for example "1-2".
 * When id does not start with letters, it returns id.
 */
export function getIdFromColumnName(id: string, arr: true): number[];
export function getIdFromColumnName(id: string, arr?: boolean): string;
// id becomes the result, so it can be a list
export function getIdFromColumnName(id: string | number[], arr?: boolean) {
    // Get the letters
    const t = /^[a-zA-Z]+/.exec(id as string);

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
        let number = parseInt(/[0-9]+$/.exec(id as string) as unknown as string | null);
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
}

/**
 * Convert jss id to excel like column name
 * @param cellId - An id, for example "1-2", or the coordinates, for example [1, 2].
 * @returns The cell name, for example "B3".
 */
export function getColumnNameFromId(cellId: string | (string | number)[]): string {
    if (!Array.isArray(cellId)) {
        cellId = cellId.split('-');
    }

    return getColumnName(parseInt(cellId[0])) + (parseInt(cellId[1]) + 1);
}
