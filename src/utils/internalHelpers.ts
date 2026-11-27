import { getColumnName } from './helpers.js';

/**
 * Make a copy of the array with the items of arr inserted at idx. The callers insert into the array or at its end.
 * @param o - The array.
 * @param idx - The position.
 * @param arr - The items to insert.
 */
export function injectArray<T>(o: T[], idx: number, arr: T[]): T[] {
    return o.slice(0, idx).concat(arr, o.slice(idx));
}

/**
 * Get the coordinates of a cell name, for example [1, 2] for "B3". When the name does not start with letters, it returns
 * the name, as before.
 * @param id - The cell name.
 */
export function getIdFromColumnName(id: string): number[] {
    // Get the letters
    const t = /^[a-zA-Z]+/.exec(id);

    if (!t) {
        // The callers give cell names, which start with letters
        return id as unknown as number[];
    }

    // Base 26 calculation. The letters give 1 or more, so the code starts on zero.
    let code = 0;
    for (let i = 0; i < t[0].length; i++) {
        code += parseInt(t[0].charCodeAt(i) - 64) * Math.pow(26, t[0].length - 1 - i);
    }
    code--;

    // Number
    // parseInt converts the match to its text, or null to NaN
    let number = parseInt(/[0-9]+$/.exec(id) as unknown as string | null);
    if (number > 0) {
        number--;
    }

    return [code, number];
}

/**
 * Convert the coordinates of a cell to its name
 * @param cellId - The coordinates, for example [1, 2].
 * @returns The cell name, for example "B3".
 */
export function getColumnNameFromId(cellId: (string | number)[]): string {
    return getColumnName(parseInt(cellId[0])) + (parseInt(cellId[1]) + 1);
}

/**
 * The new position of a row or column after a move of one row or column.
 * @param position - The old position.
 * @param from - The row or column that moved.
 * @param to - Its new position.
 */
export function movedIndex(position: number, from: number, to: number) {
    if (position === from) {
        return to;
    }
    if (from < to && position > from && position <= to) {
        return position - 1;
    }
    if (from > to && position >= to && position < from) {
        return position + 1;
    }
    return position;
}
