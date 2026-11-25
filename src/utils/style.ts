import dispatch from './dispatch.js';
import { getColumnNameFromId, getIdFromColumnName } from './internalHelpers.js';
import { setHistory } from './history.js';
import type { Worksheet } from '../types/internal';

/** A style declaration read and written by property name, for example style['color'] */
type StyleProperties = Record<string, string>;

/**
 * Get style information from cell(s)
 * @param cell - The cell name, for example "B3". Without it, the style of all cells.
 * @param key - A CSS property, for example "color". Without it, the whole style.
 */
export const getStyle = function (this: Worksheet, cell?: string | number[], key?: string): Record<string, string> | string {
    const obj = this;

    // Cell
    if (!cell) {
        // Control vars
        const data: Record<string, string> = {};

        // Column and row length
        const x = obj.options.data[0].length;
        const y = obj.options.data.length;

        // Go through the columns to get the data
        for (let j = 0; j < y; j++) {
            for (let i = 0; i < x; i++) {
                // Value
                const v = key ? (obj.records[j][i].element.style as unknown as StyleProperties)[key] : obj.records[j][i].element.getAttribute('style');

                // Any meta data for this column?
                if (v) {
                    // Column name
                    const k = getColumnNameFromId([i, j]);
                    // Value
                    data[k] = v;
                }
            }
        }

        return data;
    } else {
        cell = getIdFromColumnName(cell as string, true);

        // Each cell gets a style attribute when it is created (the text alignment)
        return key
            ? (obj.records[cell[1]][cell[0]].element.style as unknown as StyleProperties)[key]
            : obj.records[cell[1]][cell[0]].element.getAttribute('style')!;
    }
};

/**
 * Set the style of cell(s)
 * @param o - The cell name, for example "B3", or the styles by cell name.
 * @param k - A CSS property, with a cell name.
 * @param v - The value, with a cell name.
 * @param force - Also style read-only cells, and set a value that the cell already has.
 * @param ignoreHistoryAndEvents - Add no history. The onchangestyle event still occurs.
 */
export const setStyle = function (
    this: Worksheet,
    o: string | Record<string, string | string[]>,
    k?: string | null,
    v?: string | null,
    force?: boolean,
    ignoreHistoryAndEvents?: boolean
) {
    const obj = this;

    // Lists of 'property:value' entries by cell name, joined into one string per cell below
    const newValue: Record<string, string[][]> = {};
    const oldValue: Record<string, string[][]> = {};

    /**
     * Apply style
     * @param cellId - The cell name.
     * @param key - A CSS property.
     */
    const applyStyle = function (cellId: string, key: string, value: string) {
        // Position
        const cell = getIdFromColumnName(cellId, true);

        if (obj.records[cell[1]] && obj.records[cell[1]][cell[0]] && (obj.records[cell[1]][cell[0]].element.classList.contains('readonly') == false || force)) {
            // Current value
            const currentValue = (obj.records[cell[1]][cell[0]].element.style as unknown as StyleProperties)[key];

            // Change layout
            if (currentValue == value && !force) {
                value = '';
                (obj.records[cell[1]][cell[0]].element.style as unknown as StyleProperties)[key] = '';
            } else {
                (obj.records[cell[1]][cell[0]].element.style as unknown as StyleProperties)[key] = value;
            }

            // History
            if (!oldValue[cellId]) {
                oldValue[cellId] = [];
            }
            if (!newValue[cellId]) {
                newValue[cellId] = [];
            }

            oldValue[cellId].push([key + ':' + currentValue]);
            newValue[cellId].push([key + ':' + value]);
        }
    };

    if (k && v) {
        // Get object from string
        if (typeof o == 'string') {
            applyStyle(o, k, v);
        }
    } else {
        const keys = Object.keys(o);
        for (let i = 0; i < keys.length; i++) {
            let style: string | string[] = (o as Record<string, string | string[]>)[keys[i]];
            if (typeof style == 'string') {
                style = style.split(';');
            }
            for (let j = 0; j < style.length; j++) {
                // Split into a new variable, so that the styles the application gave stay as they are
                let property: string | string[] = style[j];
                if (typeof property == 'string') {
                    // Split at the first colon only, because a value can have colons, for example url(http://...)
                    property = property.split(/:(.*)/);
                }
                // Apply value
                if (property[0].trim()) {
                    applyStyle(keys[i], property[0].trim(), property[1]);
                }
            }
        }
    }

    // The lists become strings in place
    let keys = Object.keys(oldValue);
    for (let i = 0; i < keys.length; i++) {
        (oldValue as Record<string, unknown>)[keys[i]] = oldValue[keys[i]].join(';');
    }
    keys = Object.keys(newValue);
    for (let i = 0; i < keys.length; i++) {
        (newValue as Record<string, unknown>)[keys[i]] = newValue[keys[i]].join(';');
    }

    if (!ignoreHistoryAndEvents) {
        // Keeping history of changes
        setHistory.call(obj, {
            action: 'setStyle',
            oldValue: oldValue as unknown as Record<string, string>,
            newValue: newValue as unknown as Record<string, string>,
        });
    }

    dispatch.call(obj, 'onchangestyle', obj, newValue);
};

/**
 * Remove the style of cells.
 * @param o - The cells, as the keys.
 * @param ignoreHistoryAndEvents - Add no history. The onchangestyle event still occurs.
 */
export const resetStyle = function (this: Worksheet, o: Record<string, any>, ignoreHistoryAndEvents?: boolean) {
    const obj = this;

    const keys = Object.keys(o);
    for (let i = 0; i < keys.length; i++) {
        // Position
        const cell = getIdFromColumnName(keys[i], true);
        if (obj.records[cell[1]] && obj.records[cell[1]][cell[0]]) {
            obj.records[cell[1]][cell[0]].element.setAttribute('style', '');
        }
    }
    obj.setStyle(o, null, null, false, ignoreHistoryAndEvents);
};
