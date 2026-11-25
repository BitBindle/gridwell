import dispatch from './dispatch.js';
import type jspreadsheet from '../jspreadsheet';
import type { Worksheet } from '../types/internal';

/**
 * Get meta information from cell(s)
 * @param cell - The cell name, for example "B3". Without it, the meta information of all cells.
 * @param key - The key of the information.
 */
export const getMeta = function (this: Worksheet, cell?: string, key?: string): any {
    const obj = this;

    if (!cell) {
        return obj.options.meta;
    } else {
        if (key) {
            return obj.options.meta && obj.options.meta[cell] && obj.options.meta[cell][key] ? obj.options.meta[cell][key] : null;
        } else {
            return obj.options.meta && obj.options.meta[cell] ? obj.options.meta[cell] : null;
        }
    }
};

/**
 * Update meta information
 * @param affectedCells - The new cell names, by old cell name.
 */
export const updateMeta = function (this: Worksheet, affectedCells: Record<string, string>) {
    const obj = this;

    if (obj.options.meta) {
        const newMeta: Record<string, jspreadsheet.MetaInformation> = {};
        const keys = Object.keys(obj.options.meta);
        for (let i = 0; i < keys.length; i++) {
            if (affectedCells[keys[i]]) {
                newMeta[affectedCells[keys[i]]] = obj.options.meta[keys[i]];
            } else {
                newMeta[keys[i]] = obj.options.meta[keys[i]];
            }
        }
        // Update meta information
        obj.options.meta = newMeta;
    }
};

/**
 * Set meta information to cell(s)
 * @param o - The cell name, for example "B3", or the information of many cells: { B3: { key: value } }.
 * @param k - The key, with a cell name.
 * @param v - The value, with a cell name.
 */
export const setMeta = function (this: Worksheet, o: string | Record<string, jspreadsheet.MetaInformation>, k?: string, v?: string) {
    const obj = this;

    if (!obj.options.meta) {
        obj.options.meta = {};
    }

    if (k && v) {
        // Set data value. o is a cell name.
        if (!obj.options.meta[o as string]) {
            obj.options.meta[o as string] = {};
        }
        obj.options.meta[o as string][k] = v;

        dispatch.call(obj, 'onchangemeta', obj, { [o as string]: { [k]: v } });
    } else {
        // Apply that for all cells
        const keys = Object.keys(o);
        for (let i = 0; i < keys.length; i++) {
            if (!obj.options.meta[keys[i]]) {
                obj.options.meta[keys[i]] = {};
            }

            // o is the information by cell name
            const prop = Object.keys((o as Record<string, jspreadsheet.MetaInformation>)[keys[i]]);
            for (let j = 0; j < prop.length; j++) {
                obj.options.meta[keys[i]][prop[j]] = (o as Record<string, jspreadsheet.MetaInformation>)[keys[i]][prop[j]];
            }
        }

        dispatch.call(obj, 'onchangemeta', obj, o);
    }
};
