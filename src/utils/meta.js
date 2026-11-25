// @ts-check
import dispatch from './dispatch.js';
/** @import { Worksheet } from '../types/internal' */

/**
 * Get meta information from cell(s)
 * @this {Worksheet}
 * @param {string} [cell] - The cell name, for example "B3". Without it, the meta information of all cells.
 * @param {string} [key] - The key of the information.
 * @returns {any}
 */
export const getMeta = function (cell, key) {
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
 * @this {Worksheet}
 * @param {Record<string, string>} affectedCells - The new cell names, by old cell name.
 */
export const updateMeta = function (affectedCells) {
    const obj = this;

    if (obj.options.meta) {
        /** @type {Record<string, any>} */
        const newMeta = {};
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
 * @this {Worksheet}
 * @param {any} o - The cell name, for example "B3", or the information of many cells: { B3: { key: value } }.
 * @param {string} [k] - The key, with a cell name.
 * @param {any} [v] - The value, with a cell name.
 */
export const setMeta = function (o, k, v) {
    const obj = this;

    if (!obj.options.meta) {
        obj.options.meta = {};
    }

    if (k && v) {
        // Set data value
        if (!obj.options.meta[o]) {
            obj.options.meta[o] = {};
        }
        obj.options.meta[o][k] = v;

        dispatch.call(obj, 'onchangemeta', obj, { [o]: { [k]: v } });
    } else {
        // Apply that for all cells
        const keys = Object.keys(o);
        for (let i = 0; i < keys.length; i++) {
            if (!obj.options.meta[keys[i]]) {
                obj.options.meta[keys[i]] = {};
            }

            const prop = Object.keys(o[keys[i]]);
            for (let j = 0; j < prop.length; j++) {
                obj.options.meta[keys[i]][prop[j]] = o[keys[i]][prop[j]];
            }
        }

        dispatch.call(obj, 'onchangemeta', obj, o);
    }
};
