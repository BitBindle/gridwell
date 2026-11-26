import type jspreadsheet from '../jspreadsheet';
import type { Worksheet } from '../types/internal';

/**
 * The meta information of the cells of a worksheet: values that the application keeps by cell name, in options.meta.
 *
 * Each worksheet has one, as worksheet.meta. getMeta and setMeta call it, and the references (src/utils/references.ts)
 * rename its cells when rows or columns move.
 */
export class Meta {
    readonly worksheet: Worksheet;

    constructor(worksheet: Worksheet) {
        this.worksheet = worksheet;
    }

    /**
     * Get meta information from cell(s)
     * @param cell - The cell name, for example "B3". Without it, the meta information of all cells.
     * @param key - The key of the information.
     */
    get(cell?: string, key?: string) {
        const worksheet = this.worksheet;

        if (!cell) {
            return worksheet.options.meta;
        } else {
            if (key) {
                return worksheet.options.meta && worksheet.options.meta[cell] && worksheet.options.meta[cell][key] ? worksheet.options.meta[cell][key] : null;
            } else {
                return worksheet.options.meta && worksheet.options.meta[cell] ? worksheet.options.meta[cell] : null;
            }
        }
    }

    /**
     * Update meta information
     * @param affectedCells - The new cell names, by old cell name.
     */
    update(affectedCells: Record<string, string>) {
        const worksheet = this.worksheet;

        if (worksheet.options.meta) {
            const newMeta: Record<string, jspreadsheet.MetaInformation> = {};
            const keys = Object.keys(worksheet.options.meta);
            for (let i = 0; i < keys.length; i++) {
                if (affectedCells[keys[i]]) {
                    newMeta[affectedCells[keys[i]]] = worksheet.options.meta[keys[i]];
                } else {
                    newMeta[keys[i]] = worksheet.options.meta[keys[i]];
                }
            }
            // Update meta information
            worksheet.options.meta = newMeta;
        }
    }

    /**
     * Set meta information to cell(s)
     * @param o - The cell name, for example "B3", or the information of many cells: { B3: { key: value } }.
     * @param k - The key, with a cell name.
     * @param v - The value, with a cell name.
     */
    set(o: string | Record<string, jspreadsheet.MetaInformation>, k?: string, v?: string) {
        const worksheet = this.worksheet;

        if (!worksheet.options.meta) {
            worksheet.options.meta = {};
        }

        if (k && v) {
            // Set data value. o is a cell name.
            if (!worksheet.options.meta[o as string]) {
                worksheet.options.meta[o as string] = {};
            }
            worksheet.options.meta[o as string][k] = v;

            worksheet.dispatch('onchangemeta', worksheet, { [o as string]: { [k]: v } });
        } else {
            // Apply that for all cells
            const keys = Object.keys(o);
            for (let i = 0; i < keys.length; i++) {
                if (!worksheet.options.meta[keys[i]]) {
                    worksheet.options.meta[keys[i]] = {};
                }

                // o is the information by cell name
                const prop = Object.keys((o as Record<string, jspreadsheet.MetaInformation>)[keys[i]]);
                for (let j = 0; j < prop.length; j++) {
                    worksheet.options.meta[keys[i]][prop[j]] = (o as Record<string, jspreadsheet.MetaInformation>)[keys[i]][prop[j]];
                }
            }

            worksheet.dispatch('onchangemeta', worksheet, o as Record<string, jspreadsheet.MetaInformation>);
        }
    }
}
