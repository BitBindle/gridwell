import { getColumnNameFromId, getIdFromColumnName } from '../utils/internalHelpers.js';
import type { Worksheet } from '../types/internal';

/** A style declaration read and written by property name, for example style['color'] */
type StyleProperties = Record<string, string>;

/**
 * The styles of the cells of a worksheet, kept in the style attribute of each cell.
 *
 * Each worksheet has one, as worksheet.styles. getStyle, setStyle and resetStyle call it.
 */
export class Styles {
    readonly worksheet: Worksheet;

    constructor(worksheet: Worksheet) {
        this.worksheet = worksheet;
    }

    /**
     * Get style information from cell(s)
     * @param cell - The cell name, for example "B3". Without it, the style of all cells.
     * @param key - A CSS property, for example "color". Without it, the whole style.
     */
    get(cell?: string | number[], key?: string): Record<string, string> | string {
        const worksheet = this.worksheet;

        // Cell
        if (!cell) {
            // Control vars
            const data: Record<string, string> = {};

            // Column and row length
            const x = worksheet.options.data[0].length;
            const y = worksheet.options.data.length;

            // Go through the columns to get the data
            for (let j = 0; j < y; j++) {
                for (let i = 0; i < x; i++) {
                    // Value
                    const v = key
                        ? (worksheet.records[j][i].element.style as unknown as StyleProperties)[key]
                        : worksheet.records[j][i].element.getAttribute('style');

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
                ? (worksheet.records[cell[1]][cell[0]].element.style as unknown as StyleProperties)[key]
                : worksheet.records[cell[1]][cell[0]].element.getAttribute('style')!;
        }
    }

    /**
     * Set the style of cell(s)
     * @param o - The cell name, for example "B3", or the styles by cell name.
     * @param k - A CSS property, with a cell name.
     * @param v - The value, with a cell name.
     * @param force - Also style read-only cells, and set a value that the cell already has.
     * @param ignoreHistoryAndEvents - Add no history. The onchangestyle event still occurs.
     */
    set(o: string | Record<string, string | string[]>, k?: string | null, v?: string | null, force?: boolean, ignoreHistoryAndEvents?: boolean) {
        const worksheet = this.worksheet;

        // Lists of 'property:value' entries by cell name, joined into one string per cell below
        const newValue: Record<string, string[][]> = {};
        const oldValue: Record<string, string[][]> = {};

        /**
         * Apply style
         * @param cellId - The cell name.
         * @param key - A CSS property.
         */
        function applyStyle(cellId: string, key: string, value: string) {
            // Position
            const cell = getIdFromColumnName(cellId, true);

            if (worksheet.records[cell[1]]?.[cell[0]] && (worksheet.records[cell[1]][cell[0]].element.classList.contains('readonly') === false || force)) {
                // Current value
                const currentValue = (worksheet.records[cell[1]][cell[0]].element.style as unknown as StyleProperties)[key];

                // Change layout
                if (currentValue === value && !force) {
                    value = '';
                    (worksheet.records[cell[1]][cell[0]].element.style as unknown as StyleProperties)[key] = '';
                } else {
                    (worksheet.records[cell[1]][cell[0]].element.style as unknown as StyleProperties)[key] = value;
                }

                // History
                if (!oldValue[cellId]) {
                    oldValue[cellId] = [];
                }
                if (!newValue[cellId]) {
                    newValue[cellId] = [];
                }

                oldValue[cellId].push([`${key}:${currentValue}`]);
                newValue[cellId].push([`${key}:${value}`]);
            }
        }

        if (k && v) {
            // Get object from string
            if (typeof o === 'string') {
                applyStyle(o, k, v);
            }
        } else {
            const keys = Object.keys(o);
            for (const cellName of keys) {
                let style: string | string[] = (o as Record<string, string | string[]>)[cellName];
                if (typeof style === 'string') {
                    style = style.split(';');
                }
                for (const declaration of style) {
                    // Split into a new variable, so that the styles the application gave stay as they are
                    let property: string | string[] = declaration;
                    if (typeof property === 'string') {
                        // Split at the first colon only, because a value can have colons, for example url(http://...)
                        property = property.split(/:(.*)/);
                    }
                    // Apply value
                    if (property[0].trim()) {
                        applyStyle(cellName, property[0].trim(), property[1]);
                    }
                }
            }
        }

        // The lists become strings in place
        let keys = Object.keys(oldValue);
        for (const cellName of keys) {
            (oldValue as Record<string, unknown>)[cellName] = oldValue[cellName].join(';');
        }
        keys = Object.keys(newValue);
        for (const cellName of keys) {
            (newValue as Record<string, unknown>)[cellName] = newValue[cellName].join(';');
        }

        if (!ignoreHistoryAndEvents) {
            // Keeping history of changes
            worksheet.history.add({
                action: 'setStyle',
                oldValue: oldValue as unknown as Record<string, string>,
                newValue: newValue as unknown as Record<string, string>,
            });
        }

        worksheet.dispatch('onchangestyle', worksheet, newValue as unknown as Record<string, string>);
    }

    /**
     * Remove the style of cells.
     * @param o - The cells, as the keys.
     * @param ignoreHistoryAndEvents - Add no history. The onchangestyle event still occurs.
     */
    reset(o: Record<string, string | string[]>, ignoreHistoryAndEvents?: boolean) {
        const worksheet = this.worksheet;

        for (const key of Object.keys(o)) {
            // Position
            const cell = getIdFromColumnName(key, true);
            if (worksheet.records[cell[1]]?.[cell[0]]) {
                worksheet.records[cell[1]][cell[0]].element.setAttribute('style', '');
            }
        }
        worksheet.setStyle(o, null, null, false, ignoreHistoryAndEvents);
    }
}
