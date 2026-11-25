import { getCoordsFromCellName } from '../utils/helpers.js';
import { getIdFromColumnName } from '../utils/internalHelpers.js';
import type { Worksheet } from '../types/internal';

/**
 * Finding the cells of a worksheet: the td of a cell, its label, and whether it is read-only.
 *
 * Each worksheet has one, as worksheet.cells. getCell, getCellFromCoords, getLabel, setReadOnly and isReadOnly call it.
 */
export class Cells {
    readonly worksheet: Worksheet;

    constructor(worksheet: Worksheet) {
        this.worksheet = worksheet;
    }

    /**
     * Make a cell read-only, or editable again.
     * @param cell - The cell name, for example "B3", or the td of the cell.
     * @param state - Whether the cell is read-only.
     */
    setReadOnly(cell: string | HTMLElement, state: boolean) {
        const obj = this.worksheet;

        let record;

        if (typeof cell === 'string') {
            const coords = getCoordsFromCellName(cell);

            record = obj.records[coords![1]!][coords![0]];
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
    }

    /**
     * Check if a cell is read-only.
     * @param x - The column, or the cell name, for example "B3".
     * @param y - The row.
     */
    isReadOnly(x: number | string, y?: number): boolean {
        const obj = this.worksheet;

        if (typeof x === 'string' && typeof y === 'undefined') {
            const coords = getCoordsFromCellName(x);

            [x, y] = coords as [number, number];
        }

        return obj.records[y!][x as number].element.classList.contains('readonly');
    }

    /**
     * Get the cell object
     *
     * @param object cell
     * @return string value
     */
    get(x: number, y: number) {
        const obj = this.worksheet;

        if (typeof x === 'string') {
            // Convert in case name is excel liked ex. A10, BB92
            const cell = getIdFromColumnName(x, true);

            x = cell[0];
            y = cell[1];
        }

        return obj.records[y][x].element;
    }

    /**
     * Get the cell object from coords
     *
     * @param object cell
     * @return string value
     */
    getFromCoords(x: number, y: number) {
        const obj = this.worksheet;

        return obj.records[y][x].element;
    }

    /**
     * Get label
     *
     * @param object cell
     * @return string value
     */
    getLabel(x: number, y: number) {
        const obj = this.worksheet;

        if (typeof x === 'string') {
            // Convert in case name is excel liked ex. A10, BB92
            const cell = getIdFromColumnName(x, true);

            x = cell[0];
            y = cell[1];
        }

        return obj.records[y][x].element.innerHTML;
    }
}
