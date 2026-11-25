import type { WorksheetState } from '../types/internal';

/** The column and the row of a corner of the selection, then of the other corner: [x1, y1, x2, y2]. */
type Range = [number, number, number, number];

/**
 * The selection of a worksheet: the selected range, the highlighted cells, the cells that a drag of the corner fills,
 * and the cell that has the keyboard focus.
 *
 * Each worksheet has one, as worksheet.selection. Changing the selection updates the classes, the ARIA attributes and
 * the toolbar, and dispatches the events, so the functions in src/utils/selection.ts and src/utils/accessibility.ts do
 * that, and keep the state here.
 */
export class Selection {
    /**
     * The selected range, or null when nothing is selected. The column [0] and row [1] of the cell where the selection
     * starts, and the column [2] and row [3] of the cell where it ends. The end can be left of or above the start.
     */
    range: Range | null = null;

    /** The selected range from left to right and top to bottom, or null before the first selection. */
    bounds: Range | null = null;

    /** The records of the selected cells. */
    highlighted: WorksheetState['records'][number] = [];

    /** The cells that a drag of the selection corner fills, while the drag goes on. */
    fill: HTMLTableCellElement[] = [];

    /** The cell that has the roving tabindex (the tab stop of the grid), or null before the grid has one. */
    activeCell: HTMLTableCellElement | null = null;

    /** Whether the keyboard focus is in the grid or in its editor. */
    focusInGrid = false;

    /** The column whose header a press selected first, while the selection is of whole columns, or null. */
    column: number | null = null;

    /** The row whose number a press selected first, while the selection is of whole rows, or null. */
    row: number | null = null;

    /** Remove the fill border from the cells that a drag of the corner was going to fill, and forget them. */
    clearFill() {
        // Remove current selection
        for (let i = 0; i < this.fill.length; i++) {
            this.fill[i].classList.remove('selection');
            this.fill[i].classList.remove('selection-left');
            this.fill[i].classList.remove('selection-right');
            this.fill[i].classList.remove('selection-top');
            this.fill[i].classList.remove('selection-bottom');
        }

        this.fill = [];
    }
}
