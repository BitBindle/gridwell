import type { DragColumnInfo, DragRowInfo, ResizeColumnInfo, ResizeRowInfo } from '../jspreadsheet';

/**
 * The state of a mouse gesture on a worksheet: the column or row that the mouse resizes or moves, the drag of the
 * selection corner, and the element whose mouse cursor shows a resize.
 *
 * Each worksheet has one, as worksheet.pointer. The mouse handlers in src/events/mouse.ts set it when a button goes down
 * and clear it when it goes up.
 */
export class Pointer {
    /**
     * The column or row that the mouse resizes, or null. Each member says that the other index is missing, so that a
     * test of column or row tells the type check which one it is.
     */
    resizing: null | (ResizeColumnInfo & { row?: undefined }) | (ResizeRowInfo & { column?: undefined }) = null;

    /** The column or row that the mouse moves, or null. */
    dragging: null | (DragColumnInfo & { row?: undefined }) | (DragRowInfo & { column?: undefined }) = null;

    /** The header or row element that shows a resize cursor, or null. */
    cursor: HTMLElement | null = null;

    /** True while the mouse drags the selection corner to fill cells. */
    corner = false;
}
