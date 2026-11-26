import type { Worksheet } from '../types/internal';

/*
 * The zoom of the grid. The wrapper worksheet.zoomed, inside the scroll container worksheet.content, has the CSS property
 * zoom. getBoundingClientRect, the pointer coordinates, scrollTop, scrollLeft and the size of the content are in screen
 * pixels. offsetTop, offsetHeight, the styles of the elements in the wrapper and the column widths are in grid pixels.
 */

/** The zoom factor of the spreadsheet of a worksheet: 1 at 100 %. A value that is not a number above 0 gives 1. */
export function zoomOf(worksheet: Worksheet) {
    const zoom = Number(worksheet.parent.config.zoom);
    return Number.isFinite(zoom) && zoom > 0 ? zoom : 1;
}

/** A length in screen pixels as a length in grid pixels. */
export function toGrid(worksheet: Worksheet, screen: number) {
    return screen / zoomOf(worksheet);
}

/** A length in grid pixels as a length in screen pixels. */
export function toScreen(worksheet: Worksheet, grid: number) {
    return grid * zoomOf(worksheet);
}

/** Give the wrapper of a worksheet the zoom factor of its spreadsheet. */
export function applyZoom(worksheet: Worksheet) {
    const zoom = zoomOf(worksheet);
    worksheet.zoomed.style.zoom = zoom === 1 ? '' : String(zoom);
    worksheet.zoomed.style.setProperty('--jss-zoom', String(zoom));
}
