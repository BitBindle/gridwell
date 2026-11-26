import libraryBase from '../utils/libraryBase.js';
import { getElement, getIndex } from './helpers.js';
import type { TargetEvent } from '../types/internal';

/**
 * The mouse handlers act on the active worksheet, gridwell.current. A press on a worksheet makes it the active
 * worksheet (mouseDownControls), and the other handlers read it once. Each gesture is a method of worksheet.pointer: a
 * press on a header, a row number or a cell, the resize or the move of a column or a row, and the fill with the
 * selection corner.
 */

/** The button that is down: buttons, or button, or which in old browsers. */
function getMouseButton(e: MouseEvent) {
    if (e.buttons) {
        return e.buttons;
    } else if (e.button) {
        return e.button;
    } else {
        return e.which;
    }
}

/** The cell that contains the element: the first element up the tree with data-x and data-y. */
function getCell(element: HTMLElement): HTMLElement | undefined {
    if (element.parentNode) {
        const x = element.getAttribute('data-x');
        const y = element.getAttribute('data-y');
        if (x && y) {
            return element;
        } else {
            return getCell(element.parentNode as HTMLElement);
        }
    }
}

export function mouseUpControls(e: TargetEvent<MouseEvent>) {
    const current = libraryBase.gridwell.current;

    if (current) {
        // Update cell size
        if (current.pointer.resizing) {
            current.pointer.endResize();
        } else if (current.pointer.dragging) {
            current.pointer.endDrag(e);
        } else {
            current.pointer.endFill();
        }
    }

    // Clear any time control
    if (libraryBase.gridwell.timeControl) {
        clearTimeout(libraryBase.gridwell.timeControl);
        libraryBase.gridwell.timeControl = null;
    }

    // Mouse up
    libraryBase.gridwell.isMouseAction = false;
}

export function mouseDownControls(e: TargetEvent<MouseEvent>) {
    e = e || window.event;

    const mouseButton = getMouseButton(e);

    // Get elements
    const jssTable = getElement(e.target);

    // A press on a worksheet makes it the active worksheet. A press outside the spreadsheets leaves no active worksheet.
    const previous = libraryBase.gridwell.current;
    if (jssTable[0]) {
        if (previous !== jssTable[0].jssWorksheet) {
            if (previous) {
                if (previous.editor.isOpen()) {
                    previous.editor.close(previous.editor.cell, true);
                }
                previous.resetSelection();
            }
            libraryBase.gridwell.current = jssTable[0].jssWorksheet;
        }
    } else {
        if (previous) {
            if (previous.editor.isOpen()) {
                previous.editor.close(previous.editor.cell, true);
            }

            if (!e.target.closest('.jss_object')) {
                previous.selection.reset(true);
                libraryBase.gridwell.current = null;
            }
        }
    }

    const current = libraryBase.gridwell.current;

    if (current && mouseButton === 1) {
        if (e.target.classList.contains('jss_selectall')) {
            current.selection.selectAll();
        } else if (e.target.classList.contains('jss_corner')) {
            if (current.options.editable !== false) {
                current.pointer.corner = true;
            }
        } else {
            // Header found
            if (jssTable[1] === 1) {
                current.pointer.pressHeader(e);
            } else {
                current.selection.column = null;
            }

            // Body found
            if (jssTable[1] === 2) {
                if (e.target.classList.contains('jss_row')) {
                    current.pointer.pressRowNumber(e);
                } else {
                    current.pointer.pressCell(e);
                }
            } else {
                current.selection.row = null;
            }

            // Pagination
            if (e.target.classList.contains('jss_page')) {
                current.pointer.pressPage(e);
            }
        }

        if (current.editor.isOpen()) {
            libraryBase.gridwell.isMouseAction = false;
        } else {
            libraryBase.gridwell.isMouseAction = true;
        }
    } else {
        libraryBase.gridwell.isMouseAction = false;
    }
}

// Mouse move controls
export function mouseMoveControls(e: TargetEvent<MouseEvent>) {
    e = e || window.event;

    if (!getMouseButton(e)) {
        libraryBase.gridwell.isMouseAction = false;
    }

    const current = libraryBase.gridwell.current;

    if (current) {
        if (libraryBase.gridwell.isMouseAction === true) {
            // Resizing is ongoing
            if (current.pointer.resizing) {
                current.pointer.moveResize(e);
            } else if (current.pointer.dragging) {
                current.pointer.moveDrag(e);
            }
        } else {
            current.pointer.updateCursor(e);
        }
    }
}

export function mouseOverControls(e: TargetEvent<MouseEvent>) {
    e = e || window.event;

    if (!getMouseButton(e)) {
        libraryBase.gridwell.isMouseAction = false;
    }

    const current = libraryBase.gridwell.current;

    if (current && libraryBase.gridwell.isMouseAction === true) {
        // Get elements
        const jssTable = getElement(e.target);

        if (jssTable[0]) {
            // Avoid cross reference
            if (current !== jssTable[0].jssWorksheet) {
                return false;
            }

            current.pointer.extendSelection(e, jssTable[1]);
        }
    }

    // Clear any time control
    if (libraryBase.gridwell.timeControl) {
        clearTimeout(libraryBase.gridwell.timeControl);
        libraryBase.gridwell.timeControl = null;
    }
}

export function doubleClickControls(e: TargetEvent<MouseEvent>) {
    const current = libraryBase.gridwell.current;

    // Jss is selected
    if (current) {
        // Corner action
        if (e.target.classList.contains('jss_corner')) {
            current.pointer.fillDown();
        } else if (e.target.classList.contains('jss_column_filter')) {
            // Column
            const columnId = parseInt(e.target.getAttribute('data-x'));
            // Open filter
            current.view.openFilter(columnId);
        } else {
            // Get table
            const jssTable = getElement(e.target);

            // Double click over header
            if (jssTable[1] === 1 && current.options.columnSorting !== false) {
                // Check valid column header coords
                const columnId = getIndex(e.target, 'data-x');
                if (columnId != null) {
                    current.orderBy(columnId);
                }
            }

            // Double click over body
            if (jssTable[1] === 2 && current.options.editable !== false) {
                if (!current.editor.isOpen()) {
                    const cell = getCell(e.target);
                    if (cell?.classList.contains('highlight')) {
                        current.editor.open(cell as HTMLTableCellElement, undefined, e);
                    }
                }
            }
        }
    }
}
