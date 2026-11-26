import { getFreezeWidth, getIndexWidth } from './freeze.js';
import type { SelectedWorksheet, Worksheet, WorksheetMethod } from '../types/internal';

/**
 * Whether the worksheet has a selection
 */
const hasSelection = function (obj: Worksheet): obj is SelectedWorksheet {
    return obj.selection.range ? true : false;
};

/**
 * Move the selection from a cell that a merged range hides to the first cell of the range
 */
const toFirstCell = function (this: SelectedWorksheet) {
    const obj = this;

    const element = obj.records[obj.selection.range[1]][obj.selection.range[0]].element;

    if (element.getAttribute('data-merged')) {
        obj.selection.range[0] = parseInt(element.getAttribute('data-x'));
        obj.selection.range[1] = parseInt(element.getAttribute('data-y'));
    }
};

const upGet = function (this: Worksheet, x: number, y: number) {
    const obj = this;

    x = parseInt(x);
    y = parseInt(y);
    for (let j = y - 1; j >= 0; j--) {
        if (obj.records[j][x].element.style.display != 'none' && obj.rows[j].element.style.display != 'none') {
            if (obj.records[j][x].element.getAttribute('data-merged')) {
                if (obj.records[j][x].element == obj.records[y][x].element) {
                    continue;
                }
            }
            y = j;
            break;
        }
    }

    return y;
};

/** The corner of the selection that moves: 0 for the cell in selection.range[0] and [1], 1 for the cell in [2] and [3]. */
type Corner = 0 | 1;

/** 0 moves to the last visible cell in that direction (with Ctrl), 1 moves to the next visible cell. */
type Step = 0 | 1;

const upVisible = function (this: SelectedWorksheet, group: Corner, direction: Step) {
    const obj = this;

    let x, y;

    if (group == 0) {
        x = parseInt(obj.selection.range[0]);
        y = parseInt(obj.selection.range[1]);
    } else {
        x = parseInt(obj.selection.range[2]);
        y = parseInt(obj.selection.range[3]);
    }

    if (direction == 0) {
        for (let j = 0; j < y; j++) {
            if (obj.records[j][x].element.style.display != 'none' && obj.rows[j].element.style.display != 'none') {
                y = j;
                break;
            }
        }
    } else {
        y = upGet.call(obj, x, y);
    }

    if (group == 0) {
        obj.selection.range[0] = x;
        obj.selection.range[1] = y;
    } else {
        obj.selection.range[2] = x;
        obj.selection.range[3] = y;
    }
};

export const up: WorksheetMethod<'up'> = function (shiftKey, ctrlKey) {
    const obj = this;

    // Without a selection, there is nothing to move
    if (!hasSelection(obj)) {
        return;
    }

    if (shiftKey) {
        if (obj.selection.range[3] > 0) {
            upVisible.call(obj, 1, ctrlKey ? 0 : 1);
        }
    } else {
        if (obj.selection.range[1] > 0) {
            upVisible.call(obj, 0, ctrlKey ? 0 : 1);
        }
        toFirstCell.call(obj);
        obj.selection.range[2] = obj.selection.range[0];
        obj.selection.range[3] = obj.selection.range[1];
    }

    // Update selection
    obj.updateSelectionFromCoords(obj.selection.range[0], obj.selection.range[1], obj.selection.range[2], obj.selection.range[3]);

    // Change page
    if (obj.options.lazyLoading == true) {
        if (obj.selection.range[1] == 0 || obj.selection.range[3] == 0) {
            obj.view.loadPage(0);
            obj.updateSelectionFromCoords(obj.selection.range[0], obj.selection.range[1], obj.selection.range[2], obj.selection.range[3]);
        } else {
            if (obj.view.loadValidation()) {
                obj.updateSelectionFromCoords(obj.selection.range[0], obj.selection.range[1], obj.selection.range[2], obj.selection.range[3]);
            } else {
                const item = parseInt((obj.tbody.firstChild as HTMLElement).getAttribute('data-y'));
                if (obj.selection.range[1] - item < 30) {
                    obj.view.loadUp();
                    obj.updateSelectionFromCoords(obj.selection.range[0], obj.selection.range[1], obj.selection.range[2], obj.selection.range[3]);
                }
            }
        }
    } else if (obj.options.pagination! > 0) {
        const pageNumber = obj.whichPage(obj.selection.range[3]);
        if (pageNumber != obj.view.pageNumber) {
            obj.page(pageNumber);
        }
    }

    updateScroll.call(obj, 1);
};

export const rightGet = function (this: Worksheet, x: number, y: number) {
    const obj = this;

    x = parseInt(x);
    y = parseInt(y);

    for (let i = x + 1; i < obj.headers.length; i++) {
        if (obj.records[y][i].element.style.display != 'none') {
            if (obj.records[y][i].element.getAttribute('data-merged')) {
                if (obj.records[y][i].element == obj.records[y][x].element) {
                    continue;
                }
            }
            x = i;
            break;
        }
    }

    return x;
};

const rightVisible = function (this: SelectedWorksheet, group: Corner, direction: Step) {
    const obj = this;

    let x, y;

    if (group == 0) {
        x = parseInt(obj.selection.range[0]);
        y = parseInt(obj.selection.range[1]);
    } else {
        x = parseInt(obj.selection.range[2]);
        y = parseInt(obj.selection.range[3]);
    }

    if (direction == 0) {
        for (let i = obj.headers.length - 1; i > x; i--) {
            if (obj.records[y][i].element.style.display != 'none') {
                x = i;
                break;
            }
        }
    } else {
        x = rightGet.call(obj, x, y);
    }

    if (group == 0) {
        obj.selection.range[0] = x;
        obj.selection.range[1] = y;
    } else {
        obj.selection.range[2] = x;
        obj.selection.range[3] = y;
    }
};

export const right: WorksheetMethod<'right'> = function (shiftKey, ctrlKey) {
    const obj = this;

    // Without a selection, there is nothing to move
    if (!hasSelection(obj)) {
        return;
    }

    if (shiftKey) {
        if (obj.selection.range[2] < obj.headers.length - 1) {
            rightVisible.call(obj, 1, ctrlKey ? 0 : 1);
        }
    } else {
        if (obj.selection.range[0] < obj.headers.length - 1) {
            rightVisible.call(obj, 0, ctrlKey ? 0 : 1);
        }
        toFirstCell.call(obj);
        obj.selection.range[2] = obj.selection.range[0];
        obj.selection.range[3] = obj.selection.range[1];
    }

    obj.updateSelectionFromCoords(obj.selection.range[0], obj.selection.range[1], obj.selection.range[2], obj.selection.range[3]);
    updateScroll.call(obj, 2);
};

export const downGet = function (this: Worksheet, x: number, y: number) {
    const obj = this;

    x = parseInt(x);
    y = parseInt(y);
    for (let j = y + 1; j < obj.rows.length; j++) {
        if (obj.records[j][x].element.style.display != 'none' && obj.rows[j].element.style.display != 'none') {
            if (obj.records[j][x].element.getAttribute('data-merged')) {
                if (obj.records[j][x].element == obj.records[y][x].element) {
                    continue;
                }
            }
            y = j;
            break;
        }
    }

    return y;
};

const downVisible = function (this: SelectedWorksheet, group: Corner, direction: Step) {
    const obj = this;

    let x, y;

    if (group == 0) {
        x = parseInt(obj.selection.range[0]);
        y = parseInt(obj.selection.range[1]);
    } else {
        x = parseInt(obj.selection.range[2]);
        y = parseInt(obj.selection.range[3]);
    }

    if (direction == 0) {
        for (let j = obj.rows.length - 1; j > y; j--) {
            if (obj.records[j][x].element.style.display != 'none' && obj.rows[j].element.style.display != 'none') {
                y = j;
                break;
            }
        }
    } else {
        y = downGet.call(obj, x, y);
    }

    if (group == 0) {
        obj.selection.range[0] = x;
        obj.selection.range[1] = y;
    } else {
        obj.selection.range[2] = x;
        obj.selection.range[3] = y;
    }
};

export const down: WorksheetMethod<'down'> = function (shiftKey, ctrlKey) {
    const obj = this;

    // Without a selection, there is nothing to move
    if (!hasSelection(obj)) {
        return;
    }

    if (shiftKey) {
        if (obj.selection.range[3] < obj.records.length - 1) {
            downVisible.call(obj, 1, ctrlKey ? 0 : 1);
        }
    } else {
        if (obj.selection.range[1] < obj.records.length - 1) {
            downVisible.call(obj, 0, ctrlKey ? 0 : 1);
        }
        toFirstCell.call(obj);
        obj.selection.range[2] = obj.selection.range[0];
        obj.selection.range[3] = obj.selection.range[1];
    }

    obj.updateSelectionFromCoords(obj.selection.range[0], obj.selection.range[1], obj.selection.range[2], obj.selection.range[3]);

    // Change page
    if (obj.options.lazyLoading == true) {
        if (obj.selection.range[1] == obj.records.length - 1 || obj.selection.range[3] == obj.records.length - 1) {
            obj.view.loadPage(-1);
            obj.updateSelectionFromCoords(obj.selection.range[0], obj.selection.range[1], obj.selection.range[2], obj.selection.range[3]);
        } else {
            if (obj.view.loadValidation()) {
                obj.updateSelectionFromCoords(obj.selection.range[0], obj.selection.range[1], obj.selection.range[2], obj.selection.range[3]);
            } else {
                const item = parseInt((obj.tbody.lastChild as HTMLElement).getAttribute('data-y'));
                if (item - obj.selection.range[3] < 30) {
                    obj.view.loadDown();
                    obj.updateSelectionFromCoords(obj.selection.range[0], obj.selection.range[1], obj.selection.range[2], obj.selection.range[3]);
                }
            }
        }
    } else if (obj.options.pagination! > 0) {
        const pageNumber = obj.whichPage(obj.selection.range[3]);
        if (pageNumber != obj.view.pageNumber) {
            obj.page(pageNumber);
        }
    }

    updateScroll.call(obj, 3);
};

const leftGet = function (this: Worksheet, x: number, y: number) {
    const obj = this;

    x = parseInt(x);
    y = parseInt(y);
    for (let i = x - 1; i >= 0; i--) {
        if (obj.records[y][i].element.style.display != 'none') {
            if (obj.records[y][i].element.getAttribute('data-merged')) {
                if (obj.records[y][i].element == obj.records[y][x].element) {
                    continue;
                }
            }
            x = i;
            break;
        }
    }

    return x;
};

const leftVisible = function (this: SelectedWorksheet, group: Corner, direction: Step) {
    const obj = this;

    let x, y;

    if (group == 0) {
        x = parseInt(obj.selection.range[0]);
        y = parseInt(obj.selection.range[1]);
    } else {
        x = parseInt(obj.selection.range[2]);
        y = parseInt(obj.selection.range[3]);
    }

    if (direction == 0) {
        for (let i = 0; i < x; i++) {
            if (obj.records[y][i].element.style.display != 'none') {
                x = i;
                break;
            }
        }
    } else {
        x = leftGet.call(obj, x, y);
    }

    if (group == 0) {
        obj.selection.range[0] = x;
        obj.selection.range[1] = y;
    } else {
        obj.selection.range[2] = x;
        obj.selection.range[3] = y;
    }
};

export const left: WorksheetMethod<'left'> = function (shiftKey, ctrlKey) {
    const obj = this;

    // Without a selection, there is nothing to move
    if (!hasSelection(obj)) {
        return;
    }

    if (shiftKey) {
        if (obj.selection.range[2] > 0) {
            leftVisible.call(obj, 1, ctrlKey ? 0 : 1);
        }
    } else {
        if (obj.selection.range[0] > 0) {
            leftVisible.call(obj, 0, ctrlKey ? 0 : 1);
        }
        toFirstCell.call(obj);
        obj.selection.range[2] = obj.selection.range[0];
        obj.selection.range[3] = obj.selection.range[1];
    }

    obj.updateSelectionFromCoords(obj.selection.range[0], obj.selection.range[1], obj.selection.range[2], obj.selection.range[3]);
    updateScroll.call(obj, 0);
};

export const first: WorksheetMethod<'first'> = function (shiftKey, ctrlKey) {
    const obj = this;

    // Without a selection, there is nothing to move
    if (!hasSelection(obj)) {
        return;
    }

    if (shiftKey) {
        if (ctrlKey) {
            obj.selection.range[3] = 0;
        } else {
            leftVisible.call(obj, 1, 0);
        }
    } else {
        if (ctrlKey) {
            obj.selection.range[1] = 0;
        } else {
            leftVisible.call(obj, 0, 0);
        }
        toFirstCell.call(obj);
        obj.selection.range[2] = obj.selection.range[0];
        obj.selection.range[3] = obj.selection.range[1];
    }

    // Change page
    if (obj.options.lazyLoading == true && (obj.selection.range[1] == 0 || obj.selection.range[3] == 0)) {
        obj.view.loadPage(0);
    } else if (obj.options.pagination! > 0) {
        const pageNumber = obj.whichPage(obj.selection.range[3]);
        if (pageNumber != obj.view.pageNumber) {
            obj.page(pageNumber);
        }
    }

    obj.updateSelectionFromCoords(obj.selection.range[0], obj.selection.range[1], obj.selection.range[2], obj.selection.range[3]);
    updateScroll.call(obj, 1);
};

export const last: WorksheetMethod<'last'> = function (shiftKey, ctrlKey) {
    const obj = this;

    // Without a selection, there is nothing to move
    if (!hasSelection(obj)) {
        return;
    }

    if (shiftKey) {
        if (ctrlKey) {
            obj.selection.range[3] = obj.records.length - 1;
        } else {
            rightVisible.call(obj, 1, 0);
        }
    } else {
        if (ctrlKey) {
            obj.selection.range[1] = obj.records.length - 1;
        } else {
            rightVisible.call(obj, 0, 0);
        }
        toFirstCell.call(obj);
        obj.selection.range[2] = obj.selection.range[0];
        obj.selection.range[3] = obj.selection.range[1];
    }

    // Change page
    if (obj.options.lazyLoading == true && (obj.selection.range[1] == obj.records.length - 1 || obj.selection.range[3] == obj.records.length - 1)) {
        obj.view.loadPage(-1);
    } else if (obj.options.pagination! > 0) {
        const pageNumber = obj.whichPage(obj.selection.range[3]);
        if (pageNumber != obj.view.pageNumber) {
            obj.page(pageNumber);
        }
    }

    obj.updateSelectionFromCoords(obj.selection.range[0], obj.selection.range[1], obj.selection.range[2], obj.selection.range[3]);
    updateScroll.call(obj, 3);
};

/**
 * Update scroll position based on the selection
 */
export const updateScroll = function (this: Worksheet, direction: 0 | 1 | 2 | 3) {
    const obj = this;

    // Jspreadsheet Container information
    const contentRect = obj.content.getBoundingClientRect();
    const x1 = contentRect.left;
    const y1 = contentRect.top;
    const w1 = contentRect.width;
    const h1 = contentRect.height;

    // Direction Left or Up
    const reference = obj.records[obj.selection.range![3]][obj.selection.range![2]].element;

    // Reference
    const referenceRect = reference.getBoundingClientRect();
    const x2 = referenceRect.left;
    const y2 = referenceRect.top;
    const w2 = referenceRect.width;
    const h2 = referenceRect.height;

    let x, y;

    // Direction
    if (direction == 0 || direction == 1) {
        x = x2 - x1 + obj.content.scrollLeft;
        y = y2 - y1 + obj.content.scrollTop - 2;
    } else {
        x = x2 - x1 + obj.content.scrollLeft + w2;
        y = y2 - y1 + obj.content.scrollTop + h2;
    }

    // Top position check
    if (y > obj.content.scrollTop + 30 && y < obj.content.scrollTop + h1) {
        // In the viewport
    } else {
        // Out of viewport
        if (y < obj.content.scrollTop + 30) {
            obj.content.scrollTop = y - h2;
        } else {
            obj.content.scrollTop = y - (h1 - 2);
        }
    }

    // Freeze columns?
    const freezed = getFreezeWidth.call(obj);

    // The frozen columns stay at the left of the content only when scrollLeft is more than the width of the index column.
    // Then a cell whose left border is under the frozen columns goes to the right of the frozen columns.
    const indexWidth = getIndexWidth.call(obj);
    const frozen = freezed > 0 && obj.content.scrollLeft > indexWidth;
    const left = x2 - x1 + obj.content.scrollLeft;
    const underFreezed = frozen && left < obj.content.scrollLeft + freezed;

    // Left position check - TODO: change that to the bottom border of the element
    if (
        (frozen && obj.selection.range![2] < obj.options.freezeColumns!) ||
        (!underFreezed && x > obj.content.scrollLeft + freezed && x < obj.content.scrollLeft + w1)
    ) {
        // In the viewport
    } else if (underFreezed) {
        obj.content.scrollLeft = left - freezed - 1 < indexWidth ? 0 : left - freezed - 1;
    } else {
        // Out of viewport
        if (x < obj.content.scrollLeft + 30) {
            obj.content.scrollLeft = x;
            if (obj.content.scrollLeft < 50) {
                obj.content.scrollLeft = 0;
            }
        } else if (x < obj.content.scrollLeft + freezed) {
            obj.content.scrollLeft = x - freezed - 1;
        } else {
            obj.content.scrollLeft = x - (w1 - 20);
        }
    }
};
