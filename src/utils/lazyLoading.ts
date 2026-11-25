import { closeEditor } from './editor.js';
import { updateTabStop } from './accessibility.js';
import type { Worksheet } from '../types/internal';

/**
 * Remove a row from the table. If the row holds the open editor, save and close the editor first.
 */
// The callers remove a row when the body has more than a page of rows, so it has a first and a last row
const removeRow = function (this: Worksheet, row: ChildNode) {
    const obj = this;

    if (obj.editor.isOpen() && row.contains(obj.editor.cell)) {
        closeEditor.call(obj, obj.editor.cell, true);
    }

    obj.tbody.removeChild(row);
};

/**
 * Go to a page in a lazyLoading
 */
export const loadPage = function (this: Worksheet, pageNumber: number) {
    const obj = this;

    // Search: the indexes of the rows found, or the rows
    let results: number[] | Worksheet['rows'];

    if ((obj.options.search == true || obj.options.filters == true) && obj.results) {
        results = obj.results;
    } else {
        results = obj.rows;
    }

    // Per page
    const quantityPerPage = 100;

    // pageNumber
    if (pageNumber == null || pageNumber == -1) {
        // Last page
        pageNumber = Math.ceil(results.length / quantityPerPage) - 1;
    }

    let startRow = pageNumber * quantityPerPage;
    let finalRow = pageNumber * quantityPerPage + quantityPerPage;
    if (finalRow > results.length) {
        finalRow = results.length;
    }
    startRow = finalRow - 100;
    if (startRow < 0) {
        startRow = 0;
    }

    // Appeding items
    for (let j = startRow; j < finalRow; j++) {
        if ((obj.options.search == true || obj.options.filters == true) && obj.results) {
            obj.tbody.appendChild(obj.rows[results[j] as number].element);
        } else {
            obj.tbody.appendChild(obj.rows[j].element);
        }

        if (obj.tbody.children.length > quantityPerPage) {
            removeRow.call(obj, obj.tbody.firstChild!);
        }
    }

    // The tab stop of the grid can be in a removed row
    updateTabStop.call(obj);
};

export const loadValidation = function (this: Worksheet) {
    const obj = this;

    if (obj.selection.range) {
        let currentRow = parseInt((obj.tbody.firstChild as HTMLElement).getAttribute('data-y'));
        let selectedRow = parseInt(obj.selection.range[3]);
        let total = obj.rows.length;

        // With search or filters, the pages are pages of the results
        if ((obj.options.search == true || obj.options.filters == true) && obj.results) {
            currentRow = obj.results.indexOf(currentRow);
            selectedRow = obj.results.indexOf(selectedRow);
            total = obj.results.length;

            if (selectedRow < 0) {
                return false;
            }
        }

        const currentPage = currentRow / 100;
        const selectedPage = parseInt(selectedRow / 100);
        const totalPages = parseInt(total / 100);

        if (currentPage != selectedPage && selectedPage <= totalPages) {
            if (Array.prototype.indexOf.call(obj.tbody.children, obj.rows[obj.selection.range[3] as number].element) < 0) {
                loadPage.call(obj, selectedPage);
                return true;
            }
        }
    }

    return false;
};

export const loadUp = function (this: Worksheet) {
    const obj = this;

    // Search: the indexes of the rows found, or the rows
    let results: number[] | Worksheet['rows'];

    if ((obj.options.search == true || obj.options.filters == true) && obj.results) {
        results = obj.results;
    } else {
        results = obj.rows;
    }
    let test = 0;
    if (results.length > 100) {
        // Get the first element in the page
        let item = parseInt((obj.tbody.firstChild as HTMLElement).getAttribute('data-y'));
        if ((obj.options.search == true || obj.options.filters == true) && obj.results) {
            item = (results as number[]).indexOf(item);
        }
        if (item > 0) {
            for (let j = 0; j < 30; j++) {
                item = item - 1;
                if (item > -1) {
                    if ((obj.options.search == true || obj.options.filters == true) && obj.results) {
                        obj.tbody.insertBefore(obj.rows[results[item] as number].element, obj.tbody.firstChild);
                    } else {
                        obj.tbody.insertBefore(obj.rows[item].element, obj.tbody.firstChild);
                    }
                    if (obj.tbody.children.length > 100) {
                        removeRow.call(obj, obj.tbody.lastChild!);
                        test = 1;
                    }
                }
            }
        }
    }

    // The tab stop of the grid can be in a removed row
    updateTabStop.call(obj);

    return test;
};

export const loadDown = function (this: Worksheet) {
    const obj = this;

    // Search: the indexes of the rows found, or the rows
    let results: number[] | Worksheet['rows'];

    if ((obj.options.search == true || obj.options.filters == true) && obj.results) {
        results = obj.results;
    } else {
        results = obj.rows;
    }
    let test = 0;
    if (results.length > 100) {
        // Get the last element in the page
        let item = parseInt((obj.tbody.lastChild as HTMLElement).getAttribute('data-y'));
        if ((obj.options.search == true || obj.options.filters == true) && obj.results) {
            item = (results as number[]).indexOf(item);
        }
        if (item < obj.rows.length - 1) {
            for (let j = 0; j <= 30; j++) {
                if (item < results.length) {
                    if ((obj.options.search == true || obj.options.filters == true) && obj.results) {
                        obj.tbody.appendChild(obj.rows[results[item] as number].element);
                    } else {
                        obj.tbody.appendChild(obj.rows[item].element);
                    }
                    if (obj.tbody.children.length > 100) {
                        removeRow.call(obj, obj.tbody.firstChild!);
                        test = 1;
                    }
                }
                item = item + 1;
            }
        }
    }

    // The tab stop of the grid can be in a removed row
    updateTabStop.call(obj);

    return test;
};
