import { getIndexWidth } from './freeze.js';
import { updateCornerPosition } from './selection.js';
import type { Worksheet } from '../types/internal';

/**
 * Place the frozen columns at the left of the content after a scroll or a change of the columns.
 */
export const updateFreezePosition = function (this: Worksheet) {
    const obj = this;

    const scrollLeft = obj.content.scrollLeft;
    const indexWidth = getIndexWidth.call(obj);
    // The worksheet can have fewer columns than freezeColumns, for example after deleteColumn
    const freezeColumns = Math.min(obj.options.freezeColumns!, obj.headers.length);
    let width = 0;
    if (scrollLeft > indexWidth) {
        // The last frozen column that is not hidden shows the line of the frozen pane
        let last = freezeColumns - 1;
        while (last > 0 && obj.headers[last].style.display === 'none') {
            last--;
        }

        for (let i = 0; i < freezeColumns; i++) {
            if (i > 0) {
                // Must check if the previous column is hidden or not to determin whether the width shoule be added or not!
                if (!obj.options.columns || !obj.options.columns[i - 1] || obj.options.columns[i - 1].type !== 'hidden') {
                    let columnWidth;
                    if (obj.options.columns && obj.options.columns[i - 1] && obj.options.columns[i - 1].width !== undefined) {
                        columnWidth = parseInt(obj.options.columns[i - 1].width);
                    } else {
                        columnWidth = obj.options.defaultColWidth !== undefined ? parseInt(obj.options.defaultColWidth) : 100;
                    }

                    width += parseInt(columnWidth);
                }
            }
            // Only the cells of the last column, or of a column that was the last before, get a change of the line
            const isLast = i === last;
            const wasLast = obj.headers[i].classList.contains('jss_freezed_last');
            obj.headers[i].classList.add('jss_freezed');
            obj.headers[i].classList.toggle('jss_freezed_last', isLast);
            obj.headers[i].style.left = width + 'px';
            for (let j = 0; j < obj.rows.length; j++) {
                if (obj.rows[j] && obj.records[j][i]) {
                    // The cells are relative, so all the frozen cells move by the same distance as the scroll
                    const shifted = scrollLeft - indexWidth - 1 + 'px';
                    obj.records[j][i].element.classList.add('jss_freezed');
                    if (isLast || wasLast) {
                        obj.records[j][i].element.classList.toggle('jss_freezed_last', isLast);
                    }
                    obj.records[j][i].element.style.left = shifted;
                }
            }
        }
    } else {
        for (let i = 0; i < freezeColumns; i++) {
            obj.headers[i].classList.remove('jss_freezed', 'jss_freezed_last');
            obj.headers[i].style.left = '';
            for (let j = 0; j < obj.rows.length; j++) {
                if (obj.records[j][i]) {
                    obj.records[j][i].element.classList.remove('jss_freezed', 'jss_freezed_last');
                    obj.records[j][i].element.style.left = '';
                }
            }
        }
    }

    // Place the corner in the correct place
    updateCornerPosition.call(obj);
};
