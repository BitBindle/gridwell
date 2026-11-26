import { getCoordsFromCellName } from '../utils/helpers.js';
import { getColumnNameFromId, getIdFromColumnName } from '../utils/internalHelpers.js';
import type { Worksheet } from '../types/internal';

/**
 * The comments of the cells of a worksheet, kept in the title of each cell and in options.comments.
 *
 * Each worksheet has one, as worksheet.comments. getComments and setComments call it.
 */
export class Comments {
    readonly worksheet: Worksheet;

    constructor(worksheet: Worksheet) {
        this.worksheet = worksheet;
    }

    /**
     * Get cell comments, null cell for all
     * @param cell - The cell name, for example "B3".
     */
    get(cell?: string | number[]): Record<string, string> | string {
        const obj = this.worksheet;

        if (cell) {
            if (typeof cell !== 'string') {
                return this.get();
            }

            cell = getIdFromColumnName(cell, true);

            return obj.records[cell[1]][cell[0]].element.getAttribute('title') || '';
        } else {
            const data: Record<string, string> = {};
            for (let j = 0; j < obj.options.data.length; j++) {
                for (let i = 0; i < obj.options.columns.length; i++) {
                    const comments = obj.records[j][i].element.getAttribute('title');
                    if (comments) {
                        const cell = getColumnNameFromId([i, j]);
                        data[cell] = comments;
                    }
                }
            }
            return data;
        }
    }

    /**
     * Set cell comments
     * @param cellId - The cell name, for example "B3", or the comments by cell name.
     * @param comments - The comment, with a cell name.
     */
    set(cellId: string | Record<string, string | null | undefined>, comments?: string | null) {
        const obj = this.worksheet;

        let commentsObj;

        if (typeof cellId == 'string') {
            commentsObj = { [cellId]: comments };
        } else {
            commentsObj = cellId;
        }

        const oldValue: Record<string, string | null> = {};

        Object.entries(commentsObj).forEach(function ([cellName, comment]) {
            const cellCoords = getCoordsFromCellName(cellName);

            // Keep old value
            oldValue[cellName] = obj.records[cellCoords![1]!][cellCoords![0]].element.getAttribute('title');

            // Set new values
            obj.records[cellCoords![1]!][cellCoords![0]].element.setAttribute('title', comment ? comment : '');

            // Remove class if there is no comment
            if (comment) {
                obj.records[cellCoords![1]!][cellCoords![0]].element.classList.add('jss_comments');

                if (!obj.options.comments) {
                    obj.options.comments = {};
                }

                obj.options.comments[cellName] = comment;
            } else {
                obj.records[cellCoords![1]!][cellCoords![0]].element.classList.remove('jss_comments');

                if (obj.options.comments && obj.options.comments[cellName]) {
                    delete obj.options.comments[cellName];
                }
            }
        });

        // Save history
        obj.history.add({
            action: 'setComments',
            newValue: commentsObj,
            oldValue: oldValue,
        });

        // Set comments
        obj.dispatch('oncomments', obj, commentsObj, oldValue);
    }
}
