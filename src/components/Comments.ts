import { getCoordsFromCellName } from '../utils/helpers.js';
import { getColumnNameFromId, getIdFromColumnName } from '../utils/internalHelpers.js';
import { recordRows } from './RowSource.js';
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
        const worksheet = this.worksheet;

        if (cell) {
            if (typeof cell !== 'string') {
                return this.get();
            }

            cell = getIdFromColumnName(cell);

            return worksheet.records[cell[1]][cell[0]].element.getAttribute('title') || '';
        } else {
            const data: Record<string, string> = {};
            for (const j of recordRows(worksheet)) {
                for (let i = 0; i < worksheet.options.columns.length; i++) {
                    // A column outside the column window has no cell
                    const comments = worksheet.records[j][i]?.element.getAttribute('title');
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
        const worksheet = this.worksheet;

        let commentsObj;

        if (typeof cellId === 'string') {
            commentsObj = { [cellId]: comments };
        } else {
            commentsObj = cellId;
        }

        const oldValue = this.apply(commentsObj);

        // Save history
        worksheet.history.add({
            action: 'setComments',
            newValue: commentsObj,
            oldValue,
        });

        // Set comments
        worksheet.dispatch('oncomments', worksheet, commentsObj, oldValue);
    }

    /**
     * Load the comments of the comments option, without history or events.
     * @param comments - The comments by cell name.
     */
    load(comments: Record<string, string | null | undefined>) {
        this.apply(comments);
    }

    /**
     * Set the comments of the cells, and return the comments they had.
     * @param commentsObj - The comments by cell name. A falsy value removes the comment.
     */
    private apply(commentsObj: Record<string, string | null | undefined>) {
        const worksheet = this.worksheet;

        const oldValue: Record<string, string | null> = {};

        for (const [cellName, comment] of Object.entries(commentsObj)) {
            const cellCoords = getCoordsFromCellName(cellName);

            // In the row source mode, a cell outside the window has no record
            if (!worksheet.records[cellCoords![1]!]?.[cellCoords![0]]) {
                continue;
            }

            // Keep old value
            oldValue[cellName] = worksheet.records[cellCoords![1]!][cellCoords![0]].element.getAttribute('title');

            // Set new values
            worksheet.records[cellCoords![1]!][cellCoords![0]].element.setAttribute('title', comment ? comment : '');

            // Remove class if there is no comment
            if (comment) {
                worksheet.records[cellCoords![1]!][cellCoords![0]].element.classList.add('jss_comments');

                if (!worksheet.options.comments) {
                    worksheet.options.comments = {};
                }

                worksheet.options.comments[cellName] = comment;
            } else {
                worksheet.records[cellCoords![1]!][cellCoords![0]].element.classList.remove('jss_comments');

                if (worksheet.options.comments?.[cellName]) {
                    delete worksheet.options.comments[cellName];
                }
            }
        }

        return oldValue;
    }
}
