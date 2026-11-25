import { updateCornerPosition } from './selection.js';
import type { Worksheet } from '../types/internal';

/**
 * Run after the cells change: add the spare rows and columns that minSpareRows and minSpareCols ask for, update the
 * footers and place the selection corner again.
 */
export const updateTable = function (this: Worksheet) {
    const obj = this;

    // Check for spare
    if (obj.options.minSpareRows! > 0) {
        let numBlankRows = 0;
        for (let j = obj.rows.length - 1; j >= 0; j--) {
            let test = false;
            for (let i = 0; i < obj.headers.length; i++) {
                if (obj.options.data[j][i]) {
                    test = true;
                }
            }
            if (test) {
                break;
            } else {
                numBlankRows++;
            }
        }

        if (obj.options.minSpareRows! - numBlankRows > 0) {
            obj.insertRow(obj.options.minSpareRows! - numBlankRows);
        }
    }

    if (obj.options.minSpareCols! > 0) {
        let numBlankCols = 0;
        for (let i = obj.headers.length - 1; i >= 0; i--) {
            let test = false;
            for (let j = 0; j < obj.rows.length; j++) {
                if (obj.options.data[j][i]) {
                    test = true;
                }
            }
            if (test) {
                break;
            } else {
                numBlankCols++;
            }
        }

        if (obj.options.minSpareCols! - numBlankCols > 0) {
            obj.insertColumn(obj.options.minSpareCols! - numBlankCols);
        }
    }

    // Update footers
    if (obj.options.footers) {
        obj.formulas.setFooter();
    }

    if (obj.options.columns.length < obj.options.minDimensions[0]) {
        obj.options.minDimensions[0] = obj.options.columns.length;
    }

    // Update corner position
    setTimeout(function () {
        updateCornerPosition.call(obj);
    }, 0);
};
