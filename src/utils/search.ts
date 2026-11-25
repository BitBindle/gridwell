import { resetFilters } from './filter.js';
import { getIdFromColumnName } from './internalHelpers.js';
import { updateResult } from './internal.js';
import { isRowMerged } from './merges.js';
import type { Worksheet } from '../types/internal';

/**
 * Search
 * @param query - The text to find. An empty text shows all rows again.
 */
export const search = function (this: Worksheet, query: string) {
    const obj = this;

    // Reset any filter
    if (obj.options.filters) {
        resetFilters.call(obj);
    }

    // Reset selection
    obj.resetSelection();

    // Total of results
    obj.pageNumber = 0;
    obj.results = [];

    if (query) {
        if (obj.searchInput.value !== query) {
            obj.searchInput.value = query;
        }

        /**
         * Search filter
         * @param item - The data of the row.
         * @param index - The row.
         */
        const search = function (item: any, query: RegExp, index: number) {
            for (let i = 0; i < item.length; i++) {
                if (('' + item[i]).toLowerCase().search(query) >= 0 || ('' + obj.records[index][i].element.innerHTML).toLowerCase().search(query) >= 0) {
                    return true;
                }
            }
            return false;
        };

        /**
         * Result
         * @param k - The row.
         */
        const addToResult = function (k: number) {
            if (obj.results.indexOf(k) == -1) {
                obj.results.push(k);
            }
        };

        let parsedQuery: string | RegExp = query.replace(/[-[\]{}()*+?.,\\^$|#\s]/g, '\\$&');
        parsedQuery = new RegExp(parsedQuery, 'i');

        // Filter
        obj.options.data.forEach(function (v, k) {
            if (search(v, parsedQuery, k)) {
                // Merged rows found
                const rows = isRowMerged.call(obj, k);
                if (rows.length) {
                    for (let i = 0; i < rows.length; i++) {
                        const row = getIdFromColumnName(rows[i], true);
                        for (let j = 0; j < obj.options.mergeCells[rows[i]][1]; j++) {
                            addToResult(row[1] + j);
                        }
                    }
                } else {
                    // Normal row found
                    addToResult(k);
                }
            }
        });
    } else {
        obj.results = null;
    }

    updateResult.call(obj);
};

/**
 * Reset search
 */
export const resetSearch = function (this: Worksheet) {
    const obj = this;

    obj.searchInput.value = '';
    obj.search('');
    obj.results = null;
};
