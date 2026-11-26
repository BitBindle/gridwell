import jSuites from 'jsuites';
import dispatch from '../utils/dispatch.js';
import { getIdFromColumnName } from '../utils/internalHelpers.js';
import { refreshSelection, updateCornerPosition } from '../utils/selection.js';
import { updateTabStop } from '../utils/accessibility.js';
import type jspreadsheet from '../jspreadsheet';
import type { CellValue } from '../jspreadsheet';
import type { Worksheet } from '../types/internal';

/**
 * Which rows a worksheet shows: the search, the filters, the pagination, the order of the rows and the lazy loading. It
 * keeps the rows that a search or the filters found, the values of the filters and the page.
 *
 * Each worksheet has one, as worksheet.view. The public methods search, resetSearch, openFilter, resetFilters, page,
 * whichPage, quantiyOfPages and orderBy call it.
 */
export class View {
    readonly worksheet: Worksheet;

    /** The indexes of the rows that a search or the filters found, or null when the worksheet shows every row. */
    results: number[] | null = null;

    /** The values that each column is filtered by, by column index. null or a missing entry is a column without filter. */
    filters: (string[] | null)[] = [];

    /** The page that the worksheet shows, from 0, when it has pagination. */
    pageNumber = 0;

    constructor(worksheet: Worksheet) {
        this.worksheet = worksheet;
    }

    /**
     * Search
     * @param query - The text to find. An empty text shows all rows again.
     */
    search(query: string) {
        const obj = this.worksheet;

        // Reset any filter
        if (obj.options.filters) {
            this.resetFilters();
        }

        // Reset selection
        obj.resetSelection();

        // Total of results
        obj.view.pageNumber = 0;
        obj.view.results = [];

        if (query) {
            if (obj.searchInput.value !== query) {
                obj.searchInput.value = query;
            }

            /**
             * Search filter
             * @param item - The data of the row.
             * @param index - The row.
             */
            const search = function (item: CellValue[], query: RegExp, index: number) {
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
                if (obj.view.results!.indexOf(k) == -1) {
                    obj.view.results!.push(k);
                }
            };

            let parsedQuery: string | RegExp = query.replace(/[-[\]{}()*+?.,\\^$|#\s]/g, '\\$&');
            parsedQuery = new RegExp(parsedQuery, 'i');

            // Filter
            obj.options.data.forEach(function (v, k) {
                if (search(v, parsedQuery, k)) {
                    // Merged rows found
                    const rows = obj.merges.isRowMerged(k);
                    if (rows.length) {
                        for (let i = 0; i < rows.length; i++) {
                            const row = getIdFromColumnName(rows[i], true);
                            for (let j = 0; j < obj.options.mergeCells![rows[i]][1]; j++) {
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
            obj.view.results = null;
        }

        this.updateResult();
    }

    /**
     * Reset search
     */
    resetSearch() {
        const obj = this.worksheet;

        obj.searchInput.value = '';
        obj.search('');
        obj.view.results = null;
    }

    /**
     * Open the column filter
     */
    openFilter(columnId: number) {
        const obj = this.worksheet;

        if (!obj.options.filters) {
            console.log('Jspreadsheet: filters not enabled.');
        } else {
            // Make sure is integer
            columnId = parseInt(columnId);
            // Reset selection
            obj.resetSelection();
            // Load options
            let optionsFiltered = [];
            if (obj.options.columns[columnId].type == 'checkbox') {
                optionsFiltered.push({ id: 'true', name: 'True' });
                optionsFiltered.push({ id: 'false', name: 'False' });
            } else {
                // Labels by value. An array used as a map, as before.
                const options = [] as unknown as Record<string, string>;
                let hasBlanks = false;
                for (let j = 0; j < obj.options.data.length; j++) {
                    const k = obj.options.data[j][columnId];
                    const v = obj.records[j][columnId].element.innerHTML;
                    if (k && v) {
                        options[k as string] = v;
                    } else {
                        hasBlanks = true;
                    }
                }
                const keys = Object.keys(options);
                optionsFiltered = [];
                for (let j = 0; j < keys.length; j++) {
                    optionsFiltered.push({ id: keys[j], name: options[keys[j]] });
                }
                // Has blank options
                if (hasBlanks) {
                    optionsFiltered.push({ value: '', id: '', name: '(Blanks)' });
                }
            }

            // Create dropdown
            const div = document.createElement('div');
            obj.filter!.children[columnId + 1].innerHTML = '';
            obj.filter!.children[columnId + 1].appendChild(div);
            (obj.filter!.children[columnId + 1] as HTMLElement).style.paddingLeft = '0px';
            (obj.filter!.children[columnId + 1] as HTMLElement).style.paddingRight = '0px';
            (obj.filter!.children[columnId + 1] as HTMLElement).style.overflow = 'initial';

            const opt = {
                data: optionsFiltered,
                multiple: true,
                autocomplete: true,
                opened: true,
                value: obj.view.filters[columnId] !== undefined ? obj.view.filters[columnId] : null,
                width: '100%',
                position: obj.options.tableOverflow == true || obj.parent.config.fullscreen == true ? true : false,
                // The element of the dropdown, which holds the dropdown
                onclose: function (o: HTMLElement & { dropdown: ReturnType<typeof jSuites.dropdown> }) {
                    obj.view.resetFilters();
                    // A list, because the dropdown has multiple
                    obj.view.filters[columnId] = o.dropdown.getValue(true) as string[];
                    obj.filter!.children[columnId + 1].innerHTML = o.dropdown.getText() as string;
                    (obj.filter!.children[columnId + 1] as HTMLElement).style.paddingLeft = '';
                    (obj.filter!.children[columnId + 1] as HTMLElement).style.paddingRight = '';
                    (obj.filter!.children[columnId + 1] as HTMLElement).style.overflow = '';
                    obj.view.closeFilter(columnId);
                    refreshSelection.call(obj);
                },
            };

            // Dynamic dropdown
            jSuites.dropdown(div, opt);
        }
    }

    closeFilter(columnId?: number) {
        const obj = this.worksheet;

        if (!columnId) {
            for (let i = 0; i < obj.filter!.children.length; i++) {
                if (obj.view.filters[i]) {
                    columnId = i;
                }
            }
        }

        // Search filter
        const search = function (query: string[], x: number, y: number) {
            for (let i = 0; i < query.length; i++) {
                const value = '' + obj.options.data[y][x];
                const label = '' + obj.records[y][x].element.innerHTML;
                if (query[i] == value || query[i] == label) {
                    return true;
                }
            }
            return false;
        };

        // The column has a filter: openFilter passes its column, and orderBy calls this only while a filter gives results
        const query = obj.view.filters[columnId!]!;
        obj.view.results = [];
        for (let j = 0; j < obj.options.data.length; j++) {
            if (search(query, columnId!, j)) {
                obj.view.results.push(j);
            }
        }
        if (!obj.view.results.length) {
            obj.view.results = null;
        }

        this.updateResult();
    }

    resetFilters() {
        const obj = this.worksheet;

        if (obj.options.filters) {
            for (let i = 0; i < obj.filter!.children.length; i++) {
                obj.filter!.children[i].innerHTML = '&nbsp;';
                obj.view.filters[i] = null;
            }
        }

        obj.view.results = null;
        this.updateResult();
    }

    /**
     * Which page the row is
     */
    whichPage(row: number) {
        const obj = this.worksheet;

        // Search
        if ((obj.options.search == true || obj.options.filters == true) && obj.view.results) {
            row = obj.view.results.indexOf(row);
        }

        return Math.ceil((parseInt(row) + 1) / parseInt(obj.options.pagination)) - 1;
    }

    /**
     * Update the pagination
     */
    updatePagination() {
        const obj = this.worksheet;

        // Reset container
        obj.pagination.children[0].innerHTML = '';
        obj.pagination.children[1].innerHTML = '';

        // Start pagination
        if (obj.options.pagination) {
            // Searchable
            let results: number;

            if ((obj.options.search == true || obj.options.filters == true) && obj.view.results) {
                results = obj.view.results.length;
            } else {
                results = obj.rows.length;
            }

            if (!results) {
                // No records found
                obj.pagination.children[0].innerHTML = jSuites.translate('No records found');
            } else {
                // Pagination container
                const quantyOfPages = Math.ceil(results / obj.options.pagination);

                let startNumber, finalNumber;

                if (obj.view.pageNumber < 6) {
                    startNumber = 1;
                    finalNumber = quantyOfPages < 10 ? quantyOfPages : 10;
                } else if (quantyOfPages - obj.view.pageNumber < 5) {
                    startNumber = quantyOfPages - 9;
                    finalNumber = quantyOfPages;
                    if (startNumber < 1) {
                        startNumber = 1;
                    }
                } else {
                    startNumber = obj.view.pageNumber - 4;
                    finalNumber = obj.view.pageNumber + 5;
                }

                // First
                if (startNumber > 1) {
                    const paginationItem = document.createElement('div');
                    paginationItem.className = 'jss_page';
                    paginationItem.innerHTML = '<';
                    paginationItem.title = '1';
                    obj.pagination.children[1].appendChild(paginationItem);
                }

                // Get page links
                for (let i = startNumber; i <= finalNumber; i++) {
                    const paginationItem = document.createElement('div');
                    paginationItem.className = 'jss_page';
                    paginationItem.innerHTML = String(i);
                    obj.pagination.children[1].appendChild(paginationItem);

                    if (obj.view.pageNumber == i - 1) {
                        paginationItem.classList.add('jss_page_selected');
                    }
                }

                // Last
                if (finalNumber < quantyOfPages) {
                    const paginationItem = document.createElement('div');
                    paginationItem.className = 'jss_page';
                    paginationItem.innerHTML = '>';
                    paginationItem.title = String(quantyOfPages);
                    obj.pagination.children[1].appendChild(paginationItem);
                }

                // Text
                const format = function (format: string, ...args: number[]): string {
                    return format.replace(/{(\d+)}/g, function (match: string, number: number) {
                        return typeof args[number] != 'undefined' ? String(args[number]) : match;
                    });
                };

                obj.pagination.children[0].innerHTML = format(jSuites.translate('Showing page {0} of {1} entries'), obj.view.pageNumber + 1, quantyOfPages);
            }
        }
    }

    /**
     * Go to page
     */
    page(pageNumber: number) {
        const obj = this.worksheet;

        const oldPage = obj.view.pageNumber;

        // Search: the indexes of the rows found, or the rows
        let results: number[] | Worksheet['rows'];

        if ((obj.options.search == true || obj.options.filters == true) && obj.view.results) {
            results = obj.view.results;
        } else {
            results = obj.rows;
        }

        // Per page
        const quantityPerPage = parseInt(obj.options.pagination);

        // pageNumber
        if (pageNumber == null || pageNumber == -1) {
            // Last page
            pageNumber = Math.ceil(results.length / quantityPerPage) - 1;
        }

        // Page number
        obj.view.pageNumber = pageNumber;

        let startRow = pageNumber * quantityPerPage;
        let finalRow = pageNumber * quantityPerPage + quantityPerPage;
        if (finalRow > results.length) {
            finalRow = results.length;
        }
        if (startRow < 0) {
            startRow = 0;
        }

        // Reset container
        while (obj.tbody.firstChild) {
            obj.tbody.removeChild(obj.tbody.firstChild);
        }

        // Appeding items
        for (let j = startRow; j < finalRow; j++) {
            if ((obj.options.search == true || obj.options.filters == true) && obj.view.results) {
                obj.tbody.appendChild(obj.rows[results[j] as number].element);
            } else {
                obj.tbody.appendChild(obj.rows[j].element);
            }
        }

        // The tab stop of the grid can be in a row of the old page
        updateTabStop.call(obj);

        if (obj.options.pagination! > 0) {
            this.updatePagination();
        }

        // Update corner position
        updateCornerPosition.call(obj);

        // Events. setData sets the page number of a worksheet with pagination.
        dispatch.call(obj, 'onchangepage', obj, pageNumber, oldPage, obj.options.pagination!);
    }

    countPages() {
        const obj = this.worksheet;

        let results: number;
        if ((obj.options.search == true || obj.options.filters == true) && obj.view.results) {
            results = obj.view.results.length;
        } else {
            results = obj.rows.length;
        }

        return Math.ceil(results / obj.options.pagination!);
    }

    updateResult() {
        const obj = this.worksheet;

        let total = 0;
        let index = 0;

        // Page 1
        if (obj.options.lazyLoading == true) {
            total = 100;
        } else if (obj.options.pagination! > 0) {
            total = obj.options.pagination!;
        } else {
            if (obj.view.results) {
                total = obj.view.results.length;
            } else {
                total = obj.rows.length;
            }
        }

        // Reset current nodes
        while (obj.tbody.firstChild) {
            obj.tbody.removeChild(obj.tbody.firstChild);
        }

        // Hide all records from the table
        for (let j = 0; j < obj.rows.length; j++) {
            if (!obj.view.results || obj.view.results.indexOf(j) > -1) {
                if (index < total) {
                    obj.tbody.appendChild(obj.rows[j].element);
                    index++;
                }
                obj.rows[j].element.style.display = '';
            } else {
                obj.rows[j].element.style.display = 'none';
            }
        }

        // The tab stop of the grid can be in a row that is not in the results
        updateTabStop.call(obj);

        // Update pagination
        if (obj.options.pagination! > 0) {
            this.updatePagination();
        }

        updateCornerPosition.call(obj);

        dispatch.call(obj, 'onupdateresult', obj, obj.view.results);

        return total;
    }

    /**
     * Update order arrow
     */
    updateOrderArrow(column: number, order: boolean | number) {
        const obj = this.worksheet;

        // Remove order
        for (let i = 0; i < obj.headers.length; i++) {
            obj.headers[i].classList.remove('arrow-up');
            obj.headers[i].classList.remove('arrow-down');
        }

        // No order specified then toggle order
        if (order) {
            obj.headers[column].classList.add('arrow-up');
        } else {
            obj.headers[column].classList.add('arrow-down');
        }
    }

    /**
     * Update rows position
     */
    updateOrder(rows: number[]) {
        const obj = this.worksheet;

        // History
        const data: Worksheet['options']['data'] = [];
        for (let j = 0; j < rows.length; j++) {
            data[j] = obj.options.data[rows[j]];
        }
        obj.options.data = data;

        const records: Worksheet['records'] = [];
        for (let j = 0; j < rows.length; j++) {
            records[j] = obj.records[rows[j]];

            for (let i = 0; i < records[j].length; i++) {
                records[j][i].y = j;
            }
        }
        obj.records = records;

        const rowObjects: Worksheet['rows'] = [];
        for (let j = 0; j < rows.length; j++) {
            rowObjects[j] = obj.rows[rows[j]];
            rowObjects[j].y = j;
        }
        obj.rows = rowObjects;

        // Update references
        obj.formulas.updateTableReferences();

        // Redo search
        if (obj.view.results && obj.view.results.length) {
            if (obj.searchInput.value) {
                obj.search(obj.searchInput.value);
            } else {
                this.closeFilter();
            }
        } else {
            // Create page
            obj.view.results = null;
            obj.view.pageNumber = 0;

            if (obj.options.pagination! > 0) {
                obj.page(0);
            } else if (obj.options.lazyLoading == true) {
                this.loadPage(0);
            } else {
                for (let j = 0; j < obj.rows.length; j++) {
                    obj.tbody.appendChild(obj.rows[j].element);
                }
                // A moved row loses the focus
                updateTabStop.call(obj);
            }
        }
    }

    /**
     * Sort data and reload table
     */
    orderBy(column: number, order?: 0 | 1): boolean | undefined {
        const obj = this.worksheet;

        if (column >= 0) {
            // Merged cells
            if (obj.options.mergeCells && Object.keys(obj.options.mergeCells).length > 0) {
                if (!confirm(jSuites.translate('This action will destroy any existing merged cells. Are you sure?'))) {
                    return false;
                } else {
                    // Remove merged cells
                    obj.destroyMerge();
                }
            }

            // Direction
            if (order == null) {
                order = obj.headers[column].classList.contains('arrow-down') ? 1 : 0;
            } else {
                order = order ? 1 : 0;
            }

            // Test order
            // The row and the value to sort by
            let temp: jspreadsheet.SortingItem[] = [];
            if (
                obj.options.columns &&
                obj.options.columns[column] &&
                ((obj.options.columns[column].type as string) == 'number' ||
                    obj.options.columns[column].type == 'numeric' ||
                    (obj.options.columns[column].type as string) == 'percentage' ||
                    (obj.options.columns[column].type as string) == 'autonumber' ||
                    obj.options.columns[column].type == 'color')
            ) {
                for (let j = 0; j < obj.options.data.length; j++) {
                    temp[j] = [j, Number(obj.options.data[j][column])];
                }
            } else if (
                obj.options.columns &&
                obj.options.columns[column] &&
                (obj.options.columns[column].type == 'calendar' ||
                    obj.options.columns[column].type == 'checkbox' ||
                    obj.options.columns[column].type == 'radio')
            ) {
                for (let j = 0; j < obj.options.data.length; j++) {
                    temp[j] = [j, obj.options.data[j][column]];
                }
            } else {
                for (let j = 0; j < obj.options.data.length; j++) {
                    temp[j] = [j, obj.records[j][column].element.textContent.toLowerCase()];
                }
            }

            // Default sorting method
            if (typeof obj.parent.config.sorting !== 'function') {
                obj.parent.config.sorting = function (direction) {
                    return function (a, b) {
                        const valueA = a[1];
                        const valueB = b[1];

                        if (!direction) {
                            return valueA === '' && valueB !== '' ? 1 : valueA !== '' && valueB === '' ? -1 : valueA > valueB ? 1 : valueA < valueB ? -1 : 0;
                        } else {
                            return valueA === '' && valueB !== '' ? 1 : valueA !== '' && valueB === '' ? -1 : valueA > valueB ? -1 : valueA < valueB ? 1 : 0;
                        }
                    };
                };
            }

            temp = temp.sort(obj.parent.config.sorting(order));

            // Save history
            const newValue = [];
            for (let j = 0; j < temp.length; j++) {
                newValue[j] = temp[j][0];
            }

            // Save history
            obj.history.add({
                action: 'orderBy',
                rows: newValue,
                column: column,
                order: order,
            });

            // Update order
            this.updateOrderArrow(column, order);
            this.updateOrder(newValue);

            // On sort event
            dispatch.call(
                obj,
                'onsort',
                obj,
                column,
                order,
                newValue.map((row) => row)
            );

            return true;
        }
    }

    /**
     * Remove a row from the table. If the row holds the open editor, save and close the editor first.
     */
    // The callers remove a row when the body has more than a page of rows, so it has a first and a last row
    private removeRow(row: ChildNode) {
        const obj = this.worksheet;

        if (obj.editor.isOpen() && row.contains(obj.editor.cell)) {
            obj.editor.close(obj.editor.cell, true);
        }

        obj.tbody.removeChild(row);
    }

    /**
     * Go to a page in a lazyLoading
     */
    loadPage(pageNumber: number) {
        const obj = this.worksheet;

        // Search: the indexes of the rows found, or the rows
        let results: number[] | Worksheet['rows'];

        if ((obj.options.search == true || obj.options.filters == true) && obj.view.results) {
            results = obj.view.results;
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
            if ((obj.options.search == true || obj.options.filters == true) && obj.view.results) {
                obj.tbody.appendChild(obj.rows[results[j] as number].element);
            } else {
                obj.tbody.appendChild(obj.rows[j].element);
            }

            if (obj.tbody.children.length > quantityPerPage) {
                this.removeRow(obj.tbody.firstChild!);
            }
        }

        // The tab stop of the grid can be in a removed row
        updateTabStop.call(obj);
    }

    loadValidation() {
        const obj = this.worksheet;

        if (obj.selection.range) {
            let currentRow = parseInt((obj.tbody.firstChild as HTMLElement).getAttribute('data-y'));
            let selectedRow = parseInt(obj.selection.range[3]);
            let total = obj.rows.length;

            // With search or filters, the pages are pages of the results
            if ((obj.options.search == true || obj.options.filters == true) && obj.view.results) {
                currentRow = obj.view.results.indexOf(currentRow);
                selectedRow = obj.view.results.indexOf(selectedRow);
                total = obj.view.results.length;

                if (selectedRow < 0) {
                    return false;
                }
            }

            const currentPage = currentRow / 100;
            const selectedPage = parseInt(selectedRow / 100);
            const totalPages = parseInt(total / 100);

            if (currentPage != selectedPage && selectedPage <= totalPages) {
                if (Array.prototype.indexOf.call(obj.tbody.children, obj.rows[obj.selection.range[3] as number].element) < 0) {
                    this.loadPage(selectedPage);
                    return true;
                }
            }
        }

        return false;
    }

    loadUp() {
        const obj = this.worksheet;

        // Search: the indexes of the rows found, or the rows
        let results: number[] | Worksheet['rows'];

        if ((obj.options.search == true || obj.options.filters == true) && obj.view.results) {
            results = obj.view.results;
        } else {
            results = obj.rows;
        }
        let test = 0;
        if (results.length > 100) {
            // Get the first element in the page
            let item = parseInt((obj.tbody.firstChild as HTMLElement).getAttribute('data-y'));
            if ((obj.options.search == true || obj.options.filters == true) && obj.view.results) {
                item = (results as number[]).indexOf(item);
            }
            if (item > 0) {
                for (let j = 0; j < 30; j++) {
                    item = item - 1;
                    if (item > -1) {
                        if ((obj.options.search == true || obj.options.filters == true) && obj.view.results) {
                            obj.tbody.insertBefore(obj.rows[results[item] as number].element, obj.tbody.firstChild);
                        } else {
                            obj.tbody.insertBefore(obj.rows[item].element, obj.tbody.firstChild);
                        }
                        if (obj.tbody.children.length > 100) {
                            this.removeRow(obj.tbody.lastChild!);
                            test = 1;
                        }
                    }
                }
            }
        }

        // The tab stop of the grid can be in a removed row
        updateTabStop.call(obj);

        return test;
    }

    loadDown() {
        const obj = this.worksheet;

        // Search: the indexes of the rows found, or the rows
        let results: number[] | Worksheet['rows'];

        if ((obj.options.search == true || obj.options.filters == true) && obj.view.results) {
            results = obj.view.results;
        } else {
            results = obj.rows;
        }
        let test = 0;
        if (results.length > 100) {
            // Get the last element in the page
            let item = parseInt((obj.tbody.lastChild as HTMLElement).getAttribute('data-y'));
            if ((obj.options.search == true || obj.options.filters == true) && obj.view.results) {
                item = (results as number[]).indexOf(item);
            }
            if (item < obj.rows.length - 1) {
                for (let j = 0; j <= 30; j++) {
                    if (item < results.length) {
                        if ((obj.options.search == true || obj.options.filters == true) && obj.view.results) {
                            obj.tbody.appendChild(obj.rows[results[item] as number].element);
                        } else {
                            obj.tbody.appendChild(obj.rows[item].element);
                        }
                        if (obj.tbody.children.length > 100) {
                            this.removeRow(obj.tbody.firstChild!);
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
    }
}
