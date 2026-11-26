import { translate } from '../utils/dictionary.js';
import { Dropdown } from '../utils/dropdown.js';
import { getIdFromColumnName } from '../utils/internalHelpers.js';
import libraryBase from '../utils/libraryBase.js';
import type jspreadsheet from '../jspreadsheet';
import type { CellValue } from '../jspreadsheet';
import type { Worksheet } from '../types/internal';

/**
 * The row that a load of the lazy loading keeps in place, and its offsetTop: the last row that shows for a load at the
 * bottom, the first row that shows for a load at the top. A hidden row has no offsetParent. Without a row that shows,
 * the result is null.
 */
function anchorRow(tbody: HTMLTableSectionElement, last: boolean) {
    const rows = tbody.children as HTMLCollectionOf<HTMLTableRowElement>;
    for (let k = 0; k < rows.length; k++) {
        const row = rows[last ? rows.length - 1 - k : k];
        if (row.offsetParent !== null) {
            return { row, top: row.offsetTop };
        }
    }
    return null;
}

/**
 * Which rows a worksheet shows: the search, the filters, the pagination, the order of the rows, the lazy loading and the scrolling. It
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

    /** The scroll position after the last check or load of the lazy loading. A change of it gives the direction of a scroll. */
    private lazyScrollTop = 0;

    /** The direction of the last wheel event since the last check of the lazy loading: 1 down, -1 up, 0 none. */
    private lazyWheel = 0;

    constructor(worksheet: Worksheet) {
        this.worksheet = worksheet;
    }

    /**
     * Search
     * @param query - The text to find. An empty text shows all rows again.
     */
    search(query: string) {
        const worksheet = this.worksheet;

        // Reset any filter
        if (worksheet.options.filters) {
            this.resetFilters();
        }

        // Reset selection
        worksheet.resetSelection();

        // Total of results
        worksheet.view.pageNumber = 0;
        worksheet.view.results = [];

        if (query) {
            if (worksheet.searchInput.value !== query) {
                worksheet.searchInput.value = query;
            }

            /**
             * Search filter
             * @param item - The data of the row.
             * @param index - The row.
             */
            function search(item: CellValue[], query: RegExp, index: number) {
                for (let i = 0; i < item.length; i++) {
                    if (`${item[i]}`.toLowerCase().search(query) >= 0 || `${worksheet.records[index][i].element.innerHTML}`.toLowerCase().search(query) >= 0) {
                        return true;
                    }
                }
                return false;
            }

            /**
             * Result
             * @param k - The row.
             */
            function addToResult(k: number) {
                if (!worksheet.view.results!.includes(k)) {
                    worksheet.view.results!.push(k);
                }
            }

            let parsedQuery: string | RegExp = query.replace(/[-[\]{}()*+?.,\\^$|#\s]/g, '\\$&');
            parsedQuery = new RegExp(parsedQuery, 'i');

            // Filter
            worksheet.options.data.forEach((v, k) => {
                if (search(v, parsedQuery, k)) {
                    // Merged rows found
                    const rows = worksheet.merges.isRowMerged(k);
                    if (rows.length) {
                        for (const cellName of rows) {
                            const row = getIdFromColumnName(cellName);
                            for (let j = 0; j < worksheet.options.mergeCells![cellName][1]; j++) {
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
            worksheet.view.results = null;
        }

        this.updateResult();
    }

    /**
     * Reset search
     */
    resetSearch() {
        const worksheet = this.worksheet;

        worksheet.searchInput.value = '';
        worksheet.search('');
        worksheet.view.results = null;
    }

    /**
     * Open the column filter
     */
    openFilter(columnId: number) {
        const worksheet = this.worksheet;

        if (!worksheet.options.filters) {
            console.log('Jspreadsheet: filters not enabled.');
        } else {
            // Make sure is integer
            columnId = parseInt(columnId);
            // Reset selection
            worksheet.resetSelection();
            // Load options
            let optionsFiltered = [];
            if (worksheet.options.columns[columnId].type === 'checkbox') {
                optionsFiltered.push({ id: 'true', name: 'True' });
                optionsFiltered.push({ id: 'false', name: 'False' });
            } else {
                // Labels by value. An array used as a map, as before.
                const options = [] as unknown as Record<string, string>;
                let hasBlanks = false;
                for (let j = 0; j < worksheet.options.data.length; j++) {
                    const k = worksheet.options.data[j][columnId];
                    // The text of the cell, which the dropdown shows as text
                    const v = worksheet.records[j][columnId].element.textContent;
                    if (k && v) {
                        options[k as string] = v;
                    } else {
                        hasBlanks = true;
                    }
                }
                optionsFiltered = [];
                for (const [id, name] of Object.entries(options)) {
                    optionsFiltered.push({ id, name });
                }
                // Has blank options
                if (hasBlanks) {
                    optionsFiltered.push({ id: '', name: translate('(Blanks)') });
                }
            }

            // The cell of the filter, which shows the dropdown while it is open
            const filterCell = worksheet.filter!.children[columnId + 1] as HTMLElement;
            const previous = filterCell.textContent;
            function restore(text: string | null) {
                filterCell.textContent = text;
                filterCell.style.paddingLeft = '';
                filterCell.style.paddingRight = '';
                filterCell.style.overflow = '';
            }

            const div = document.createElement('div');
            filterCell.textContent = '';
            filterCell.appendChild(div);
            filterCell.style.paddingLeft = '0px';
            filterCell.style.paddingRight = '0px';
            filterCell.style.overflow = 'initial';

            const dropdown = new Dropdown(div, {
                data: optionsFiltered,
                multiple: true,
                autocomplete: true,
                value: worksheet.view.filters[columnId] !== undefined ? worksheet.view.filters[columnId] : null,
                width: '100%',
                position: worksheet.options.tableOverflow === true || worksheet.parent.config.fullscreen === true,
                onclose: () => {
                    worksheet.view.resetFilters();
                    worksheet.view.filters[columnId] = dropdown.getValue(true);
                    restore(dropdown.getText());
                    worksheet.view.closeFilter(columnId);
                    worksheet.selection.refresh();
                },
                // Escape leaves the filters as they were
                oncancel: () => {
                    restore(previous);
                },
            });
            dropdown.header.setAttribute('aria-label', worksheet.accessibility.getLabelText('filterColumn', worksheet.getHeader(columnId)));
            dropdown.open();
        }
    }

    closeFilter(columnId?: number) {
        const worksheet = this.worksheet;

        if (!columnId) {
            for (let i = 0; i < worksheet.filter!.children.length; i++) {
                if (worksheet.view.filters[i]) {
                    columnId = i;
                }
            }
        }

        // Search filter
        function search(query: string[], x: number, y: number) {
            for (const term of query) {
                const value = `${worksheet.options.data[y][x]}`;
                const label = `${worksheet.records[y][x].element.innerHTML}`;
                if (term === value || term === label) {
                    return true;
                }
            }
            return false;
        }

        // The column has a filter: openFilter passes its column, and orderBy calls this only while a filter gives results
        const query = worksheet.view.filters[columnId!]!;
        worksheet.view.results = [];
        for (let j = 0; j < worksheet.options.data.length; j++) {
            if (search(query, columnId!, j)) {
                worksheet.view.results.push(j);
            }
        }
        if (!worksheet.view.results.length) {
            worksheet.view.results = null;
        }

        this.updateResult();
    }

    resetFilters() {
        const worksheet = this.worksheet;

        if (worksheet.options.filters) {
            for (let i = 0; i < worksheet.filter!.children.length; i++) {
                worksheet.filter!.children[i].innerHTML = '&nbsp;';
                worksheet.view.filters[i] = null;
            }
        }

        worksheet.view.results = null;
        this.updateResult();
    }

    /**
     * Which page the row is
     */
    whichPage(row: number) {
        const worksheet = this.worksheet;

        // Search
        if ((worksheet.options.search === true || worksheet.options.filters === true) && worksheet.view.results) {
            row = worksheet.view.results.indexOf(row);
        }

        return Math.ceil((parseInt(row) + 1) / parseInt(worksheet.options.pagination)) - 1;
    }

    /**
     * Update the pagination
     */
    updatePagination() {
        const worksheet = this.worksheet;

        // Reset container
        worksheet.pagination.children[0].innerHTML = '';
        worksheet.pagination.children[1].innerHTML = '';

        // Start pagination
        if (worksheet.options.pagination) {
            // Searchable
            let results: number;

            if ((worksheet.options.search === true || worksheet.options.filters === true) && worksheet.view.results) {
                results = worksheet.view.results.length;
            } else {
                results = worksheet.rows.length;
            }

            if (!results) {
                // No records found
                worksheet.pagination.children[0].innerHTML = translate('No records found');
            } else {
                // Pagination container
                const quantyOfPages = Math.ceil(results / worksheet.options.pagination);

                let startNumber, finalNumber;

                if (worksheet.view.pageNumber < 6) {
                    startNumber = 1;
                    finalNumber = quantyOfPages < 10 ? quantyOfPages : 10;
                } else if (quantyOfPages - worksheet.view.pageNumber < 5) {
                    startNumber = quantyOfPages - 9;
                    finalNumber = quantyOfPages;
                    if (startNumber < 1) {
                        startNumber = 1;
                    }
                } else {
                    startNumber = worksheet.view.pageNumber - 4;
                    finalNumber = worksheet.view.pageNumber + 5;
                }

                // First
                if (startNumber > 1) {
                    const paginationItem = document.createElement('div');
                    paginationItem.className = 'jss_page';
                    paginationItem.innerHTML = '<';
                    paginationItem.title = '1';
                    worksheet.pagination.children[1].appendChild(paginationItem);
                }

                // Get page links
                for (let i = startNumber; i <= finalNumber; i++) {
                    const paginationItem = document.createElement('div');
                    paginationItem.className = 'jss_page';
                    paginationItem.innerHTML = String(i);
                    worksheet.pagination.children[1].appendChild(paginationItem);

                    if (worksheet.view.pageNumber === i - 1) {
                        paginationItem.classList.add('jss_page_selected');
                    }
                }

                // Last
                if (finalNumber < quantyOfPages) {
                    const paginationItem = document.createElement('div');
                    paginationItem.className = 'jss_page';
                    paginationItem.innerHTML = '>';
                    paginationItem.title = String(quantyOfPages);
                    worksheet.pagination.children[1].appendChild(paginationItem);
                }

                // Text
                function format(format: string, ...args: number[]): string {
                    return format.replace(/{(\d+)}/g, (match: string, number: number) => {
                        return typeof args[number] !== 'undefined' ? String(args[number]) : match;
                    });
                }

                worksheet.pagination.children[0].innerHTML = format(translate('Showing page {0} of {1} entries'), worksheet.view.pageNumber + 1, quantyOfPages);
            }
        }
    }

    /**
     * Go to page
     */
    page(pageNumber: number) {
        const worksheet = this.worksheet;

        const oldPage = worksheet.view.pageNumber;

        // Search: the indexes of the rows found, or the rows
        let results: number[] | Worksheet['rows'];

        if ((worksheet.options.search === true || worksheet.options.filters === true) && worksheet.view.results) {
            results = worksheet.view.results;
        } else {
            results = worksheet.rows;
        }

        // Per page
        const quantityPerPage = parseInt(worksheet.options.pagination);

        // pageNumber
        if (pageNumber == null || pageNumber === -1) {
            // Last page
            pageNumber = Math.ceil(results.length / quantityPerPage) - 1;
        }

        // Page number
        worksheet.view.pageNumber = pageNumber;

        let startRow = pageNumber * quantityPerPage;
        let finalRow = pageNumber * quantityPerPage + quantityPerPage;
        if (finalRow > results.length) {
            finalRow = results.length;
        }
        if (startRow < 0) {
            startRow = 0;
        }

        // Reset container
        while (worksheet.tbody.firstChild) {
            worksheet.tbody.removeChild(worksheet.tbody.firstChild);
        }

        // Appeding items
        for (let j = startRow; j < finalRow; j++) {
            if ((worksheet.options.search === true || worksheet.options.filters === true) && worksheet.view.results) {
                worksheet.tbody.appendChild(worksheet.rows[results[j] as number].element);
            } else {
                worksheet.tbody.appendChild(worksheet.rows[j].element);
            }
        }

        // The tab stop of the grid can be in a row of the old page
        worksheet.selection.updateTabStop();

        if (worksheet.options.pagination! > 0) {
            this.updatePagination();
        }

        // Update corner position
        worksheet.selection.updateCorner();

        // Events. setData sets the page number of a worksheet with pagination.
        worksheet.dispatch('onchangepage', worksheet, pageNumber, oldPage, worksheet.options.pagination!);
    }

    countPages() {
        const worksheet = this.worksheet;

        let results: number;
        if ((worksheet.options.search === true || worksheet.options.filters === true) && worksheet.view.results) {
            results = worksheet.view.results.length;
        } else {
            results = worksheet.rows.length;
        }

        return Math.ceil(results / worksheet.options.pagination!);
    }

    updateResult() {
        const worksheet = this.worksheet;

        let total;
        let index = 0;

        // Page 1
        if (worksheet.options.lazyLoading === true) {
            total = 100;
        } else if (worksheet.options.pagination! > 0) {
            total = worksheet.options.pagination!;
        } else {
            if (worksheet.view.results) {
                total = worksheet.view.results.length;
            } else {
                total = worksheet.rows.length;
            }
        }

        // Reset current nodes
        while (worksheet.tbody.firstChild) {
            worksheet.tbody.removeChild(worksheet.tbody.firstChild);
        }

        // Hide all records from the table
        for (let j = 0; j < worksheet.rows.length; j++) {
            if (!worksheet.view.results || worksheet.view.results.includes(j)) {
                if (index < total) {
                    worksheet.tbody.appendChild(worksheet.rows[j].element);
                    index++;
                }
                worksheet.rows[j].element.style.display = '';
            } else {
                worksheet.rows[j].element.style.display = 'none';
            }
        }

        // The tab stop of the grid can be in a row that is not in the results
        worksheet.selection.updateTabStop();

        // Update pagination
        if (worksheet.options.pagination! > 0) {
            this.updatePagination();
        }

        worksheet.selection.updateCorner();

        worksheet.dispatch('onupdateresult', worksheet, worksheet.view.results);

        return total;
    }

    /**
     * Update order arrow
     */
    updateOrderArrow(column: number, order: boolean | number) {
        const worksheet = this.worksheet;

        // Remove order
        for (const header of worksheet.headers) {
            header.classList.remove('arrow-up');
            header.classList.remove('arrow-down');
        }

        // No order specified then toggle order
        if (order) {
            worksheet.headers[column].classList.add('arrow-up');
        } else {
            worksheet.headers[column].classList.add('arrow-down');
        }
    }

    /**
     * Update rows position
     */
    updateOrder(rows: number[]) {
        const worksheet = this.worksheet;

        // History
        const data: Worksheet['options']['data'] = [];
        for (let j = 0; j < rows.length; j++) {
            data[j] = worksheet.options.data[rows[j]];
        }
        worksheet.options.data = data;

        const records: Worksheet['records'] = [];
        for (let j = 0; j < rows.length; j++) {
            records[j] = worksheet.records[rows[j]];

            for (const cell of records[j]) {
                cell.y = j;
            }
        }
        worksheet.records = records;

        const rowObjects: Worksheet['rows'] = [];
        for (let j = 0; j < rows.length; j++) {
            rowObjects[j] = worksheet.rows[rows[j]];
            rowObjects[j].y = j;
        }
        worksheet.rows = rowObjects;

        // Update references
        worksheet.formulas.updateTableReferences();

        // Redo search
        if (worksheet.view.results?.length) {
            if (worksheet.searchInput.value) {
                worksheet.search(worksheet.searchInput.value);
            } else {
                this.closeFilter();
            }
        } else {
            // Create page
            worksheet.view.results = null;
            worksheet.view.pageNumber = 0;

            if (worksheet.options.pagination! > 0) {
                worksheet.page(0);
            } else if (worksheet.options.lazyLoading === true) {
                this.loadPage(0);
            } else {
                for (const row of worksheet.rows) {
                    worksheet.tbody.appendChild(row.element);
                }
                // A moved row loses the focus
                worksheet.selection.updateTabStop();
            }
        }
    }

    /**
     * Sort data and reload table
     */
    orderBy(column: number, order?: 0 | 1): boolean | undefined {
        const worksheet = this.worksheet;

        if (column >= 0) {
            // Merged cells
            if (worksheet.options.mergeCells && Object.keys(worksheet.options.mergeCells).length > 0) {
                if (!confirm(translate('This action will destroy any existing merged cells. Are you sure?'))) {
                    return false;
                } else {
                    // Remove merged cells
                    worksheet.destroyMerge();
                }
            }

            // Direction
            if (order == null) {
                order = worksheet.headers[column].classList.contains('arrow-down') ? 1 : 0;
            } else {
                order = order ? 1 : 0;
            }

            // Test order
            // The row and the value to sort by
            let temp: jspreadsheet.SortingItem[] = [];
            if (
                worksheet.options.columns?.[column] &&
                ((worksheet.options.columns[column].type as string) === 'number' ||
                    worksheet.options.columns[column].type === 'numeric' ||
                    (worksheet.options.columns[column].type as string) === 'percentage' ||
                    (worksheet.options.columns[column].type as string) === 'autonumber' ||
                    worksheet.options.columns[column].type === 'color')
            ) {
                for (let j = 0; j < worksheet.options.data.length; j++) {
                    temp[j] = [j, Number(worksheet.options.data[j][column])];
                }
            } else if (
                worksheet.options.columns?.[column] &&
                (worksheet.options.columns[column].type === 'calendar' ||
                    worksheet.options.columns[column].type === 'checkbox' ||
                    worksheet.options.columns[column].type === 'radio')
            ) {
                for (let j = 0; j < worksheet.options.data.length; j++) {
                    temp[j] = [j, worksheet.options.data[j][column]];
                }
            } else {
                for (let j = 0; j < worksheet.options.data.length; j++) {
                    temp[j] = [j, worksheet.records[j][column].element.textContent.toLowerCase()];
                }
            }

            // Default sorting method
            if (typeof worksheet.parent.config.sorting !== 'function') {
                worksheet.parent.config.sorting = (direction) => {
                    return (a, b) => {
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

            temp = temp.sort(worksheet.parent.config.sorting(order));

            // Save history
            const newValue = [];
            for (let j = 0; j < temp.length; j++) {
                newValue[j] = temp[j][0];
            }

            // Save history
            worksheet.history.add({
                action: 'orderBy',
                rows: newValue,
                column,
                order,
            });

            // Update order
            this.updateOrderArrow(column, order);
            this.updateOrder(newValue);

            // On sort event
            worksheet.dispatch(
                'onsort',
                worksheet,
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
        const worksheet = this.worksheet;

        if (worksheet.editor.isOpen() && row.contains(worksheet.editor.cell)) {
            worksheet.editor.close(worksheet.editor.cell, true);
        }

        worksheet.tbody.removeChild(row);
    }

    /**
     * Go to a page in a lazyLoading
     */
    loadPage(pageNumber: number) {
        const worksheet = this.worksheet;

        // Search: the indexes of the rows found, or the rows
        let results: number[] | Worksheet['rows'];

        if ((worksheet.options.search === true || worksheet.options.filters === true) && worksheet.view.results) {
            results = worksheet.view.results;
        } else {
            results = worksheet.rows;
        }

        // Per page
        const quantityPerPage = 100;

        // pageNumber
        if (pageNumber == null || pageNumber === -1) {
            // Last page
            pageNumber = Math.ceil(results.length / quantityPerPage) - 1;
        }

        let finalRow = pageNumber * quantityPerPage + quantityPerPage;
        if (finalRow > results.length) {
            finalRow = results.length;
        }
        let startRow = finalRow - 100;
        if (startRow < 0) {
            startRow = 0;
        }

        // Appeding items
        for (let j = startRow; j < finalRow; j++) {
            if ((worksheet.options.search === true || worksheet.options.filters === true) && worksheet.view.results) {
                worksheet.tbody.appendChild(worksheet.rows[results[j] as number].element);
            } else {
                worksheet.tbody.appendChild(worksheet.rows[j].element);
            }

            // More than one page can be attached, for example after an insert next to the attached rows
            while (worksheet.tbody.children.length > quantityPerPage) {
                this.removeRow(worksheet.tbody.firstChild!);
            }
        }

        // The tab stop of the grid can be in a removed row
        worksheet.selection.updateTabStop();
    }

    loadValidation() {
        const worksheet = this.worksheet;

        if (worksheet.selection.range) {
            let currentRow = parseInt((worksheet.tbody.firstChild as HTMLElement).getAttribute('data-y'));
            let selectedRow = parseInt(worksheet.selection.range[3]);
            let total = worksheet.rows.length;

            // With search or filters, the pages are pages of the results
            if ((worksheet.options.search === true || worksheet.options.filters === true) && worksheet.view.results) {
                currentRow = worksheet.view.results.indexOf(currentRow);
                selectedRow = worksheet.view.results.indexOf(selectedRow);
                total = worksheet.view.results.length;

                if (selectedRow < 0) {
                    return false;
                }
            }

            const currentPage = currentRow / 100;
            const selectedPage = parseInt(selectedRow / 100);
            const totalPages = parseInt(total / 100);

            if (currentPage !== selectedPage && selectedPage <= totalPages) {
                if (worksheet.rows[worksheet.selection.range[3] as number].element.parentNode !== worksheet.tbody) {
                    this.loadPage(selectedPage);
                    return true;
                }
            }
        }

        return false;
    }

    loadUp() {
        const worksheet = this.worksheet;

        // Search: the indexes of the rows found, or the rows
        let results: number[] | Worksheet['rows'];

        if ((worksheet.options.search === true || worksheet.options.filters === true) && worksheet.view.results) {
            results = worksheet.view.results;
        } else {
            results = worksheet.rows;
        }
        let test = 0;
        if (results.length > 100) {
            // Get the first element in the page
            let item = parseInt((worksheet.tbody.firstChild as HTMLElement).getAttribute('data-y'));
            if ((worksheet.options.search === true || worksheet.options.filters === true) && worksheet.view.results) {
                item = (results as number[]).indexOf(item);
            }
            if (item > 0) {
                const anchor = anchorRow(worksheet.tbody, false);
                for (let j = 0; j < 30; j++) {
                    item = item - 1;
                    if (item > -1) {
                        if ((worksheet.options.search === true || worksheet.options.filters === true) && worksheet.view.results) {
                            worksheet.tbody.insertBefore(worksheet.rows[results[item] as number].element, worksheet.tbody.firstChild);
                        } else {
                            worksheet.tbody.insertBefore(worksheet.rows[item].element, worksheet.tbody.firstChild);
                        }
                        if (worksheet.tbody.children.length > 100) {
                            this.removeRow(worksheet.tbody.lastChild!);
                            test = 1;
                        }
                    }
                }
                this.keepAnchor(anchor);
            }
        }

        // The tab stop of the grid can be in a removed row
        worksheet.selection.updateTabStop();

        return test;
    }

    loadDown() {
        const worksheet = this.worksheet;

        // Search: the indexes of the rows found, or the rows
        let results: number[] | Worksheet['rows'];

        if ((worksheet.options.search === true || worksheet.options.filters === true) && worksheet.view.results) {
            results = worksheet.view.results;
        } else {
            results = worksheet.rows;
        }
        let test = 0;
        if (results.length > 100) {
            // Get the last element in the page
            let item = parseInt((worksheet.tbody.lastChild as HTMLElement).getAttribute('data-y'));
            if ((worksheet.options.search === true || worksheet.options.filters === true) && worksheet.view.results) {
                item = (results as number[]).indexOf(item);
            }
            if (item < worksheet.rows.length - 1) {
                const anchor = anchorRow(worksheet.tbody, true);
                for (let j = 0; j <= 30; j++) {
                    if (item < results.length) {
                        if ((worksheet.options.search === true || worksheet.options.filters === true) && worksheet.view.results) {
                            worksheet.tbody.appendChild(worksheet.rows[results[item] as number].element);
                        } else {
                            worksheet.tbody.appendChild(worksheet.rows[item].element);
                        }
                        if (worksheet.tbody.children.length > 100) {
                            this.removeRow(worksheet.tbody.firstChild!);
                            test = 1;
                        }
                    }
                    item = item + 1;
                }
                this.keepAnchor(anchor);
            }
        }

        // The tab stop of the grid can be in a removed row
        worksheet.selection.updateTabStop();

        return test;
    }

    /**
     * Move the view by the distance that the anchor row of a load moved, so the rows in the view stay in place. If the
     * load removed the anchor row, the view does not move.
     */
    private keepAnchor(anchor: ReturnType<typeof anchorRow>) {
        const content = this.worksheet.content;
        if (anchor && anchor.row.offsetParent !== null) {
            content.scrollTop = content.scrollTop + anchor.row.offsetTop - anchor.top;
        }
        // The browser can clamp the position, so the next check must see no scroll
        this.lazyScrollTop = content.scrollTop;
    }

    /**
     * Load more rows of the lazy loading when the content is scrolled near its top or its bottom. The rows load only in
     * the direction of the scroll, so a load at the bottom and a load at the top cannot alternate.
     */
    wheel(e?: WheelEvent) {
        const worksheet = this.worksheet;

        if (worksheet.options.lazyLoading === true) {
            if (e?.type === 'wheel' && e.deltaY) {
                this.lazyWheel = Math.sign(e.deltaY);
            }

            if (libraryBase.jspreadsheet.timeControlLoading == null) {
                libraryBase.jspreadsheet.timeControlLoading = setTimeout(() => {
                    const content = worksheet.content;
                    // The last wheel event, else the change of the scroll position since the last check or load
                    const direction = this.lazyWheel || Math.sign(content.scrollTop - this.lazyScrollTop);
                    this.lazyWheel = 0;
                    this.lazyScrollTop = content.scrollTop;

                    if (direction > 0 && content.scrollTop + content.clientHeight >= content.scrollHeight - 10) {
                        if (this.loadDown()) {
                            worksheet.selection.updateCorner();
                        }
                    } else if (direction < 0 && content.scrollTop < 10) {
                        if (this.loadUp()) {
                            worksheet.selection.updateCorner();
                        }
                    }

                    libraryBase.jspreadsheet.timeControlLoading = null;
                }, 100);
            }
        }
    }

    /**
     * The content of the worksheet scrolls: load more rows, place the frozen columns and close a floating editor.
     */
    scroll(e: Event) {
        const worksheet = this.worksheet;

        this.wheel();

        if (worksheet.options.freezeColumns! > 0 && worksheet.content.scrollLeft !== worksheet.lastScrollLeft) {
            worksheet.lastScrollLeft = worksheet.content.scrollLeft;
            worksheet.columns.updateFreezePosition();
        }

        // Close a floating editor. The text editor is inside the cell and scrolls with it.
        if (worksheet.options.lazyLoading === true || worksheet.options.tableOverflow === true) {
            if (worksheet.editor.isOpen() && this.hasFloatingEditor(worksheet.editor.x) && (e.target as HTMLElement).className.substr(0, 9) !== 'jdropdown') {
                worksheet.editor.close(worksheet.editor.cell, true);
            }
        }
    }

    /**
     * Check if the editor of a column floats over the page. It does not move with the cell.
     */
    private hasFloatingEditor(x: number) {
        const worksheet = this.worksheet;

        const column = worksheet.options.columns?.[x];
        if (!column?.type) {
            return false;
        }

        return typeof column.type === 'object' || ['dropdown', 'calendar', 'color', 'html', 'image'].includes(column.type);
    }
}
