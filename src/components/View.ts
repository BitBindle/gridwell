import { translate } from '../utils/dictionary.js';
import { Dropdown } from '../utils/dropdown.js';
import { getIdFromColumnName } from '../utils/internalHelpers.js';
import libraryBase from '../utils/libraryBase.js';
import type jspreadsheet from '../jspreadsheet';
import type { CellValue } from '../jspreadsheet';
import type { Worksheet } from '../types/internal';
import { askConfirm } from '../utils/dialogs.js';
import { applyZoom, toScreen, zoomOf } from '../utils/zoom.js';

/**
 * The row that a load of the lazy loading keeps in place, and its offsetTop: the last row that shows for a load at the
 * bottom, the first row after the frozen rows that shows for a load at the top. A hidden row has no offsetParent. Without
 * a row that shows, the result is null.
 */
function anchorRow(tbody: HTMLTableSectionElement, last: boolean, frozen: number) {
    const rows = tbody.children as HTMLCollectionOf<HTMLTableRowElement>;
    // A frozen row sticks at the top, so it does not move with the other rows
    for (let k = 0; k < rows.length - frozen; k++) {
        const row = rows[last ? rows.length - 1 - k : frozen + k];
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

    /** The scroll position from the top at the last scroll, so the fill handle is checked only for a vertical scroll. */
    private lastScrollTop = 0;

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

        // The rows of the page. An open editor in a row that leaves the page detaches.
        const wanted: HTMLTableRowElement[] = [];
        for (let j = startRow; j < finalRow; j++) {
            if ((worksheet.options.search === true || worksheet.options.filters === true) && worksheet.view.results) {
                wanted.push(worksheet.rows[results[j] as number].element);
            } else {
                wanted.push(worksheet.rows[j].element);
            }
        }
        this.showRows(wanted);

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
        const before = worksheet.options.lazyLoading === true ? this.attachedRows() : null;

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

        // The first rows of the results go in the table. An open editor in a row that leaves detaches.
        const shown = worksheet.view.results ? new Set(worksheet.view.results) : null;
        const wanted: HTMLTableRowElement[] = [];
        for (let j = 0; j < worksheet.rows.length; j++) {
            if (!shown || shown.has(j)) {
                if (index < total) {
                    wanted.push(worksheet.rows[j].element);
                    index++;
                }
                worksheet.rows[j].element.style.display = '';
            } else {
                worksheet.rows[j].element.style.display = 'none';
            }
        }
        this.showRows(wanted);

        // The tab stop of the grid can be in a row that is not in the results
        worksheet.selection.updateTabStop();

        // Update pagination
        if (worksheet.options.pagination! > 0) {
            this.updatePagination();
        }

        worksheet.selection.updateCorner();

        worksheet.dispatch('onupdateresult', worksheet, worksheet.view.results);

        if (before) {
            this.attached(before);
        }

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

        // The application sorts the rows of the row source mode
        if (worksheet.rowSource) {
            return false;
        }

        if (column >= 0) {
            // Merged cells
            if (worksheet.options.mergeCells && Object.keys(worksheet.options.mergeCells).length > 0) {
                if (!askConfirm(worksheet.parent.config, translate('This action will destroy any existing merged cells. Are you sure?'))) {
                    return false;
                } else {
                    // Remove merged cells. The selection of the user stays.
                    worksheet.merges.destroy(false, true);
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
     * The number of frozen rows that lazy loading keeps first in the table body. With the results of a search or the
     * filters, the body holds only the results, so it keeps no frozen row.
     */
    private frozenRows() {
        return this.showsResults() ? 0 : this.worksheet.rowManager.getFreezeCount();
    }

    /**
     * True when the table body holds the results of a search or the filters.
     */
    private showsResults() {
        const worksheet = this.worksheet;
        return (worksheet.options.search === true || worksheet.options.filters === true) && !!worksheet.view.results;
    }

    /**
     * Put the frozen rows first in the table body, and the other attached rows after them in their order. Lazy loading
     * needs this after an insert, a move or a delete in the frozen rows.
     * @param left - The rows that were frozen before the change and are not frozen now.
     */
    attachFrozen(left: HTMLTableRowElement[] = []) {
        const worksheet = this.worksheet;
        const frozen = this.frozenRows();
        // The rows of the results have gaps in data-y, so the runs below do not apply to them. When the last frozen
        // rows leave, they leave the table as the other rows do.
        if (worksheet.options.lazyLoading !== true || this.showsResults() || (!frozen && !left.length)) {
            return;
        }
        const before = this.attachedRows();

        const tbody = worksheet.tbody;
        function y(row: Element) {
            return parseInt(row.getAttribute('data-y'));
        }
        const others = Array.from(tbody.children)
            .filter((row) => y(row) >= frozen)
            .sort((a, b) => y(a) - y(b));

        // The runs of consecutive rows after the frozen rows
        const runs: Element[][] = [];
        for (const row of others) {
            const run = runs[runs.length - 1];
            if (run && y(row) === y(run[run.length - 1]) + 1) {
                run.push(row);
            } else {
                runs.push([row]);
            }
        }

        // A run that holds a row that left the frozen rows, with the rows inserted before it, is not next to the rows that
        // scroll, so it leaves the table. Without such a run, the longest run stays.
        let keep = runs;
        if (runs.length > 1) {
            const scrolled = runs.filter((run) => !run.some((row) => left.includes(row as HTMLTableRowElement)));
            keep = scrolled.length ? scrolled : runs;
            if (keep.length > 1) {
                keep = [keep.reduce((longest, run) => (run.length > longest.length ? run : longest))];
            }
        }
        for (const run of runs) {
            if (!keep.includes(run)) {
                for (const row of run) {
                    this.removeRow(row);
                }
            }
        }

        const wanted = [...worksheet.rows.slice(0, frozen).map((row) => row.element), ...keep.flat()];
        for (let k = 0; k < wanted.length; k++) {
            if (tbody.children[k] !== wanted[k]) {
                tbody.insertBefore(wanted[k], tbody.children[k] || null);
            }
        }

        this.attached(before);
    }

    /**
     * The first row of the table body that lazy loading can remove or insert before: the row after the frozen rows.
     */
    firstScrollRow() {
        const tbody = this.worksheet.tbody;
        return (tbody.children[this.frozenRows()] || tbody.lastChild) as HTMLTableRowElement;
    }

    /**
     * Put these rows in the table body, in this order, and remove the other rows. The rows that stay are not removed
     * first, so an open editor in such a row stays open. An editor in a row that leaves detaches.
     */
    private showRows(wanted: HTMLTableRowElement[]) {
        const tbody = this.worksheet.tbody;
        const keep = new Set<Element>(wanted);
        for (const row of Array.from(tbody.children)) {
            if (!keep.has(row)) {
                this.removeRow(row);
            }
        }
        // A cursor on the next row of the body, because the list of the children changes after each insert
        let next = tbody.firstChild;
        for (const row of wanted) {
            if (row === next) {
                next = row.nextSibling;
            } else {
                tbody.insertBefore(row, next);
            }
        }
        // An editor that was open in a row outside the body, for example after a move, detaches too
        this.worksheet.editor.settle();
    }

    /**
     * Remove a row from the table. If the row holds the open editor, detach the editor first: oneditordetach can take the
     * edit, else the editor saves.
     */
    private removeRow(row: ChildNode) {
        const worksheet = this.worksheet;

        if (worksheet.editor.isOpen() && row.contains(worksheet.editor.cell)) {
            worksheet.editor.detach();
        }

        worksheet.tbody.removeChild(row);
    }

    /**
     * The rows in the table body, before a load of the lazy loading.
     */
    private attachedRows() {
        return new Set<Element>(this.worksheet.tbody.children);
    }

    /**
     * Tell the application the rows that a load of the lazy loading put in the table body, and that were not in it before
     * the load, as [from, to] pairs of row indexes, with onrowsattached.
     * @param before - The rows in the table body before the load.
     */
    private attached(before: Set<Element>) {
        const worksheet = this.worksheet;
        const indexes = Array.from(worksheet.tbody.children)
            .filter((row) => !before.has(row))
            .map((row) => parseInt(row.getAttribute('data-y')!))
            .sort((a, b) => a - b);
        const ranges: [number, number][] = [];
        for (const y of indexes) {
            const last = ranges[ranges.length - 1];
            if (last && last[1] + 1 >= y) {
                last[1] = Math.max(last[1], y);
            } else {
                ranges.push([y, y]);
            }
        }
        if (ranges.length) {
            worksheet.dispatch('onrowsattached', worksheet, ranges);
        }
    }

    /**
     * Fill the table body with the frozen rows and the rows from 10 rows above a row, up to a page of 100 rows, or with
     * the last rows. With a search or the filters, the rows are rows of the results, from the first result at or after the
     * row. The scroll position does not change. Without lazy loading, nothing occurs.
     * @param y - Index of the row.
     */
    loadRow(y: number) {
        const worksheet = this.worksheet;
        if (worksheet.options.lazyLoading !== true || typeof y !== 'number' || !Number.isFinite(y)) {
            return;
        }
        const before = this.attachedRows();

        const quantityPerPage = 100;
        const results = (worksheet.options.search === true || worksheet.options.filters === true) && worksheet.view.results ? worksheet.view.results : null;
        const total = results ? results.length : worksheet.rows.length;
        const frozen = this.frozenRows();
        let position = Math.floor(y);
        if (results) {
            position = results.findIndex((row) => row >= position);
            if (position < 0) {
                position = total;
            }
        }
        const start = Math.max(frozen, Math.min(position - 10, total - (quantityPerPage - frozen)));
        const end = Math.min(total, start + quantityPerPage - frozen);

        const wanted: HTMLTableRowElement[] = [];
        for (let j = 0; j < frozen; j++) {
            wanted.push(worksheet.rows[j].element);
        }
        for (let j = start; j < end; j++) {
            wanted.push(worksheet.rows[results ? results[j] : j].element);
        }

        this.showRows(wanted);

        // The tab stop of the grid can be in a removed row, and the selection corner can be in another place
        worksheet.selection.updateTabStop();
        worksheet.selection.updateCorner();

        this.attached(before);
    }

    /**
     * Change the zoom factor of the spreadsheet, and apply it to the built worksheets. The frozen columns and the
     * selection corner use the factor, so they are placed again.
     */
    setZoom(zoom: number) {
        const spreadsheet = this.worksheet.parent;
        const factor = Number(zoom);
        const next = Number.isFinite(factor) && factor > 0 ? factor : 1;
        for (const worksheet of spreadsheet.worksheets) {
            // A pop-up editor at a fixed place does not move with its cell, so it saves and closes, as after a scroll
            if (worksheet.editor?.isOpen() && worksheet.view.hasFloatingEditor(worksheet.editor.x)) {
                worksheet.editor.close(worksheet.editor.cell, true);
            }
        }
        const previous = spreadsheet.worksheets[0] ? zoomOf(spreadsheet.worksheets[0]) : 1;
        spreadsheet.config.zoom = next;
        for (const worksheet of spreadsheet.worksheets) {
            if (worksheet.zoomed) {
                applyZoom(worksheet);
                // The scroll position is in screen pixels, so the same cells stay in view
                worksheet.content.scrollLeft = (worksheet.content.scrollLeft * next) / previous;
                worksheet.content.scrollTop = (worksheet.content.scrollTop * next) / previous;
                worksheet.columns.updateFreezePosition();
                // The view holds another number of rows and columns
                if (worksheet.rowSource) {
                    worksheet.renderRows();
                }
                worksheet.selection.updateCorner();
            }
        }
    }

    /**
     * Go to a page in a lazyLoading
     */
    loadPage(pageNumber: number) {
        const worksheet = this.worksheet;
        const before = this.attachedRows();

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

        // The frozen rows stay first. Only the rows outside the new page leave, so an editor in a row of the new page
        // stays open. More than one page can be attached before, for example after an insert next to the attached rows.
        const frozen = this.frozenRows();
        const wanted: HTMLTableRowElement[] = [];
        for (let j = 0; j < frozen; j++) {
            wanted.push(worksheet.rows[j].element);
        }
        // A page has 100 rows with the frozen rows
        for (let j = Math.max(startRow, frozen, finalRow - (quantityPerPage - frozen)); j < finalRow; j++) {
            if ((worksheet.options.search === true || worksheet.options.filters === true) && worksheet.view.results) {
                wanted.push(worksheet.rows[results[j] as number].element);
            } else {
                wanted.push(worksheet.rows[j].element);
            }
        }
        this.showRows(wanted);

        // The tab stop of the grid can be in a removed row
        worksheet.selection.updateTabStop();

        this.attached(before);
    }

    loadValidation() {
        const worksheet = this.worksheet;

        if (worksheet.selection.range) {
            let currentRow = parseInt(this.firstScrollRow().getAttribute('data-y'));
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
        const before = this.attachedRows();

        // Search: the indexes of the rows found, or the rows
        let results: number[] | Worksheet['rows'];

        if ((worksheet.options.search === true || worksheet.options.filters === true) && worksheet.view.results) {
            results = worksheet.view.results;
        } else {
            results = worksheet.rows;
        }
        let test = 0;
        if (results.length > 100) {
            // Get the first element in the page, after the frozen rows
            const frozen = this.frozenRows();
            let item = parseInt(this.firstScrollRow().getAttribute('data-y'));
            if ((worksheet.options.search === true || worksheet.options.filters === true) && worksheet.view.results) {
                item = (results as number[]).indexOf(item);
            }
            if (item > frozen) {
                const anchor = anchorRow(worksheet.tbody, false, frozen);
                for (let j = 0; j < 30; j++) {
                    item = item - 1;
                    if (item >= frozen) {
                        if ((worksheet.options.search === true || worksheet.options.filters === true) && worksheet.view.results) {
                            worksheet.tbody.insertBefore(worksheet.rows[results[item] as number].element, this.firstScrollRow());
                        } else {
                            worksheet.tbody.insertBefore(worksheet.rows[item].element, this.firstScrollRow());
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

        this.attached(before);

        return test;
    }

    loadDown() {
        const worksheet = this.worksheet;
        const before = this.attachedRows();

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
                const anchor = anchorRow(worksheet.tbody, true, this.frozenRows());
                for (let j = 0; j <= 30; j++) {
                    if (item < results.length) {
                        if ((worksheet.options.search === true || worksheet.options.filters === true) && worksheet.view.results) {
                            worksheet.tbody.appendChild(worksheet.rows[results[item] as number].element);
                        } else {
                            worksheet.tbody.appendChild(worksheet.rows[item].element);
                        }
                        if (worksheet.tbody.children.length > 100) {
                            this.removeRow(this.firstScrollRow());
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

        this.attached(before);

        return test;
    }

    /**
     * Move the view by the distance that the anchor row of a load moved, so the rows in the view stay in place. If the
     * load removed the anchor row, the view does not move.
     */
    private keepAnchor(anchor: ReturnType<typeof anchorRow>) {
        const content = this.worksheet.content;
        if (anchor && anchor.row.offsetParent !== null) {
            // offsetTop is in grid pixels, and the scroll position in screen pixels
            content.scrollTop = content.scrollTop + toScreen(this.worksheet, anchor.row.offsetTop - anchor.top);
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
                    // loadWorksheet can replace the table of the worksheet before the timer ends
                    if (worksheet.view !== this) {
                        libraryBase.jspreadsheet.timeControlLoading = null;
                        return;
                    }
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

        // The row source mode renders its window for the scroll position
        if (worksheet.rowSource) {
            worksheet.rowSource.render(false, true);
            worksheet.rowSource.columns?.render();
        } else {
            this.wheel();
        }

        if (worksheet.options.freezeColumns! > 0 && worksheet.content.scrollLeft !== worksheet.lastScrollLeft) {
            worksheet.lastScrollLeft = worksheet.content.scrollLeft;
            worksheet.columns.updateFreezePosition();
        }

        // The fill handle hides while its cell is under the sticky header or the frozen rows
        if (worksheet.content.scrollTop !== this.lastScrollTop) {
            this.lastScrollTop = worksheet.content.scrollTop;
            worksheet.selection.updateCornerCover();
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
    hasFloatingEditor(x: number) {
        const worksheet = this.worksheet;

        const column = worksheet.options.columns?.[x];
        if (!column?.type) {
            return false;
        }

        return typeof column.type === 'object' || ['dropdown', 'calendar', 'color', 'html', 'image'].includes(column.type);
    }
}
