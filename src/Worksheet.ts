import { translate } from './utils/dictionary.js';
import { request } from './utils/request.js';
import libraryBase from './utils/libraryBase.js';
import { parseCSV } from './utils/helpers.js';
import { getIdFromColumnName } from './utils/internalHelpers.js';
import { cutControls } from './events/clipboard.js';
import type jspreadsheet from './jspreadsheet';
import { Accessibility } from './components/Accessibility.js';
import { Cells } from './components/Cells.js';
import { Clipboard } from './components/Clipboard.js';
import { Columns } from './components/Columns.js';
import { Comments } from './components/Comments.js';
import { Data } from './components/Data.js';
import { Editor } from './components/Editor.js';
import { Formulas } from './components/Formulas.js';
import { History } from './components/History.js';
import { Keyboard } from './components/Keyboard.js';
import { Merges } from './components/Merges.js';
import { Meta } from './components/Meta.js';
import { Pointer } from './components/Pointer.js';
import { Rows } from './components/Rows.js';
import { Selection } from './components/Selection.js';
import { Styles } from './components/Styles.js';
import { View } from './components/View.js';
import type { EventHandlers, EventName, Spreadsheet, WorksheetState } from './types/internal';

/** The rows that persistence posts to the server, the data of onbeforesave */
type SaveData = Parameters<NonNullable<jspreadsheet.SpreadsheetOptions['onbeforesave']>>[2];

/**
 * The fields of a worksheet and its public methods. The fields are in WorksheetState (src/types/internal.d.ts). The
 * public methods are methods of the class. Most of them call the method of a component.
 */
// eslint-disable-next-line @typescript-eslint/no-empty-object-type, @typescript-eslint/no-unsafe-declaration-merging
export interface Worksheet extends WorksheetState {}

/**
 * A worksheet: one table of a spreadsheet, with its options, its DOM and its state.
 *
 * The constructor sets up the state that the worksheet has before it is built. build then creates the DOM from the
 * options.
 */
// eslint-disable-next-line @typescript-eslint/no-unsafe-declaration-merging
export class Worksheet {
    constructor(spreadsheet: Spreadsheet, options: jspreadsheet.WorksheetOptions, element?: HTMLElement) {
        this.parent = spreadsheet;
        // build fills in the options that the internal type says are always there
        this.options = options as Worksheet['options'];
        if (element) {
            // build adds jssWorksheet
            this.element = element as Worksheet['element'];
        }
        this.history = new History(this);
        this.clipboard = new Clipboard(this);
        this.editor = new Editor(this);
        this.selection = new Selection(this);
        this.formulas = new Formulas(this);
        this.view = new View(this);
        this.pointer = new Pointer(this);
        this.cells = new Cells(this);
        this.styles = new Styles(this);
        this.comments = new Comments(this);
        this.meta = new Meta(this);
        this.merges = new Merges(this);
        this.columns = new Columns(this);
        this.rowManager = new Rows(this);
        this.data = new Data(this);
        this.accessibility = new Accessibility(this);
        this.keyboard = new Keyboard(this);
        this.lastScrollLeft = 0;
    }

    /**
     * Trigger an event of the worksheet. The handlers get the worksheet as `this`. After onafterchanges, the plugins
     * and the persistence option save the changes.
     *
     * The first argument is the name of the event, for example "onchange". The other arguments go to the event handlers.
     * It returns the result of the last handler, or null when no handler ran.
     */
    dispatch<E extends EventName>(event: E, ...args: Parameters<EventHandlers[E]>): ReturnType<EventHandlers[E]> | null {
        const spreadsheet = this.parent;

        const ret = spreadsheet.callHandlers(this, event, args);

        if (event === 'onafterchanges') {
            if (typeof spreadsheet.plugins === 'object') {
                for (const plugin of Object.values(spreadsheet.plugins)) {
                    if (typeof plugin.persistence === 'function') {
                        // The changes of onafterchanges
                        plugin.persistence(this, 'setValue', { data: args[1] as jspreadsheet.CellChange[] });
                    }
                }
            }

            if (this.options.persistence) {
                const url = this.options.persistence === true ? this.options.url : this.options.persistence;
                const data = this.prepareJson(args[1] as jspreadsheet.CellChange[]);
                this.save(url, data);
            }
        }

        // The result of a handler of the event, as the type of dispatch says
        return ret as ReturnType<EventHandlers[E]> | null;
    }

    /**
     * Prepare JSON in the correct format
     * @param data - The changed cells.
     */
    private prepareJson(data: jspreadsheet.CellChange[]) {
        const rows: SaveData = [];
        for (const change of data) {
            const x = change.x;
            const y = change.y;
            const k = this.options.columns[x].name ? this.options.columns[x].name : x;

            // Create row
            if (!rows[y]) {
                rows[y] = {
                    row: y,
                    data: {},
                };
            }
            // The value is undefined after setValue without a value, and the JSON leaves it out
            rows[y].data[k] = change.value as jspreadsheet.CellValue;
        }

        // Filter rows
        return rows.filter((el) => {
            return el != null;
        });
    }

    /**
     * Post json to a remote server
     */
    // The url is undefined when persistence is true and the worksheet has no url option
    private save(url: string | undefined, data: SaveData) {
        // Parse anything in the data before sending to the server
        const ret = this.parent.dispatch('onbeforesave', this.parent, this, data);
        if (ret) {
            // The application can return any data to send in place of the rows
            data = ret as SaveData;
        } else {
            if (ret === false) {
                return false;
            }
        }

        // Post the data as JSON in the data field of a form
        const body = new FormData();
        body.append('data', JSON.stringify(data));
        // fetch, as jSuites.ajax did, posts to the address "undefined" when there is no url (issue #140)
        request(this, String(url), 'json', body).then((response) => {
            if (response) {
                // Event
                this.dispatch('onsave', this.parent, this, data);
            }
        });
    }

    /**
     * Get the options of the worksheet.
     */
    getConfig() {
        return this.options;
    }

    /**
     * Change options of the worksheet, or of its spreadsheet when spreadsheetLevel is true.
     * @param config - The options to change.
     * @param spreadsheetLevel - Change the options of the spreadsheet of the worksheet.
     */
    setConfig(config: Partial<jspreadsheet.WorksheetOptions>, spreadsheetLevel?: false): void;
    setConfig(config: Partial<jspreadsheet.SpreadsheetOptions>, spreadsheetLevel: true): void;
    setConfig(config: Partial<jspreadsheet.SpreadsheetOptions> | Partial<jspreadsheet.WorksheetOptions>, spreadsheetLevel?: boolean) {
        if (spreadsheetLevel) {
            this.parent.setConfig(config);
            return;
        }

        for (const [key, value] of Object.entries(config)) {
            (this.options as Record<string, unknown>)[key] = value;
        }
    }

    /**
     * The index of the worksheet that the spreadsheet shows.
     */
    getWorksheetActive(): number {
        return this.parent.getWorksheetActive();
    }

    // The public methods that call a component

    selectAll() {
        return this.selection.selectAll();
    }

    updateSelectionFromCoords(x1: number | null, y1: number | null, x2?: number | null, y2?: number | null) {
        return this.selection.updateFromCoords(x1, y1, x2, y2);
    }

    resetSelection() {
        return this.selection.reset();
    }

    getSelection() {
        return this.selection.get();
    }

    getSelected(columnNameOnly?: false): { element: HTMLTableCellElement; x: number; y: number }[];
    getSelected(columnNameOnly: true): string[];
    getSelected(columnNameOnly?: boolean): { element: HTMLTableCellElement; x: number; y: number }[] | string[];
    getSelected(columnNameOnly?: boolean) {
        return this.selection.getSelected(columnNameOnly);
    }

    getSelectedColumns(visibleOnly?: boolean) {
        return this.selection.getSelectedColumns(visibleOnly);
    }

    getSelectedRows(visibleOnly?: boolean) {
        return this.selection.getSelectedRows(visibleOnly);
    }

    getData(highlighted?: boolean, processed?: boolean, delimiter?: undefined | null | '', asJson?: false): jspreadsheet.CellValue[][];
    getData(highlighted: boolean | undefined, processed: boolean | undefined, delimiter: '', asJson: true): Record<number, jspreadsheet.CellValue>[];
    getData(highlighted: boolean | undefined, processed: boolean | undefined, delimiter: string, asJson?: boolean): string;
    getData(
        highlighted: boolean | undefined,
        processed: boolean | undefined,
        delimiter: undefined | null,
        asJson: true
    ): Record<number, jspreadsheet.CellValue>[];
    getData(
        highlighted?: boolean,
        processed?: boolean,
        delimiter?: string | null,
        asJson?: boolean
    ): jspreadsheet.CellValue[][] | Record<number, jspreadsheet.CellValue>[] | string;
    getData(highlighted?: boolean, processed?: boolean, delimiter?: string | null, asJson?: boolean) {
        return this.data.get(highlighted, processed, delimiter, asJson);
    }

    setData(data?: jspreadsheet.CellValue[][] | Record<string, jspreadsheet.CellValue>[]) {
        return this.data.set(data);
    }

    getValue(cell: string, processedValue?: boolean) {
        return this.data.getValue(cell, processedValue);
    }

    getValueFromCoords(x: number, y: number, processedValue?: boolean) {
        return this.data.getValueFromCoords(x, y, processedValue);
    }

    setValue(cell: Parameters<jspreadsheet.WorksheetInstance['setValue']>[0], value?: jspreadsheet.CellValue, force?: boolean) {
        return this.data.setValue(cell, value, force);
    }

    setValueFromCoords(x: number, y: number, value: jspreadsheet.CellValue, force?: boolean) {
        return this.data.setValueFromCoords(x, y, value, force);
    }

    getWidth(column?: number) {
        return this.columns.getWidth(column);
    }

    setWidth(column: number, width: number): void;
    setWidth(column: number[], width: number | number[]): void;
    setWidth(column: number | number[], width: number | number[]) {
        return this.columns.setWidth(column, width);
    }

    insertRow(mixed?: number | jspreadsheet.CellValue[], rowNumber?: number, insertBefore?: boolean | number) {
        return this.rowManager.insert(mixed, rowNumber, insertBefore);
    }

    moveRow(rowNumber: number, newPositionNumber: number) {
        return this.rowManager.move(rowNumber, newPositionNumber);
    }

    deleteRow(rowNumber?: number, numOfRows?: number) {
        return this.rowManager.delete(rowNumber, numOfRows);
    }

    hideRow(rowNumber: number | number[]) {
        return this.rowManager.hide(rowNumber);
    }

    showRow(rowNumber: number | number[]) {
        return this.rowManager.show(rowNumber);
    }

    getRowData(rowNumber: number, processed?: boolean) {
        return this.rowManager.getData(rowNumber, processed);
    }

    setRowData(rowNumber: number, data: (jspreadsheet.CellValue | null)[], force?: boolean) {
        return this.rowManager.setData(rowNumber, data, force);
    }

    getHeight(row?: undefined): string[];
    getHeight(row: number): string;
    getHeight(row?: number) {
        return this.rowManager.getHeight(row);
    }

    setHeight(row: number, height: number) {
        return this.rowManager.setHeight(row, height);
    }

    getMerge(cellName?: string) {
        return this.merges.get(cellName);
    }

    setMerge(cellName?: string, colspan?: number, rowspan?: number) {
        return this.merges.set(cellName, colspan, rowspan);
    }

    destroyMerge() {
        return this.merges.destroy();
    }

    replaceMerges(merges: Record<string, [number, number]>) {
        return this.merges.replace(merges);
    }

    removeMerge(cellName: string, data?: jspreadsheet.CellValue[]) {
        return this.merges.remove(cellName, data);
    }

    search(query: string) {
        return this.view.search(query);
    }

    resetSearch() {
        return this.view.resetSearch();
    }

    getHeader(column: number) {
        return this.columns.getHeader(column);
    }

    getHeaders(asArray?: boolean) {
        return this.columns.getHeaders(asArray);
    }

    setHeader(column: number, newValue?: string) {
        return this.columns.setHeader(column, newValue);
    }

    getStyle(cell?: string | number[], key?: string) {
        return this.styles.get(cell, key);
    }

    setStyle(o: string, k: string, v: string, force?: boolean, ignoreHistoryAndEvents?: boolean): void;
    setStyle(o: Record<string, string | string[]>, k?: null | undefined, v?: null | undefined, force?: boolean, ignoreHistoryAndEvents?: boolean): void;
    setStyle(
        cell: string | Record<string, string | string[]>,
        property?: string | null,
        value?: string | null,
        forceOverwrite?: boolean,
        ignoreHistoryAndEvents?: boolean
    ) {
        return this.styles.set(cell, property, value, forceOverwrite, ignoreHistoryAndEvents);
    }

    resetStyle(o: Record<string, string | string[]>, ignoreHistoryAndEvents?: boolean) {
        return this.styles.reset(o, ignoreHistoryAndEvents);
    }

    insertColumn(mixed?: number | jspreadsheet.CellValue[], columnNumber?: number, insertBefore?: boolean | number, properties?: jspreadsheet.Column[]) {
        return this.columns.insert(mixed, columnNumber, insertBefore, properties);
    }

    moveColumn(o: number, d: number) {
        return this.columns.move(o, d);
    }

    deleteColumn(columnNumber?: number, numOfColumns?: number) {
        return this.columns.delete(columnNumber, numOfColumns);
    }

    getColumnData(columnNumber: number, processed?: boolean) {
        return this.columns.getData(columnNumber, processed);
    }

    setColumnData(colNumber: number, data: (jspreadsheet.CellValue | null)[], force?: boolean) {
        return this.columns.setData(colNumber, data, force);
    }

    whichPage(row: number) {
        return this.view.whichPage(row);
    }

    page(pageNumber: number) {
        return this.view.page(pageNumber);
    }

    download(includeHeaders?: boolean, processed?: boolean) {
        return this.clipboard.download(includeHeaders, processed);
    }

    getComments(cell?: string | number[]) {
        return this.comments.get(cell);
    }

    setComments(cellId: string, comments: string | null | undefined): void;
    setComments(cellId: Record<string, string | null | undefined>): void;
    setComments(cellId: string | Record<string, string | null | undefined>, comments?: string | null) {
        return this.comments.set(cellId, comments);
    }

    orderBy(column: number, order?: 0 | 1) {
        return this.view.orderBy(column, order);
    }

    undo() {
        return this.history.undo();
    }

    redo() {
        return this.history.redo();
    }

    getCell(cell: string): HTMLTableCellElement;
    getCell(x: number, y: number): HTMLTableCellElement;
    getCell(x: number | string, y?: number) {
        return this.cells.get(x, y);
    }

    getCellFromCoords(x: number, y: number) {
        return this.cells.getFromCoords(x, y);
    }

    getLabel(cell: string): string;
    getLabel(x: number, y: number): string;
    getLabel(x: number | string, y?: number) {
        return this.cells.getLabel(x, y);
    }

    getMeta(cell?: undefined): Record<string, jspreadsheet.MetaInformation> | undefined;
    getMeta(cell: string): jspreadsheet.MetaInformation | null;
    getMeta(cell?: string) {
        return this.meta.get(cell);
    }

    setMeta(o: string, k: string, v: string): void;
    setMeta(o: Record<string, jspreadsheet.MetaInformation>): void;
    setMeta(o: string | Record<string, jspreadsheet.MetaInformation>, k?: string, v?: string) {
        return this.meta.set(o, k, v);
    }

    showColumn(colNumber: number | number[]) {
        return this.columns.show(colNumber);
    }

    hideColumn(colNumber: number | number[]) {
        return this.columns.hide(colNumber);
    }

    showIndex() {
        return this.rowManager.showIndex();
    }

    hideIndex() {
        return this.rowManager.hideIndex();
    }

    openEditor(cell: HTMLTableCellElement, empty?: boolean, e?: KeyboardEvent | MouseEvent | TouchEvent) {
        return this.editor.open(cell, empty, e);
    }

    closeEditor(cell: HTMLTableCellElement, save: boolean) {
        return this.editor.close(cell, save);
    }

    copy(cut?: boolean) {
        if (cut) {
            cutControls();
        } else {
            this.clipboard.copy(true);
        }
    }

    paste(x: number, y: number, data: string) {
        return this.clipboard.paste(x, y, data);
    }

    executeFormula(expression: string, x?: number, y?: number) {
        return this.formulas.execute(expression, x, y);
    }

    getDataFromRange(range: string, processed?: boolean) {
        return this.data.getFromRange(range, processed);
    }

    quantiyOfPages() {
        return this.view.countPages();
    }

    getRange() {
        return this.selection.getRange();
    }

    isSelected(x: number, y: number) {
        return this.selection.isSelected(x, y);
    }

    /**
     * Change the number of frozen rows and frozen columns. Undefined keeps the current number.
     */
    setFreeze(rows?: number, columns?: number) {
        this.rowManager.setFreeze(rows, columns);
    }

    setReadOnly(cell: string | HTMLElement, state: boolean) {
        return this.cells.setReadOnly(cell, state);
    }

    isReadOnly(x: number, y: number): boolean;
    isReadOnly(cellName: string): boolean;
    isReadOnly(x: number | string, y?: number) {
        return this.cells.isReadOnly(x, y);
    }

    getHighlighted() {
        return this.selection.getHighlighted();
    }

    labelCell(x: number, y: number) {
        return this.accessibility.labelCell(x, y);
    }

    down(shiftKey?: boolean, ctrlKey?: boolean) {
        return this.selection.down(shiftKey, ctrlKey);
    }

    first(shiftKey?: boolean, ctrlKey?: boolean) {
        return this.selection.first(shiftKey, ctrlKey);
    }

    last(shiftKey?: boolean, ctrlKey?: boolean) {
        return this.selection.last(shiftKey, ctrlKey);
    }

    left(shiftKey?: boolean, ctrlKey?: boolean) {
        return this.selection.left(shiftKey, ctrlKey);
    }

    right(shiftKey?: boolean, ctrlKey?: boolean) {
        return this.selection.right(shiftKey, ctrlKey);
    }

    up(shiftKey?: boolean, ctrlKey?: boolean) {
        return this.selection.up(shiftKey, ctrlKey);
    }

    openFilter(columnId: number) {
        return this.view.openFilter(columnId);
    }

    resetFilters() {
        return this.view.resetFilters();
    }

    /**
     * Add a worksheet to the spreadsheet, in a new tab. The tabs add it at the end, and position moves it.
     */
    createWorksheet(options: jspreadsheet.WorksheetOptions, position?: number) {
        const spreadsheet = this.parent;

        spreadsheet.creationThroughJss = true;

        spreadsheet.createWorksheetObj(options);

        spreadsheet.element.tabs!.create(options.worksheetName);

        const last = spreadsheet.worksheets.length - 1;
        if (Number.isInteger(position) && position! >= 0 && position! < last) {
            spreadsheet.placeWorksheet(last, position!);
        }
    }

    /**
     * Move a worksheet and its tab to another position.
     */
    moveWorksheet(from: number, to: number) {
        const spreadsheet = this.parent;
        const count = spreadsheet.worksheets.length;

        if (!Number.isInteger(from) || !Number.isInteger(to) || from < 0 || to < 0 || from >= count || to >= count) {
            return false;
        }
        if (from === to) {
            return;
        }

        const worksheet = spreadsheet.worksheets[from];
        spreadsheet.placeWorksheet(from, to);

        worksheet.dispatch('onmoveworksheet', worksheet, from, to);
    }

    /**
     * Rename a worksheet, its tab and the accessible name of its table.
     */
    renameWorksheet(position: number, name: string) {
        const spreadsheet = this.parent;
        const worksheet = spreadsheet.worksheets[position];

        // A name of spaces would give the tab and the table no accessible name
        if (!worksheet || typeof name !== 'string' || !name.trim()) {
            return false;
        }
        if (spreadsheet.worksheets.some((other) => other !== worksheet && other.options.worksheetName === name)) {
            return false;
        }

        const oldName = worksheet.options.worksheetName!;
        if (oldName !== name) {
            worksheet.options.worksheetName = name;
            spreadsheet.element.tabs!.rename(position, name);
            worksheet.accessibility.setGridLabels();

            worksheet.dispatch('onrenameworksheet', worksheet, position, name, oldName);
        }
    }

    /**
     * Open the tab of the worksheet at a position.
     */
    openWorksheet(position: number) {
        const spreadsheet = this.parent;

        spreadsheet.element.tabs!.open(position);
    }

    /**
     * Remove the worksheet at a position, and its tab.
     */
    deleteWorksheet(position: number) {
        this.parent.element.tabs!.remove(position);

        const removedWorksheet = this.parent.worksheets.splice(position, 1)[0];

        // The size observer of the frozen rows holds the worksheet
        removedWorksheet?.rowManager.stopFreeze();

        // getConfig gives the options of the worksheets that the spreadsheet has. A position without a worksheet removes none.
        const index = removedWorksheet ? this.parent.config.worksheets.indexOf(removedWorksheet.options) : -1;
        if (index >= 0) {
            this.parent.config.worksheets.splice(index, 1);
        }

        this.parent.dispatch('ondeleteworksheet', removedWorksheet, position);
    }

    /**
     * The container of the table, which scrolls.
     */
    private createContent() {
        // Create table container
        this.content = document.createElement('div');
        this.content.classList.add('jss_content');
        this.content.onscroll = (e) => {
            this.view.scroll(e);
        };
        this.content.onwheel = (e) => {
            this.view.wheel(e);
        };
    }

    /**
     * The search and the number of rows on a page, above the table. The worksheet adds them only with the search option.
     */
    private createSearch() {
        // Search
        const searchContainer = document.createElement('div');

        const searchLabel = document.createElement('label');
        searchLabel.innerHTML = `${translate('Search')}: `;
        searchContainer.appendChild(searchLabel);

        this.searchInput = document.createElement('input');
        this.searchInput.classList.add('jss_search');
        searchLabel.appendChild(this.searchInput);
        this.searchInput.onfocus = () => {
            this.resetSelection();
        };

        // Pagination select option
        const paginationUpdateContainer = document.createElement('div');

        if (this.options.pagination! > 0 && this.options.paginationOptions && this.options.paginationOptions.length > 0) {
            const paginationDropdown = document.createElement('select');
            this.paginationDropdown = paginationDropdown;
            this.paginationDropdown.classList.add('jss_pagination_dropdown');
            this.paginationDropdown.onchange = () => {
                this.options.pagination = parseInt(paginationDropdown.value);
                this.page(0);
            };

            for (const paginationOption of this.options.paginationOptions) {
                const temp = document.createElement('option');
                temp.value = String(paginationOption);
                temp.innerHTML = String(paginationOption);
                this.paginationDropdown.appendChild(temp);
            }

            // Set initial pagination value
            this.paginationDropdown.value = String(this.options.pagination);

            paginationUpdateContainer.appendChild(document.createTextNode(translate('Show ')));
            paginationUpdateContainer.appendChild(this.paginationDropdown);
            paginationUpdateContainer.appendChild(document.createTextNode(` ${translate('entries')}`));
        }

        // Filter and pagination container
        const filter = document.createElement('div');
        filter.classList.add('jss_filter');
        filter.appendChild(paginationUpdateContainer);
        filter.appendChild(searchContainer);

        return filter;
    }

    /**
     * The columns and the head of the table: the nested headers, the headers and the filters.
     */
    private createHead() {
        // Colsgroup
        this.colgroupContainer = document.createElement('colgroup');
        // A col, then the select-all cell of the header
        let tempCol: HTMLElement = document.createElement('col');
        tempCol.setAttribute('width', '50');
        this.colgroupContainer.appendChild(tempCol);

        // Nested
        if (this.options.nestedHeaders && this.options.nestedHeaders.length > 0 && this.options.nestedHeaders[0]?.[0]) {
            for (const nestedHeader of this.options.nestedHeaders) {
                this.thead.appendChild(this.columns.createNestedHeader(nestedHeader));
            }
        }

        // Row
        this.headerContainer = document.createElement('tr');
        tempCol = document.createElement('td');
        tempCol.classList.add('jss_selectall');
        this.headerContainer.appendChild(tempCol);

        const numberOfColumns = this.columns.count();

        for (let i = 0; i < numberOfColumns; i++) {
            // Create header
            this.columns.createHeader(i);
            // Append cell to the container
            this.headerContainer.appendChild(this.headers[i]);
            this.colgroupContainer.appendChild(this.cols[i].colElement);
        }

        this.thead.appendChild(this.headerContainer);

        // Filters
        if (this.options.filters === true) {
            this.filter = document.createElement('tr');
            const td = document.createElement('td');
            this.filter.appendChild(td);

            for (let i = 0; i < this.options.columns.length; i++) {
                const td = document.createElement('td');
                td.innerHTML = '&nbsp;';
                td.setAttribute('data-x', i);
                td.className = 'jss_column_filter';
                if (this.options.columns[i].type === 'hidden') {
                    td.style.display = 'none';
                }
                this.filter.appendChild(td);
            }

            this.thead.appendChild(this.filter);
        }
    }

    /**
     * The table, with the roles of the grid.
     */
    private createGrid() {
        // Content table
        this.table = document.createElement('table');
        this.table.classList.add('jss_worksheet');
        this.table.setAttribute('cellpadding', '0');
        this.table.setAttribute('cellspacing', '0');
        this.table.setAttribute('unselectable', 'yes');
        //this.table.setAttribute('onselectstart', 'return false');
        this.table.appendChild(this.colgroupContainer);
        this.table.appendChild(this.thead);
        this.table.appendChild(this.tbody);
        this.accessibility.setGridRoles();
        this.selection.setFocusEvents();

        if (!this.options.textOverflow) {
            this.table.classList.add('jss_overflow');
        }
    }

    /**
     * The corner of the selection, which fills the cells.
     */
    private createCorner() {
        // Spreadsheet corner
        this.corner = document.createElement('div');
        this.corner.className = 'jss_corner';
        this.corner.setAttribute('unselectable', 'on');
        // A listener, not an inline handler, which a Content Security Policy without 'unsafe-inline' blocks
        this.corner.addEventListener('selectstart', (event) => event.preventDefault());

        if (this.options.selectionCopy === false) {
            this.corner.style.display = 'none';
        }
    }

    /**
     * The link to Gridwell below the table.
     */
    private createAbout() {
        const ads = document.createElement('a');
        ads.setAttribute('href', 'https://github.com/BitBindle/jspreadsheet-ce');
        this.ads = document.createElement('div');
        this.ads.className = 'jss_about';

        const span = document.createElement('span');
        span.textContent = 'Gridwell';
        ads.appendChild(span);
        this.ads.appendChild(ads);
    }

    /**
     * The pagination below the table, hidden without the pagination option.
     */
    private createPagination() {
        // Pagination
        this.pagination = document.createElement('div');
        this.pagination.classList.add('jss_pagination');
        const paginationInfo = document.createElement('div');
        const paginationPages = document.createElement('div');
        this.pagination.appendChild(paginationInfo);
        this.pagination.appendChild(paginationPages);

        // Hide pagination if not in use
        if (!this.options.pagination) {
            this.pagination.style.display = 'none';
        }
    }

    /**
     * The size and the overflow of the content, and the classes that allow the drag and the resize.
     */
    private setLayout() {
        // Overflow
        if (this.options.tableOverflow === true) {
            if (this.options.tableHeight) {
                (this.content.style as unknown as Record<string, string>)['overflow-y'] = 'auto';
                (this.content.style as unknown as Record<string, string>)['box-shadow'] = 'rgb(221 221 221) 2px 2px 5px 0.1px';
                this.content.style.maxHeight = typeof this.options.tableHeight === 'string' ? this.options.tableHeight : `${this.options.tableHeight}px`;
            }
            if (this.options.tableWidth) {
                (this.content.style as unknown as Record<string, string>)['overflow-x'] = 'auto';
                this.content.style.width = typeof this.options.tableWidth === 'string' ? this.options.tableWidth : `${this.options.tableWidth}px`;
            }
        }

        // With toolbars
        if (this.options.tableOverflow !== true && this.parent.config.toolbar) {
            this.element.classList.add('with-toolbar');
        }

        // Actions
        if (this.options.columnDrag !== false) {
            this.thead.classList.add('draggable');
        }
        if (this.options.columnResize !== false) {
            this.thead.classList.add('resizable');
        }
        if (this.options.rowDrag !== false) {
            this.tbody.classList.add('draggable');
        }
        if (this.options.rowResize !== false) {
            this.tbody.classList.add('resizable');
        }
    }

    /**
     * The data, and the styles, the comments and the classes of the cells.
     */
    private loadContent() {
        // Load data
        this.data.set();

        // Style
        if (this.options.style) {
            // The numbers, indexes in the style list of the spreadsheet, were changed to the styles when the spreadsheet was created
            this.setStyle(this.options.style as Record<string, string>, null, null, true, true);

            delete this.options.style;
        }

        Object.defineProperty(this.options, 'style', {
            enumerable: true,
            configurable: true,
            get: () => {
                return this.getStyle();
            },
        });

        if (this.options.comments) {
            this.comments.load(this.options.comments);
        }

        // Classes
        if (this.options.classes) {
            for (const [cellName, className] of Object.entries(this.options.classes)) {
                const cell = getIdFromColumnName(cellName);
                this.records[cell[1]][cell[0]].element.classList.add(className);
            }
        }
    }

    private createTable() {
        // Elements
        this.thead = document.createElement('thead');
        this.tbody = document.createElement('tbody');

        // Create headers controllers
        this.headers = [];
        this.cols = [];

        this.createContent();
        const filter = this.createSearch();
        this.createHead();
        this.createGrid();
        this.createCorner();
        this.createAbout();
        this.createPagination();

        // Append containers to the table
        if (this.options.search === true) {
            this.element.appendChild(filter);
        }

        // Elements
        this.content.appendChild(this.table);
        this.content.appendChild(this.corner);
        this.content.appendChild(this.clipboard.textarea);
        this.accessibility.setGridLabels();

        this.element.appendChild(this.content);
        this.element.appendChild(this.pagination);
        this.element.appendChild(this.ads);
        this.element.classList.add('jss_container');

        this.element.jssWorksheet = this;
        this.element.jspreadsheet = this;

        this.setLayout();
        this.loadContent();
    }

    /**
     * Prepare the table of the worksheet.
     */
    private prepareTable() {
        // Lazy loading
        if (this.options.lazyLoading === true && this.options.tableOverflow !== true && this.parent.config.fullscreen !== true) {
            console.error('Jspreadsheet: The lazyloading only works when tableOverflow = yes or fullscreen = yes');
            this.options.lazyLoading = false;
        }

        if (!this.options.columns) {
            this.options.columns = [];
        }

        // Number of columns
        let size = this.options.columns.length;
        let keys;

        if (typeof this.options.data?.[0] !== 'undefined') {
            if (!Array.isArray(this.options.data[0])) {
                // Data keys
                keys = Object.keys(this.options.data[0]);

                if (keys.length > size) {
                    size = keys.length;
                }
            } else {
                const numOfColumns = this.options.data[0].length;

                if (numOfColumns > size) {
                    size = numOfColumns;
                }
            }
        }

        // Minimal dimensions
        if (!this.options.minDimensions) {
            this.options.minDimensions = [0, 0];
        }

        if (this.options.minDimensions[0] > size) {
            size = this.options.minDimensions[0];
        }

        // The requests of the sources of the dropdown columns
        const requests: Promise<void>[] = [];

        // Preparations
        for (let i = 0; i < size; i++) {
            // Default column description
            if (!this.options.columns[i]) {
                this.options.columns[i] = {};
            }
            if (!this.options.columns[i].name && keys?.[i]) {
                this.options.columns[i].name = keys[i];
            }

            // Pre-load initial source for json dropdown
            if (this.options.columns[i].type === 'dropdown') {
                // if remote content
                if (this.options.columns[i].url) {
                    const column = this.options.columns[i];
                    requests.push(
                        request(this, column.url!, 'json').then((response) => {
                            if (response) {
                                if (!column.source) {
                                    column.source = [];
                                }

                                for (const option of response.result as NonNullable<typeof column.source>) {
                                    column.source.push(option);
                                }
                            }
                        })
                    );
                }
            }
        }

        // Create the table when is ready
        if (!requests.length) {
            this.createTable();
        } else {
            Promise.all(requests).then(() => {
                this.createTable();
            });
        }
    }

    async build() {
        const el = this.element;

        const spreadsheet = this.parent;

        if (typeof spreadsheet.plugins === 'object') {
            for (const plugin of Object.values(spreadsheet.plugins)) {
                if (typeof plugin.beforeinit === 'function') {
                    plugin.beforeinit(this);
                }
            }
        }

        libraryBase.jspreadsheet.current = this;

        const promises = [];

        // Load the table data based on an CSV file
        if (this.options.csv) {
            promises.push(
                request(this, this.options.csv, 'text').then((response) => {
                    if (response) {
                        // Convert data
                        const newData = parseCSV(response.result as string, this.options.csvDelimiter);

                        // Headers
                        if (this.options.csvHeaders === true && newData.length > 0) {
                            const headers = newData.shift()!;

                            if (headers.length > 0) {
                                if (!this.options.columns) {
                                    this.options.columns = [];
                                }

                                for (let i = 0; i < headers.length; i++) {
                                    if (!this.options.columns[i]) {
                                        this.options.columns[i] = {};
                                    }
                                    // Precedence over pre-configurated titles
                                    if (typeof this.options.columns[i].title === 'undefined') {
                                        this.options.columns[i].title = headers[i];
                                    }
                                }
                            }
                        }
                        // Data
                        this.options.data = newData;
                    }
                    // Without the file, the worksheet is built with its other options
                    this.prepareTable();
                })
            );
        } else if (this.options.url) {
            promises.push(
                request(this, this.options.url, 'json').then((response) => {
                    if (response) {
                        // The data, or an object with the data in its data property. setData changes rows of objects to arrays.
                        this.options.data = (response.result as { data?: jspreadsheet.CellValue[][] }).data || (response.result as jspreadsheet.CellValue[][]);
                    }
                    // Without the data, the worksheet is built with its other options
                    this.prepareTable();
                })
            );
        } else {
            // Prepare table
            this.prepareTable();
        }

        await Promise.all(promises);

        if (typeof spreadsheet.plugins === 'object') {
            for (const plugin of Object.values(spreadsheet.plugins)) {
                if (typeof plugin.init === 'function') {
                    plugin.init(this);
                }
            }
        }
    }
}
