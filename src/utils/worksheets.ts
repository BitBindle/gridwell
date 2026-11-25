import jSuites from 'jsuites';

import libraryBase from './libraryBase.js';

import { parseCSV } from './helpers.js';
import { getData, getDataFromRange, getValue, getValueFromCoords, setData, setValue, setValueFromCoords } from './data.js';
import { cutControls } from '../events/clipboard.js';
import { scrollControls, wheelControls } from '../events/scroll.js';
import { deleteRow, getHeight, getRowData, hideRow, insertRow, moveRow, setHeight, setRowData, showRow } from './rows.js';
import { download } from './download.js';
import { hideIndex, showIndex } from './rows.js';
import { getWorksheetActive } from './activeWorksheet.js';
import { getWorksheetConfig, setConfig } from './config.js';
import dispatch from './dispatch.js';
import { getIdFromColumnName } from './internalHelpers.js';
import { redo, undo } from './history.js';
import { setGridLabels, setGridRoles } from './accessibility.js';
import type jspreadsheet from '../jspreadsheet';
import { Worksheet } from '../Worksheet.js';
import type { Spreadsheet, WorksheetMethod } from '../types/internal';

/**
 * The container of the table, which scrolls.
 */
const createContent = function (this: Worksheet) {
    const obj = this;

    // Create table container
    obj.content = document.createElement('div');
    obj.content.classList.add('jss_content');
    obj.content.onscroll = function (e) {
        scrollControls.call(obj, e);
    };
    obj.content.onwheel = function (e) {
        wheelControls.call(obj, e);
    };
};

/**
 * The search and the number of rows on a page, above the table. The worksheet adds them only with the search option.
 */
const createSearch = function (this: Worksheet) {
    const obj = this;

    // Search
    const searchContainer = document.createElement('div');

    const searchLabel = document.createElement('label');
    searchLabel.innerHTML = jSuites.translate('Search') + ': ';
    searchContainer.appendChild(searchLabel);

    obj.searchInput = document.createElement('input');
    obj.searchInput.classList.add('jss_search');
    searchLabel.appendChild(obj.searchInput);
    obj.searchInput.onfocus = function () {
        obj.resetSelection();
    };

    // Pagination select option
    const paginationUpdateContainer = document.createElement('div');

    if (obj.options.pagination! > 0 && obj.options.paginationOptions && obj.options.paginationOptions.length > 0) {
        obj.paginationDropdown = document.createElement('select');
        obj.paginationDropdown.classList.add('jss_pagination_dropdown');
        obj.paginationDropdown.onchange = function () {
            obj.options.pagination = parseInt((this as HTMLSelectElement).value);
            obj.page(0);
        };

        for (let i = 0; i < obj.options.paginationOptions.length; i++) {
            const temp = document.createElement('option');
            temp.value = String(obj.options.paginationOptions[i]);
            temp.innerHTML = String(obj.options.paginationOptions[i]);
            obj.paginationDropdown.appendChild(temp);
        }

        // Set initial pagination value
        obj.paginationDropdown.value = String(obj.options.pagination);

        paginationUpdateContainer.appendChild(document.createTextNode(jSuites.translate('Show ')));
        paginationUpdateContainer.appendChild(obj.paginationDropdown);
        paginationUpdateContainer.appendChild(document.createTextNode(jSuites.translate('entries')));
    }

    // Filter and pagination container
    const filter = document.createElement('div');
    filter.classList.add('jss_filter');
    filter.appendChild(paginationUpdateContainer);
    filter.appendChild(searchContainer);

    return filter;
};

/**
 * The columns and the head of the table: the nested headers, the headers and the filters.
 */
const createHead = function (this: Worksheet) {
    const obj = this;

    // Colsgroup
    obj.colgroupContainer = document.createElement('colgroup');
    // A col, then the select-all cell of the header
    let tempCol: HTMLElement = document.createElement('col');
    tempCol.setAttribute('width', '50');
    obj.colgroupContainer.appendChild(tempCol);

    // Nested
    if (obj.options.nestedHeaders && obj.options.nestedHeaders.length > 0 && obj.options.nestedHeaders[0] && obj.options.nestedHeaders[0][0]) {
        for (let j = 0; j < obj.options.nestedHeaders.length; j++) {
            obj.thead.appendChild(obj.columns.createNestedHeader(obj.options.nestedHeaders[j]));
        }
    }

    // Row
    obj.headerContainer = document.createElement('tr');
    tempCol = document.createElement('td');
    tempCol.classList.add('jss_selectall');
    obj.headerContainer.appendChild(tempCol);

    const numberOfColumns = obj.columns.count();

    for (let i = 0; i < numberOfColumns; i++) {
        // Create header
        obj.columns.createHeader(i);
        // Append cell to the container
        obj.headerContainer.appendChild(obj.headers[i]);
        obj.colgroupContainer.appendChild(obj.cols[i].colElement);
    }

    obj.thead.appendChild(obj.headerContainer);

    // Filters
    if (obj.options.filters == true) {
        obj.filter = document.createElement('tr');
        const td = document.createElement('td');
        obj.filter.appendChild(td);

        for (let i = 0; i < obj.options.columns.length; i++) {
            const td = document.createElement('td');
            td.innerHTML = '&nbsp;';
            td.setAttribute('data-x', i);
            td.className = 'jss_column_filter';
            if (obj.options.columns[i].type == 'hidden') {
                td.style.display = 'none';
            }
            obj.filter.appendChild(td);
        }

        obj.thead.appendChild(obj.filter);
    }
};

/**
 * The table, with the roles of the grid.
 */
const createGrid = function (this: Worksheet) {
    const obj = this;

    // Content table
    obj.table = document.createElement('table');
    obj.table.classList.add('jss_worksheet');
    obj.table.setAttribute('cellpadding', '0');
    obj.table.setAttribute('cellspacing', '0');
    obj.table.setAttribute('unselectable', 'yes');
    //obj.table.setAttribute('onselectstart', 'return false');
    obj.table.appendChild(obj.colgroupContainer);
    obj.table.appendChild(obj.thead);
    obj.table.appendChild(obj.tbody);
    setGridRoles.call(obj);
    obj.selection.setFocusEvents();

    if (!obj.options.textOverflow) {
        obj.table.classList.add('jss_overflow');
    }
};

/**
 * The corner of the selection, which fills the cells.
 */
const createCorner = function (this: Worksheet) {
    const obj = this;

    // Spreadsheet corner
    obj.corner = document.createElement('div');
    obj.corner.className = 'jss_corner';
    obj.corner.setAttribute('unselectable', 'on');
    obj.corner.setAttribute('onselectstart', 'return false');

    if (obj.options.selectionCopy == false) {
        obj.corner.style.display = 'none';
    }
};

/**
 * The link to Jspreadsheet below the table.
 */
const createAbout = function (this: Worksheet) {
    const obj = this;

    // Powered by Jspreadsheet
    const ads = document.createElement('a');
    ads.setAttribute('href', 'https://bossanova.uk/jspreadsheet/');
    obj.ads = document.createElement('div');
    obj.ads.className = 'jss_about';

    const span = document.createElement('span');
    span.innerHTML = 'Jspreadsheet CE';
    ads.appendChild(span);
    obj.ads.appendChild(ads);
};

/**
 * The pagination below the table, hidden without the pagination option.
 */
const createPagination = function (this: Worksheet) {
    const obj = this;

    // Pagination
    obj.pagination = document.createElement('div');
    obj.pagination.classList.add('jss_pagination');
    const paginationInfo = document.createElement('div');
    const paginationPages = document.createElement('div');
    obj.pagination.appendChild(paginationInfo);
    obj.pagination.appendChild(paginationPages);

    // Hide pagination if not in use
    if (!obj.options.pagination) {
        obj.pagination.style.display = 'none';
    }
};

/**
 * The size and the overflow of the content, and the classes that allow the drag and the resize.
 */
const setLayout = function (this: Worksheet) {
    const obj = this;

    // Overflow
    if (obj.options.tableOverflow == true) {
        if (obj.options.tableHeight) {
            (obj.content.style as unknown as Record<string, string>)['overflow-y'] = 'auto';
            (obj.content.style as unknown as Record<string, string>)['box-shadow'] = 'rgb(221 221 221) 2px 2px 5px 0.1px';
            obj.content.style.maxHeight = typeof obj.options.tableHeight === 'string' ? obj.options.tableHeight : obj.options.tableHeight + 'px';
        }
        if (obj.options.tableWidth) {
            (obj.content.style as unknown as Record<string, string>)['overflow-x'] = 'auto';
            obj.content.style.width = typeof obj.options.tableWidth === 'string' ? obj.options.tableWidth : obj.options.tableWidth + 'px';
        }
    }

    // With toolbars
    if (obj.options.tableOverflow != true && obj.parent.config.toolbar) {
        obj.element.classList.add('with-toolbar');
    }

    // Actions
    if (obj.options.columnDrag != false) {
        obj.thead.classList.add('draggable');
    }
    if (obj.options.columnResize != false) {
        obj.thead.classList.add('resizable');
    }
    if (obj.options.rowDrag != false) {
        obj.tbody.classList.add('draggable');
    }
    if (obj.options.rowResize != false) {
        obj.tbody.classList.add('resizable');
    }
};

/**
 * The data, and the styles, the comments and the classes of the cells.
 */
const loadContent = function (this: Worksheet) {
    const obj = this;

    // Load data
    obj.setData.call(obj);

    // Style
    if (obj.options.style) {
        // The numbers, indexes in the style list of the spreadsheet, were changed to the styles when the spreadsheet was created
        obj.setStyle(obj.options.style as Record<string, string>, null, null, true, true);

        delete obj.options.style;
    }

    Object.defineProperty(obj.options, 'style', {
        enumerable: true,
        configurable: true,
        get() {
            return obj.getStyle();
        },
    });

    if (obj.options.comments) {
        obj.setComments(obj.options.comments);
    }

    // Classes
    if (obj.options.classes) {
        const k = Object.keys(obj.options.classes);
        for (let i = 0; i < k.length; i++) {
            const cell = getIdFromColumnName(k[i], true);
            obj.records[cell[1]][cell[0]].element.classList.add(obj.options.classes[k[i]]);
        }
    }
};

const createTable = function (this: Worksheet) {
    const obj = this;

    // Elements
    obj.thead = document.createElement('thead');
    obj.tbody = document.createElement('tbody');

    // Create headers controllers
    obj.headers = [];
    obj.cols = [];

    createContent.call(obj);
    const filter = createSearch.call(obj);
    createHead.call(obj);
    createGrid.call(obj);
    createCorner.call(obj);
    createAbout.call(obj);
    createPagination.call(obj);

    // Append containers to the table
    if (obj.options.search == true) {
        obj.element.appendChild(filter);
    }

    // Elements
    obj.content.appendChild(obj.table);
    obj.content.appendChild(obj.corner);
    obj.content.appendChild(obj.clipboard.textarea);
    setGridLabels.call(obj);

    obj.element.appendChild(obj.content);
    obj.element.appendChild(obj.pagination);
    obj.element.appendChild(obj.ads);
    obj.element.classList.add('jss_container');

    obj.element.jssWorksheet = obj;
    obj.element.jspreadsheet = obj;

    setLayout.call(obj);
    loadContent.call(obj);
};

/**
 * Prepare the jspreadsheet table
 *
 * @Param config
 */
const prepareTable = function (this: Worksheet) {
    const obj = this;

    // Lazy loading
    if (obj.options.lazyLoading == true && obj.options.tableOverflow != true && obj.parent.config.fullscreen != true) {
        console.error('Jspreadsheet: The lazyloading only works when tableOverflow = yes or fullscreen = yes');
        obj.options.lazyLoading = false;
    }

    if (!obj.options.columns) {
        obj.options.columns = [];
    }

    // Number of columns
    let size = obj.options.columns.length;
    let keys;

    if (obj.options.data && typeof obj.options.data[0] !== 'undefined') {
        if (!Array.isArray(obj.options.data[0])) {
            // Data keys
            keys = Object.keys(obj.options.data[0]);

            if (keys.length > size) {
                size = keys.length;
            }
        } else {
            const numOfColumns = obj.options.data[0].length;

            if (numOfColumns > size) {
                size = numOfColumns;
            }
        }
    }

    // Minimal dimensions
    if (!obj.options.minDimensions) {
        obj.options.minDimensions = [0, 0];
    }

    if (obj.options.minDimensions[0] > size) {
        size = obj.options.minDimensions[0];
    }

    // Requests
    const multiple = [];

    // Preparations
    for (let i = 0; i < size; i++) {
        // Default column description
        if (!obj.options.columns[i]) {
            obj.options.columns[i] = {};
        }
        if (!obj.options.columns[i].name && keys && keys[i]) {
            obj.options.columns[i].name = keys[i];
        }

        // Pre-load initial source for json dropdown
        if (obj.options.columns[i].type == 'dropdown') {
            // if remote content
            if (obj.options.columns[i].url) {
                multiple.push({
                    url: obj.options.columns[i].url,
                    index: i,
                    method: 'GET',
                    dataType: 'json',
                    success: function (this: { index: number }, data: unknown[]) {
                        if (!obj.options.columns[this.index].source) {
                            obj.options.columns[this.index].source = [];
                        }

                        for (let i = 0; i < data.length; i++) {
                            obj.options.columns[this.index].source.push(data[i]);
                        }
                    },
                });
            }
        }
    }

    // Create the table when is ready
    if (!multiple.length) {
        createTable.call(obj);
    } else {
        jSuites.ajax(multiple, function () {
            createTable.call(obj);
        });
    }
};

export const getNextDefaultWorksheetName = function (spreadsheet: Spreadsheet) {
    const defaultWorksheetNameRegex = /^Sheet(\d+)$/;

    let largestWorksheetNumber = 0;

    spreadsheet.worksheets.forEach(function (worksheet) {
        const regexResult = defaultWorksheetNameRegex.exec(worksheet.options.worksheetName!);
        if (regexResult) {
            largestWorksheetNumber = Math.max(largestWorksheetNumber, parseInt(regexResult[1]));
        }
    });

    return 'Sheet' + (largestWorksheetNumber + 1);
};

export const buildWorksheet = async function (this: Worksheet) {
    const obj = this;
    const el = obj.element;

    const spreadsheet = obj.parent;

    if (typeof spreadsheet.plugins === 'object') {
        Object.entries(spreadsheet.plugins).forEach(function ([, plugin]) {
            if (typeof plugin.beforeinit === 'function') {
                plugin.beforeinit(obj);
            }
        });
    }

    libraryBase.jspreadsheet.current = obj;

    const promises = [];

    // Load the table data based on an CSV file
    if (obj.options.csv) {
        const promise = new Promise<void>((resolve) => {
            // Load CSV file
            jSuites.ajax({
                url: obj.options.csv,
                method: 'GET',
                dataType: 'text',
                success: function (result: string) {
                    // Convert data
                    const newData = parseCSV(result, obj.options.csvDelimiter);

                    // Headers
                    if (obj.options.csvHeaders == true && newData.length > 0) {
                        const headers = newData.shift()!;

                        if (headers.length > 0) {
                            if (!obj.options.columns) {
                                obj.options.columns = [];
                            }

                            for (let i = 0; i < headers.length; i++) {
                                if (!obj.options.columns[i]) {
                                    obj.options.columns[i] = {};
                                }
                                // Precedence over pre-configurated titles
                                if (typeof obj.options.columns[i].title === 'undefined') {
                                    obj.options.columns[i].title = headers[i];
                                }
                            }
                        }
                    }
                    // Data
                    obj.options.data = newData;
                    // Prepare table
                    prepareTable.call(obj);

                    resolve();
                },
            });
        });

        promises.push(promise);
    } else if (obj.options.url) {
        const promise = new Promise<void>((resolve) => {
            jSuites.ajax({
                url: obj.options.url,
                method: 'GET',
                dataType: 'json',
                // The data, or an object with the data in its data property. setData changes rows of objects to arrays.
                success: function (result: { data?: jspreadsheet.CellValue[][] }) {
                    // Data
                    obj.options.data = result.data ? result.data : (result as jspreadsheet.CellValue[][]);
                    // Prepare table
                    prepareTable.call(obj);

                    resolve();
                },
            });
        });

        promises.push(promise);
    } else {
        // Prepare table
        prepareTable.call(obj);
    }

    await Promise.all(promises);

    if (typeof spreadsheet.plugins === 'object') {
        Object.entries(spreadsheet.plugins).forEach(function ([, plugin]) {
            if (typeof plugin.init === 'function') {
                plugin.init(obj);
            }
        });
    }
};

export const createWorksheetObj = function (this: Worksheet, options: jspreadsheet.WorksheetOptions) {
    const obj = this;

    const spreadsheet = obj.parent;

    if (!options.worksheetName) {
        options.worksheetName = getNextDefaultWorksheetName(obj.parent);
    }

    const newWorksheet = new Worksheet(spreadsheet, options);

    spreadsheet.config.worksheets.push(newWorksheet.options);
    spreadsheet.worksheets.push(newWorksheet);

    return newWorksheet;
};

export const createWorksheet: WorksheetMethod<'createWorksheet'> = function (options) {
    const obj = this;
    const spreadsheet = obj.parent;

    spreadsheet.creationThroughJss = true;

    createWorksheetObj.call(obj, options);

    spreadsheet.element.tabs!.create(options.worksheetName);
};

export const openWorksheet: WorksheetMethod<'openWorksheet'> = function (position) {
    const obj = this;
    const spreadsheet = obj.parent;

    spreadsheet.element.tabs!.open(position);
};

export const deleteWorksheet: WorksheetMethod<'deleteWorksheet'> = function (position) {
    const obj = this;

    obj.parent.element.tabs!.remove(position);

    const removedWorksheet = obj.parent.worksheets.splice(position, 1)[0];

    dispatch.call(obj.parent, 'ondeleteworksheet', removedWorksheet, position);
};

// never[]: each method has its own parameters
const worksheetPublicMethods: [keyof jspreadsheet.WorksheetInstance, (this: Worksheet, ...args: never[]) => unknown][] = [
    [
        'selectAll',
        function (this: Worksheet) {
            return this.selection.selectAll();
        },
    ],
    [
        'updateSelectionFromCoords',
        function (this: Worksheet, x1: number, y1: number, x2: number, y2: number) {
            return this.selection.updateFromCoords(x1, y1, x2, y2);
        },
    ],
    [
        'resetSelection',
        function (this: Worksheet) {
            return this.selection.reset();
        },
    ],
    [
        'getSelection',
        function (this: Worksheet) {
            return this.selection.get();
        },
    ],
    [
        'getSelected',
        function (this: Worksheet, columnNameOnly?: boolean) {
            return this.selection.getSelected(columnNameOnly);
        },
    ],
    [
        'getSelectedColumns',
        function (this: Worksheet, visibleOnly?: boolean) {
            return this.selection.getSelectedColumns(visibleOnly);
        },
    ],
    [
        'getSelectedRows',
        function (this: Worksheet, visibleOnly?: boolean) {
            return this.selection.getSelectedRows(visibleOnly);
        },
    ],
    ['getData', getData],
    ['setData', setData],
    ['getValue', getValue],
    ['getValueFromCoords', getValueFromCoords],
    ['setValue', setValue],
    ['setValueFromCoords', setValueFromCoords],
    [
        'getWidth',
        function (this: Worksheet, column?: number) {
            return this.columns.getWidth(column);
        },
    ],
    [
        'setWidth',
        function (this: Worksheet, column: number | number[], width: number | number[]) {
            return this.columns.setWidth(column, width);
        },
    ],
    ['insertRow', insertRow],
    [
        'moveRow',
        function (this: Worksheet, rowNumber: number, newPositionNumber: number) {
            return moveRow.call(this, rowNumber, newPositionNumber);
        },
    ],
    ['deleteRow', deleteRow],
    ['hideRow', hideRow],
    ['showRow', showRow],
    ['getRowData', getRowData],
    ['setRowData', setRowData],
    ['getHeight', getHeight],
    [
        'setHeight',
        function (this: Worksheet, row: number, height: number) {
            return setHeight.call(this, row, height);
        },
    ],
    [
        'getMerge',
        function (this: Worksheet, cellName?: string) {
            return this.merges.get(cellName);
        },
    ],
    [
        'setMerge',
        function (this: Worksheet, cellName: string, colspan: number, rowspan: number) {
            return this.merges.set(cellName, colspan, rowspan);
        },
    ],
    [
        'destroyMerge',
        function (this: Worksheet) {
            return this.merges.destroy();
        },
    ],
    [
        'removeMerge',
        function (this: Worksheet, cellName: string, data?: jspreadsheet.CellValue[]) {
            return this.merges.remove(cellName, data);
        },
    ],
    [
        'search',
        function (this: Worksheet, query: string) {
            return this.view.search(query);
        },
    ],
    [
        'resetSearch',
        function (this: Worksheet) {
            return this.view.resetSearch();
        },
    ],
    [
        'getHeader',
        function (this: Worksheet, column: number) {
            return this.columns.getHeader(column);
        },
    ],
    [
        'getHeaders',
        function (this: Worksheet, asArray?: boolean) {
            return this.columns.getHeaders(asArray);
        },
    ],
    [
        'setHeader',
        function (this: Worksheet, column: number, newValue?: string) {
            return this.columns.setHeader(column, newValue);
        },
    ],
    [
        'getStyle',
        function (this: Worksheet, cell?: string | number[], key?: string) {
            return this.styles.get(cell, key);
        },
    ],
    [
        'setStyle',
        function (
            this: Worksheet,
            cell: string | Record<string, string | string[]>,
            property: string | null | undefined,
            value: string | null | undefined,
            forceOverwrite: boolean,
            ignoreHistoryAndEvents?: boolean
        ) {
            return this.styles.set(cell, property, value, forceOverwrite, ignoreHistoryAndEvents);
        },
    ],
    [
        'resetStyle',
        function (this: Worksheet, o: Record<string, string | string[]>, ignoreHistoryAndEvents?: boolean) {
            return this.styles.reset(o, ignoreHistoryAndEvents);
        },
    ],
    [
        'insertColumn',
        function (
            this: Worksheet,
            mixed?: number | jspreadsheet.CellValue[],
            columnNumber?: number,
            insertBefore?: boolean | number,
            properties?: jspreadsheet.Column[]
        ) {
            return this.columns.insert(mixed, columnNumber, insertBefore, properties);
        },
    ],
    [
        'moveColumn',
        function (this: Worksheet, o: number, d: number) {
            return this.columns.move(o, d);
        },
    ],
    [
        'deleteColumn',
        function (this: Worksheet, columnNumber?: number, numOfColumns?: number) {
            return this.columns.delete(columnNumber, numOfColumns);
        },
    ],
    [
        'getColumnData',
        function (this: Worksheet, columnNumber: number, processed?: boolean) {
            return this.columns.getData(columnNumber, processed);
        },
    ],
    [
        'setColumnData',
        function (this: Worksheet, colNumber: number, data: (jspreadsheet.CellValue | null)[], force?: boolean) {
            return this.columns.setData(colNumber, data, force);
        },
    ],
    [
        'whichPage',
        function (this: Worksheet, row: number) {
            return this.view.whichPage(row);
        },
    ],
    [
        'page',
        function (this: Worksheet, pageNumber: number) {
            return this.view.page(pageNumber);
        },
    ],
    ['download', download],
    [
        'getComments',
        function (this: Worksheet, cell?: string | number[]) {
            return this.comments.get(cell);
        },
    ],
    [
        'setComments',
        function (this: Worksheet, cellId: string | Record<string, string | null | undefined>, comments?: string | null) {
            return this.comments.set(cellId, comments);
        },
    ],
    [
        'orderBy',
        function (this: Worksheet, column: number, order?: 0 | 1) {
            return this.view.orderBy(column, order);
        },
    ],
    ['undo', undo],
    ['redo', redo],
    [
        'getCell',
        function (this: Worksheet, x: number, y: number) {
            return this.cells.get(x, y);
        },
    ],
    [
        'getCellFromCoords',
        function (this: Worksheet, x: number, y: number) {
            return this.cells.getFromCoords(x, y);
        },
    ],
    [
        'getLabel',
        function (this: Worksheet, x: number, y: number) {
            return this.cells.getLabel(x, y);
        },
    ],
    ['getConfig', getWorksheetConfig],
    ['setConfig', setConfig],
    [
        'getMeta',
        function (this: Worksheet, cell: string) {
            return this.meta.get(cell);
        },
    ],
    [
        'setMeta',
        function (this: Worksheet, o: string | Record<string, jspreadsheet.MetaInformation>, k?: string, v?: string) {
            return this.meta.set(o, k, v);
        },
    ],
    [
        'showColumn',
        function (this: Worksheet, colNumber: number | number[]) {
            return this.columns.show(colNumber);
        },
    ],
    [
        'hideColumn',
        function (this: Worksheet, colNumber: number | number[]) {
            return this.columns.hide(colNumber);
        },
    ],
    ['showIndex', showIndex],
    ['hideIndex', hideIndex],
    ['getWorksheetActive', getWorksheetActive],
    [
        'openEditor',
        function (this: Worksheet, cell: HTMLTableCellElement, empty?: boolean, e?: KeyboardEvent | MouseEvent | TouchEvent) {
            return this.editor.open(cell, empty, e);
        },
    ],
    [
        'closeEditor',
        function (this: Worksheet, cell: HTMLTableCellElement, save: boolean) {
            return this.editor.close(cell, save);
        },
    ],
    ['createWorksheet', createWorksheet],
    ['openWorksheet', openWorksheet],
    ['deleteWorksheet', deleteWorksheet],
    [
        'copy',
        function (this: Worksheet, cut: boolean) {
            if (cut) {
                cutControls();
            } else {
                this.clipboard.copy(true);
            }
        },
    ],
    [
        'paste',
        function (this: Worksheet, x: number, y: number, data: string) {
            return this.clipboard.paste(x, y, data);
        },
    ],
    [
        'executeFormula',
        function (this: Worksheet, expression: string, x?: number, y?: number) {
            return this.formulas.execute(expression, x, y);
        },
    ],
    ['getDataFromRange', getDataFromRange],
    [
        'quantiyOfPages',
        function (this: Worksheet) {
            return this.view.countPages();
        },
    ],
    [
        'getRange',
        function (this: Worksheet) {
            return this.selection.getRange();
        },
    ],
    [
        'isSelected',
        function (this: Worksheet, x: number, y: number) {
            return this.selection.isSelected(x, y);
        },
    ],
    [
        'setReadOnly',
        function (this: Worksheet, cell: string | HTMLElement, state: boolean) {
            return this.cells.setReadOnly(cell, state);
        },
    ],
    [
        'isReadOnly',
        function (this: Worksheet, x: number | string, y?: number) {
            return this.cells.isReadOnly(x, y);
        },
    ],
    [
        'getHighlighted',
        function (this: Worksheet) {
            return this.selection.getHighlighted();
        },
    ],
    // The public dispatch takes any event name, and dispatch is typed for the events of the library
    ['dispatch', dispatch as (this: Worksheet, ...args: never[]) => unknown],
    [
        'down',
        function (this: Worksheet, shiftKey?: boolean, ctrlKey?: boolean) {
            return this.selection.down(shiftKey, ctrlKey);
        },
    ],
    [
        'first',
        function (this: Worksheet, shiftKey?: boolean, ctrlKey?: boolean) {
            return this.selection.first(shiftKey, ctrlKey);
        },
    ],
    [
        'last',
        function (this: Worksheet, shiftKey?: boolean, ctrlKey?: boolean) {
            return this.selection.last(shiftKey, ctrlKey);
        },
    ],
    [
        'left',
        function (this: Worksheet, shiftKey?: boolean, ctrlKey?: boolean) {
            return this.selection.left(shiftKey, ctrlKey);
        },
    ],
    [
        'right',
        function (this: Worksheet, shiftKey?: boolean, ctrlKey?: boolean) {
            return this.selection.right(shiftKey, ctrlKey);
        },
    ],
    [
        'up',
        function (this: Worksheet, shiftKey?: boolean, ctrlKey?: boolean) {
            return this.selection.up(shiftKey, ctrlKey);
        },
    ],
    [
        'openFilter',
        function (this: Worksheet, columnId: number) {
            return this.view.openFilter(columnId);
        },
    ],
    [
        'resetFilters',
        function (this: Worksheet) {
            return this.view.resetFilters();
        },
    ],
];

// The public methods are on the prototype, as the methods of a class: not enumerable, and an instance can replace one
for (const [methodName, method] of worksheetPublicMethods) {
    Object.defineProperty(Worksheet.prototype, methodName, { value: method, writable: true, configurable: true });
}
