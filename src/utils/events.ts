import jSuites from 'jsuites';

import { closeEditor, openEditor, setCheckRadioValue } from './editor.js';
import libraryBase from './libraryBase.js';
import { down, first, last, left, right, up } from './keys.js';
import { isColMerged, isRowMerged } from './merges.js';
import { getCellColumnIndex } from './internal.js';
import { copyData, removeCopySelection, resetSelection, selectAll, updateCornerPosition, updateSelectionFromCoords } from './selection.js';
import { copy, paste } from './copyPaste.js';
import { openFilter } from './filter.js';
import { loadDown, loadUp } from './lazyLoading.js';
import { setWidth } from './columns.js';
import { moveRow, setHeight } from './rows.js';
import version from './version.js';
import { getCellNameFromCoords } from './helpers.js';
import { getIndexWidth } from './freeze.js';
import type { TargetEvent, Worksheet } from '../types/internal';

// The element of the worksheet that contains the element, which keeps the worksheet, and the section: 1 for the header, 2 for the body
/**
 * The index in the data-x or data-y attribute of an element, or null when the element has none
 */
const getIndex = function (element: Element, name: 'data-x' | 'data-y') {
    const text = element.getAttribute(name);
    return text ? parseInt(text) : null;
};

const getElement = function (element: HTMLElement): [Worksheet['element'] | 0 | null, number] {
    let jssSection = 0;
    let jssElement: Worksheet['element'] | 0 | null = 0;

    // The path goes up to the document, which has no className
    function path(element: HTMLElement) {
        if (element.className) {
            if (element.classList.contains('jss_container')) {
                jssElement = element as Worksheet['element'];
            }

            if (element.classList.contains('jss_spreadsheet')) {
                jssElement = element.querySelector(':scope > .jtabs-content > .jtabs-selected') as Worksheet['element'] | null;
            }
        }

        if (element.tagName == 'THEAD') {
            jssSection = 1;
        } else if (element.tagName == 'TBODY') {
            jssSection = 2;
        }

        if (element.parentNode) {
            if (!jssElement) {
                path(element.parentNode as HTMLElement);
            }
        }
    }

    path(element);

    // A section of a table outside a worksheet is not a section of the grid
    return [jssElement, jssElement ? jssSection : 0];
};

const mouseUpControls = function (e: TargetEvent<MouseEvent>) {
    if (libraryBase.jspreadsheet.current) {
        // Update cell size
        if (libraryBase.jspreadsheet.current.resizing) {
            // Columns to be updated
            if (libraryBase.jspreadsheet.current.resizing.column != null) {
                // New width
                const newWidth = parseInt(
                    libraryBase.jspreadsheet.current.cols[libraryBase.jspreadsheet.current.resizing.column].colElement.getAttribute('width')
                );
                // Columns
                const columns = libraryBase.jspreadsheet.current.getSelectedColumns();
                if (columns.length > 1) {
                    const currentWidth = [];
                    for (let i = 0; i < columns.length; i++) {
                        currentWidth.push(parseInt(libraryBase.jspreadsheet.current.cols[columns[i]].colElement.getAttribute('width')));
                    }
                    // Previous width
                    const index = columns.indexOf(parseInt(libraryBase.jspreadsheet.current.resizing.column));
                    currentWidth[index] = libraryBase.jspreadsheet.current.resizing.width;
                    setWidth.call(libraryBase.jspreadsheet.current, columns, newWidth, currentWidth);
                } else {
                    setWidth.call(
                        libraryBase.jspreadsheet.current,
                        parseInt(libraryBase.jspreadsheet.current.resizing.column),
                        newWidth,
                        libraryBase.jspreadsheet.current.resizing.width
                    );
                }
                // Remove border
                libraryBase.jspreadsheet.current.headers[libraryBase.jspreadsheet.current.resizing.column].classList.remove('resizing');
                for (let j = 0; j < libraryBase.jspreadsheet.current.records.length; j++) {
                    if (libraryBase.jspreadsheet.current.records[j][libraryBase.jspreadsheet.current.resizing.column]) {
                        libraryBase.jspreadsheet.current.records[j][libraryBase.jspreadsheet.current.resizing.column].element.classList.remove('resizing');
                    }
                }
            } else {
                // Remove Class
                libraryBase.jspreadsheet.current.rows[libraryBase.jspreadsheet.current.resizing.row].element.children[0].classList.remove('resizing');
                // setHeight parses the height, and it ignores a missing attribute (NaN) as it ignored null
                const newHeight = parseInt(libraryBase.jspreadsheet.current.rows[libraryBase.jspreadsheet.current.resizing.row].element.getAttribute('height'));
                setHeight.call(
                    libraryBase.jspreadsheet.current,
                    libraryBase.jspreadsheet.current.resizing.row,
                    newHeight,
                    libraryBase.jspreadsheet.current.resizing.height
                );
                // Remove border
                libraryBase.jspreadsheet.current.resizing.element.classList.remove('resizing');
            }
            // Reset resizing helper
            libraryBase.jspreadsheet.current.resizing = null;
        } else if (libraryBase.jspreadsheet.current.dragging) {
            // Reset dragging helper
            if (libraryBase.jspreadsheet.current.dragging) {
                if (libraryBase.jspreadsheet.current.dragging.column != null) {
                    // Target
                    const columnId = getIndex(e.target, 'data-x');
                    // Remove move style
                    libraryBase.jspreadsheet.current.headers[libraryBase.jspreadsheet.current.dragging.column].classList.remove('dragging');
                    for (let j = 0; j < libraryBase.jspreadsheet.current.rows.length; j++) {
                        if (libraryBase.jspreadsheet.current.records[j][libraryBase.jspreadsheet.current.dragging.column]) {
                            libraryBase.jspreadsheet.current.records[j][libraryBase.jspreadsheet.current.dragging.column].element.classList.remove('dragging');
                        }
                    }
                    for (let i = 0; i < libraryBase.jspreadsheet.current.headers.length; i++) {
                        libraryBase.jspreadsheet.current.headers[i].classList.remove('dragging-left');
                        libraryBase.jspreadsheet.current.headers[i].classList.remove('dragging-right');
                    }
                    // Update position
                    if (columnId != null) {
                        if (libraryBase.jspreadsheet.current.dragging.column != libraryBase.jspreadsheet.current.dragging.destination) {
                            libraryBase.jspreadsheet.current.moveColumn(
                                libraryBase.jspreadsheet.current.dragging.column,
                                libraryBase.jspreadsheet.current.dragging.destination
                            );
                        }
                    }
                } else {
                    // A row that moved has a sibling. A row that did not move can be the only row.
                    if (libraryBase.jspreadsheet.current.dragging.row != libraryBase.jspreadsheet.current.dragging.destination) {
                        let position;

                        if (libraryBase.jspreadsheet.current.dragging.element.nextSibling) {
                            position = parseInt((libraryBase.jspreadsheet.current.dragging.element.nextSibling as HTMLElement).getAttribute('data-y'));
                            if (libraryBase.jspreadsheet.current.dragging.row < position) {
                                position -= 1;
                            }
                        } else {
                            position = parseInt((libraryBase.jspreadsheet.current.dragging.element.previousSibling as HTMLElement).getAttribute('data-y'));
                        }

                        moveRow.call(libraryBase.jspreadsheet.current, libraryBase.jspreadsheet.current.dragging.row, position, true);
                    }
                    libraryBase.jspreadsheet.current.dragging.element.classList.remove('dragging');
                }
                libraryBase.jspreadsheet.current.dragging = null;
            }
        } else {
            // Close any corner selection
            if (libraryBase.jspreadsheet.current.selectedCorner) {
                libraryBase.jspreadsheet.current.selectedCorner = false;

                // Data to be copied
                if (libraryBase.jspreadsheet.current.selection.length > 0) {
                    // Copy data
                    copyData.call(
                        libraryBase.jspreadsheet.current,
                        libraryBase.jspreadsheet.current.selection[0],
                        libraryBase.jspreadsheet.current.selection[libraryBase.jspreadsheet.current.selection.length - 1]
                    );

                    // Remove selection
                    removeCopySelection.call(libraryBase.jspreadsheet.current);
                }
            }
        }
    }

    // Clear any time control
    if (libraryBase.jspreadsheet.timeControl) {
        clearTimeout(libraryBase.jspreadsheet.timeControl);
        libraryBase.jspreadsheet.timeControl = null;
    }

    // Mouse up
    libraryBase.jspreadsheet.isMouseAction = false;
};

const mouseDownControls = function (e: TargetEvent<MouseEvent>) {
    e = e || window.event;

    let mouseButton;

    if (e.buttons) {
        mouseButton = e.buttons;
    } else if (e.button) {
        mouseButton = e.button;
    } else {
        mouseButton = e.which;
    }

    // Get elements
    const jssTable = getElement(e.target);

    if (jssTable[0]) {
        if (libraryBase.jspreadsheet.current != jssTable[0].jssWorksheet) {
            if (libraryBase.jspreadsheet.current) {
                if (libraryBase.jspreadsheet.current.edition) {
                    closeEditor.call(libraryBase.jspreadsheet.current, libraryBase.jspreadsheet.current.edition[0], true);
                }
                libraryBase.jspreadsheet.current.resetSelection();
            }
            libraryBase.jspreadsheet.current = jssTable[0].jssWorksheet;
        }
    } else {
        if (libraryBase.jspreadsheet.current) {
            if (libraryBase.jspreadsheet.current.edition) {
                closeEditor.call(libraryBase.jspreadsheet.current, libraryBase.jspreadsheet.current.edition[0], true);
            }

            if (!e.target.closest('.jss_object')) {
                resetSelection.call(libraryBase.jspreadsheet.current, true);
                libraryBase.jspreadsheet.current = null;
            }
        }
    }

    if (libraryBase.jspreadsheet.current && mouseButton == 1) {
        if (e.target.classList.contains('jss_selectall')) {
            if (libraryBase.jspreadsheet.current) {
                selectAll.call(libraryBase.jspreadsheet.current);
            }
        } else if (e.target.classList.contains('jss_corner')) {
            if (libraryBase.jspreadsheet.current.options.editable != false) {
                libraryBase.jspreadsheet.current.selectedCorner = true;
            }
        } else {
            // Header found
            if (jssTable[1] == 1) {
                const columnText = e.target.getAttribute('data-x');
                if (columnText) {
                    const columnId = parseInt(columnText);
                    // Update cursor
                    const info = e.target.getBoundingClientRect();
                    if (libraryBase.jspreadsheet.current.options.columnResize != false && info.width - e.offsetX < 6) {
                        // Resize helper
                        libraryBase.jspreadsheet.current.resizing = {
                            mousePosition: e.pageX,
                            column: columnId,
                            width: info.width,
                        };

                        // Border indication
                        libraryBase.jspreadsheet.current.headers[columnId].classList.add('resizing');
                        for (let j = 0; j < libraryBase.jspreadsheet.current.records.length; j++) {
                            if (libraryBase.jspreadsheet.current.records[j][columnId]) {
                                libraryBase.jspreadsheet.current.records[j][columnId].element.classList.add('resizing');
                            }
                        }
                    } else if (libraryBase.jspreadsheet.current.options.columnDrag != false && info.height - e.offsetY < 6) {
                        if (isColMerged.call(libraryBase.jspreadsheet.current, columnId).length) {
                            console.error('Jspreadsheet: This column is part of a merged cell.');
                        } else {
                            // Reset selection
                            libraryBase.jspreadsheet.current.resetSelection();
                            // Drag helper
                            libraryBase.jspreadsheet.current.dragging = {
                                element: e.target as HTMLTableCellElement,
                                column: columnId,
                                destination: columnId,
                            };
                            // Border indication
                            libraryBase.jspreadsheet.current.headers[columnId].classList.add('dragging');
                            for (let j = 0; j < libraryBase.jspreadsheet.current.records.length; j++) {
                                if (libraryBase.jspreadsheet.current.records[j][columnId]) {
                                    libraryBase.jspreadsheet.current.records[j][columnId].element.classList.add('dragging');
                                }
                            }
                        }
                    } else {
                        let o, d;

                        if (libraryBase.jspreadsheet.current.selectedHeader != null && (e.shiftKey || e.ctrlKey)) {
                            o = libraryBase.jspreadsheet.current.selectedHeader;
                            d = columnId;
                        } else {
                            // Press to rename
                            if (
                                libraryBase.jspreadsheet.current.selectedHeader == columnId &&
                                libraryBase.jspreadsheet.current.options.allowRenameColumn != false
                            ) {
                                libraryBase.jspreadsheet.timeControl = setTimeout(function () {
                                    libraryBase.jspreadsheet.current!.setHeader(columnId);
                                }, 800);
                            }

                            // Keep track of which header was selected first
                            libraryBase.jspreadsheet.current.selectedHeader = columnId;

                            // Update selection single column
                            o = columnId;
                            d = columnId;
                        }

                        // Update selection
                        updateSelectionFromCoords.call(libraryBase.jspreadsheet.current, o, 0, d, libraryBase.jspreadsheet.current.options.data.length - 1, e);
                    }
                } else {
                    if ((e.target.parentNode as HTMLElement).classList.contains('jss_nested')) {
                        let c1, c2;

                        if (e.target.getAttribute('data-column')) {
                            const column = e.target.getAttribute('data-column')!.split(',');
                            c1 = parseInt(column[0]);
                            c2 = parseInt(column[column.length - 1]);
                        } else {
                            c1 = 0;
                            c2 = libraryBase.jspreadsheet.current.options.columns.length - 1;
                        }
                        updateSelectionFromCoords.call(
                            libraryBase.jspreadsheet.current,
                            c1,
                            0,
                            c2,
                            libraryBase.jspreadsheet.current.options.data.length - 1,
                            e
                        );
                    }
                }
            } else {
                libraryBase.jspreadsheet.current.selectedHeader = null;
            }

            // Body found
            if (jssTable[1] == 2) {
                const rowId = parseInt(e.target.getAttribute('data-y'));

                if (e.target.classList.contains('jss_row')) {
                    const info = e.target.getBoundingClientRect();
                    if (libraryBase.jspreadsheet.current.options.rowResize != false && info.height - e.offsetY < 6) {
                        // Resize helper
                        libraryBase.jspreadsheet.current.resizing = {
                            element: e.target.parentNode as HTMLTableRowElement,
                            mousePosition: e.pageY,
                            row: rowId,
                            height: info.height,
                        };
                        // Border indication
                        (e.target.parentNode as HTMLElement).classList.add('resizing');
                    } else if (libraryBase.jspreadsheet.current.options.rowDrag != false && info.width - e.offsetX < 6) {
                        if (isRowMerged.call(libraryBase.jspreadsheet.current, rowId).length) {
                            console.error('Jspreadsheet: This row is part of a merged cell');
                        } else if (libraryBase.jspreadsheet.current.options.search == true && libraryBase.jspreadsheet.current.results) {
                            console.error('Jspreadsheet: Please clear your search before perform this action');
                        } else {
                            // Reset selection
                            libraryBase.jspreadsheet.current.resetSelection();
                            // Drag helper
                            libraryBase.jspreadsheet.current.dragging = {
                                element: e.target.parentNode as HTMLTableRowElement,
                                row: rowId,
                                destination: rowId,
                            };
                            // Border indication
                            (e.target.parentNode as HTMLElement).classList.add('dragging');
                        }
                    } else {
                        let o, d;
                        if (libraryBase.jspreadsheet.current.selectedRow != null && (e.shiftKey || e.ctrlKey)) {
                            o = libraryBase.jspreadsheet.current.selectedRow;
                            d = rowId;
                        } else {
                            // Keep track of which header was selected first
                            libraryBase.jspreadsheet.current.selectedRow = rowId;

                            // Update selection single column
                            o = rowId;
                            d = rowId;
                        }

                        // Update selection
                        updateSelectionFromCoords.call(libraryBase.jspreadsheet.current, null, o, null, d, e);
                    }
                } else {
                    // Jclose
                    if (e.target.classList.contains('jclose') && e.target.clientWidth - e.offsetX < 50 && e.offsetY < 50) {
                        closeEditor.call(libraryBase.jspreadsheet.current, libraryBase.jspreadsheet.current.edition![0], true);
                    } else {
                        const getCellCoords = function (element: HTMLElement) {
                            const x = element.getAttribute('data-x');
                            const y = element.getAttribute('data-y');
                            if (x && y) {
                                return [parseInt(x), parseInt(y)];
                            } else {
                                if (element.parentNode) {
                                    return getCellCoords(element.parentNode as HTMLElement);
                                }
                            }
                        };

                        const position = getCellCoords(e.target);
                        if (position) {
                            const columnId = position[0];
                            const rowId = position[1];
                            // Close edition
                            if (libraryBase.jspreadsheet.current.edition) {
                                // The edition keeps the text of data-x and data-y
                                if (
                                    parseInt(libraryBase.jspreadsheet.current.edition[2]) != columnId ||
                                    parseInt(libraryBase.jspreadsheet.current.edition[3]) != rowId
                                ) {
                                    closeEditor.call(libraryBase.jspreadsheet.current, libraryBase.jspreadsheet.current.edition[0], true);
                                }
                            }

                            if (!libraryBase.jspreadsheet.current.edition) {
                                // Update cell selection
                                if (e.shiftKey) {
                                    updateSelectionFromCoords.call(
                                        libraryBase.jspreadsheet.current,
                                        libraryBase.jspreadsheet.current.selectedCell![0],
                                        libraryBase.jspreadsheet.current.selectedCell![1],
                                        columnId,
                                        rowId,
                                        e
                                    );
                                } else {
                                    updateSelectionFromCoords.call(libraryBase.jspreadsheet.current, columnId, rowId, undefined, undefined, e);
                                }
                            }

                            // No full row selected
                            libraryBase.jspreadsheet.current.selectedHeader = null;
                            libraryBase.jspreadsheet.current.selectedRow = null;
                        }
                    }
                }
            } else {
                libraryBase.jspreadsheet.current.selectedRow = null;
            }

            // Pagination
            if (e.target.classList.contains('jss_page')) {
                if (e.target.textContent == '<') {
                    libraryBase.jspreadsheet.current.page(0);
                } else if (e.target.textContent == '>') {
                    libraryBase.jspreadsheet.current.page(parseInt(e.target.getAttribute('title')) - 1);
                } else {
                    libraryBase.jspreadsheet.current.page(parseInt(e.target.textContent) - 1);
                }
            }
        }

        if (libraryBase.jspreadsheet.current.edition) {
            libraryBase.jspreadsheet.isMouseAction = false;
        } else {
            libraryBase.jspreadsheet.isMouseAction = true;
        }
    } else {
        libraryBase.jspreadsheet.isMouseAction = false;
    }
};

// Mouse move controls
const mouseMoveControls = function (e: TargetEvent<MouseEvent>) {
    e = e || window.event;

    let mouseButton;

    if (e.buttons) {
        mouseButton = e.buttons;
    } else if (e.button) {
        mouseButton = e.button;
    } else {
        mouseButton = e.which;
    }

    if (!mouseButton) {
        libraryBase.jspreadsheet.isMouseAction = false;
    }

    if (libraryBase.jspreadsheet.current) {
        if (libraryBase.jspreadsheet.isMouseAction == true) {
            // Resizing is ongoing
            if (libraryBase.jspreadsheet.current.resizing) {
                if (libraryBase.jspreadsheet.current.resizing.column != null) {
                    const width = e.pageX - libraryBase.jspreadsheet.current.resizing.mousePosition;

                    if (libraryBase.jspreadsheet.current.resizing.width + width > 0) {
                        const tempWidth = libraryBase.jspreadsheet.current.resizing.width + width;
                        libraryBase.jspreadsheet.current.cols[libraryBase.jspreadsheet.current.resizing.column].colElement.setAttribute('width', tempWidth);

                        updateCornerPosition.call(libraryBase.jspreadsheet.current);
                    }
                } else {
                    const height = e.pageY - libraryBase.jspreadsheet.current.resizing.mousePosition;

                    if (libraryBase.jspreadsheet.current.resizing.height + height > 0) {
                        const tempHeight = libraryBase.jspreadsheet.current.resizing.height + height;
                        libraryBase.jspreadsheet.current.rows[libraryBase.jspreadsheet.current.resizing.row].element.setAttribute('height', tempHeight);

                        updateCornerPosition.call(libraryBase.jspreadsheet.current);
                    }
                }
            } else if (libraryBase.jspreadsheet.current.dragging) {
                if (libraryBase.jspreadsheet.current.dragging.column != null) {
                    const columnId = getIndex(e.target, 'data-x');
                    if (columnId != null) {
                        if (isColMerged.call(libraryBase.jspreadsheet.current, columnId).length) {
                            console.error('Jspreadsheet: This column is part of a merged cell.');
                        } else {
                            for (let i = 0; i < libraryBase.jspreadsheet.current.headers.length; i++) {
                                libraryBase.jspreadsheet.current.headers[i].classList.remove('dragging-left');
                                libraryBase.jspreadsheet.current.headers[i].classList.remove('dragging-right');
                            }

                            if (libraryBase.jspreadsheet.current.dragging.column == columnId) {
                                libraryBase.jspreadsheet.current.dragging.destination = columnId;
                            } else {
                                if (e.target.clientWidth / 2 > e.offsetX) {
                                    if (libraryBase.jspreadsheet.current.dragging.column < columnId) {
                                        libraryBase.jspreadsheet.current.dragging.destination = columnId - 1;
                                    } else {
                                        libraryBase.jspreadsheet.current.dragging.destination = columnId;
                                    }
                                    libraryBase.jspreadsheet.current.headers[columnId].classList.add('dragging-left');
                                } else {
                                    if (libraryBase.jspreadsheet.current.dragging.column < columnId) {
                                        libraryBase.jspreadsheet.current.dragging.destination = columnId;
                                    } else {
                                        libraryBase.jspreadsheet.current.dragging.destination = columnId + 1;
                                    }
                                    libraryBase.jspreadsheet.current.headers[columnId].classList.add('dragging-right');
                                }
                            }
                        }
                    }
                } else {
                    const rowId = getIndex(e.target, 'data-y');
                    if (rowId != null) {
                        if (isRowMerged.call(libraryBase.jspreadsheet.current, rowId).length) {
                            console.error('Jspreadsheet: This row is part of a merged cell.');
                        } else {
                            const target = e.target.clientHeight / 2 > e.offsetY ? e.target.parentNode!.nextSibling : e.target.parentNode;
                            if (libraryBase.jspreadsheet.current.dragging.element != target) {
                                e.target.parentNode!.parentNode!.insertBefore(libraryBase.jspreadsheet.current.dragging.element, target);
                                libraryBase.jspreadsheet.current.dragging.destination = Array.prototype.indexOf.call(
                                    libraryBase.jspreadsheet.current.dragging.element.parentNode!.children,
                                    libraryBase.jspreadsheet.current.dragging.element
                                );
                            }
                        }
                    }
                }
            }
        } else {
            const x = e.target.getAttribute('data-x');
            const y = e.target.getAttribute('data-y');
            const rect = e.target.getBoundingClientRect();

            if (libraryBase.jspreadsheet.current.cursor) {
                libraryBase.jspreadsheet.current.cursor.style.cursor = '';
                libraryBase.jspreadsheet.current.cursor = null;
            }

            if (e.target.parentNode!.parentNode && (e.target.parentNode!.parentNode as HTMLElement).className) {
                if ((e.target.parentNode!.parentNode as HTMLElement).classList.contains('resizable')) {
                    if (e.target && x && !y && rect.width - (e.clientX - rect.left) < 6) {
                        libraryBase.jspreadsheet.current.cursor = e.target;
                        libraryBase.jspreadsheet.current.cursor!.style.cursor = 'col-resize';
                    } else if (e.target && !x && y && rect.height - (e.clientY - rect.top) < 6) {
                        libraryBase.jspreadsheet.current.cursor = e.target;
                        libraryBase.jspreadsheet.current.cursor!.style.cursor = 'row-resize';
                    }
                }

                if ((e.target.parentNode!.parentNode as HTMLElement).classList.contains('draggable')) {
                    if (e.target && !x && y && rect.width - (e.clientX - rect.left) < 6) {
                        libraryBase.jspreadsheet.current.cursor = e.target;
                        libraryBase.jspreadsheet.current.cursor!.style.cursor = 'move';
                    } else if (e.target && x && !y && rect.height - (e.clientY - rect.top) < 6) {
                        libraryBase.jspreadsheet.current.cursor = e.target;
                        libraryBase.jspreadsheet.current.cursor!.style.cursor = 'move';
                    }
                }
            }
        }
    }
};

/**
 * Update copy selection
 *
 * @param int x, y
 * @return void
 */
const updateCopySelection = function (this: Worksheet, x3: number, y3: number) {
    const obj = this;

    // Remove selection
    removeCopySelection.call(obj);

    // Get elements first and last
    const x1 = obj.selectedContainer![0];
    const y1 = obj.selectedContainer![1];
    const x2 = obj.selectedContainer![2];
    const y2 = obj.selectedContainer![3];

    if (x3 != null && y3 != null) {
        let px, ux;

        if (x3 - x2 > 0) {
            px = parseInt(x2) + 1;
            ux = parseInt(x3);
        } else {
            px = parseInt(x3);
            ux = parseInt(x1) - 1;
        }

        let py, uy;

        if (y3 - y2 > 0) {
            py = parseInt(y2) + 1;
            uy = parseInt(y3);
        } else {
            py = parseInt(y3);
            uy = parseInt(y1) - 1;
        }

        if (ux - px <= uy - py) {
            px = parseInt(x1);
            ux = parseInt(x2);
        } else {
            py = parseInt(y1);
            uy = parseInt(y2);
        }

        for (let j = py; j <= uy; j++) {
            for (let i = px; i <= ux; i++) {
                if (obj.records[j][i] && obj.rows[j].element.style.display != 'none' && obj.records[j][i].element.style.display != 'none') {
                    obj.records[j][i].element.classList.add('selection');
                    obj.records[py][i].element.classList.add('selection-top');
                    obj.records[uy][i].element.classList.add('selection-bottom');
                    obj.records[j][px].element.classList.add('selection-left');
                    obj.records[j][ux].element.classList.add('selection-right');

                    // Persist selected elements
                    obj.selection.push(obj.records[j][i].element);
                }
            }
        }
    }
};

const mouseOverControls = function (e: TargetEvent<MouseEvent>) {
    e = e || window.event;

    let mouseButton;

    if (e.buttons) {
        mouseButton = e.buttons;
    } else if (e.button) {
        mouseButton = e.button;
    } else {
        mouseButton = e.which;
    }

    if (!mouseButton) {
        libraryBase.jspreadsheet.isMouseAction = false;
    }

    if (libraryBase.jspreadsheet.current && libraryBase.jspreadsheet.isMouseAction == true) {
        // Get elements
        const jssTable = getElement(e.target);

        if (jssTable[0]) {
            // Avoid cross reference
            if (libraryBase.jspreadsheet.current != jssTable[0].jssWorksheet) {
                if (libraryBase.jspreadsheet.current) {
                    return false;
                }
            }

            let columnId = getIndex(e.target, 'data-x');
            const rowId = getIndex(e.target, 'data-y');
            if (libraryBase.jspreadsheet.current.resizing || libraryBase.jspreadsheet.current.dragging) {
                // ignore
            } else {
                // Header found
                if (jssTable[1] == 1) {
                    if (libraryBase.jspreadsheet.current.selectedHeader != null) {
                        columnId = getIndex(e.target, 'data-x');
                        const o = libraryBase.jspreadsheet.current.selectedHeader;
                        const d = columnId;
                        // Update selection
                        updateSelectionFromCoords.call(libraryBase.jspreadsheet.current, o, 0, d, libraryBase.jspreadsheet.current.options.data.length - 1, e);
                    }
                }

                // Body found
                if (jssTable[1] == 2) {
                    if (e.target.classList.contains('jss_row')) {
                        if (libraryBase.jspreadsheet.current.selectedRow != null) {
                            const o = libraryBase.jspreadsheet.current.selectedRow;
                            const d = rowId;
                            // Update selection
                            updateSelectionFromCoords.call(
                                libraryBase.jspreadsheet.current,
                                0,
                                o,
                                libraryBase.jspreadsheet.current.options.data[0].length - 1,
                                d,
                                e
                            );
                        }
                    } else {
                        // Do not select edtion is in progress
                        if (!libraryBase.jspreadsheet.current.edition) {
                            if (columnId != null && rowId != null) {
                                if (libraryBase.jspreadsheet.current.selectedCorner) {
                                    updateCopySelection.call(libraryBase.jspreadsheet.current, columnId, rowId);
                                } else {
                                    if (libraryBase.jspreadsheet.current.selectedCell) {
                                        updateSelectionFromCoords.call(
                                            libraryBase.jspreadsheet.current,
                                            libraryBase.jspreadsheet.current.selectedCell[0],
                                            libraryBase.jspreadsheet.current.selectedCell[1],
                                            columnId,
                                            rowId,
                                            e
                                        );
                                    }
                                }
                            }
                        }
                    }
                }
            }
        }
    }

    // Clear any time control
    if (libraryBase.jspreadsheet.timeControl) {
        clearTimeout(libraryBase.jspreadsheet.timeControl);
        libraryBase.jspreadsheet.timeControl = null;
    }
};

const doubleClickControls = function (e: TargetEvent<MouseEvent>) {
    // Jss is selected
    if (libraryBase.jspreadsheet.current) {
        // Corner action
        if (e.target.classList.contains('jss_corner')) {
            // Any selected cells
            if (libraryBase.jspreadsheet.current.highlighted.length > 0) {
                // Copy from this
                const x1 = parseInt(libraryBase.jspreadsheet.current.highlighted[0].element.getAttribute('data-x'));
                const y1 =
                    parseInt(
                        libraryBase.jspreadsheet.current.highlighted[libraryBase.jspreadsheet.current.highlighted.length - 1].element.getAttribute('data-y')
                    ) + 1;
                // Until this
                const x2 = parseInt(
                    libraryBase.jspreadsheet.current.highlighted[libraryBase.jspreadsheet.current.highlighted.length - 1].element.getAttribute('data-x')
                );
                const y2 = libraryBase.jspreadsheet.current.records.length - 1;
                // Execute copy
                copyData.call(
                    libraryBase.jspreadsheet.current,
                    libraryBase.jspreadsheet.current.records[y1][x1].element,
                    libraryBase.jspreadsheet.current.records[y2][x2].element
                );
            }
        } else if (e.target.classList.contains('jss_column_filter')) {
            // Column
            const columnId = parseInt(e.target.getAttribute('data-x'));
            // Open filter
            openFilter.call(libraryBase.jspreadsheet.current, columnId);
        } else {
            // Get table
            const jssTable = getElement(e.target);

            // Double click over header
            if (jssTable[1] == 1 && libraryBase.jspreadsheet.current.options.columnSorting != false) {
                // Check valid column header coords
                const columnId = getIndex(e.target, 'data-x');
                if (columnId != null) {
                    libraryBase.jspreadsheet.current.orderBy(columnId);
                }
            }

            // Double click over body
            if (jssTable[1] == 2 && libraryBase.jspreadsheet.current.options.editable != false) {
                if (!libraryBase.jspreadsheet.current.edition) {
                    const getCellCoords = function (element: HTMLElement) {
                        if (element.parentNode) {
                            const x = element.getAttribute('data-x');
                            const y = element.getAttribute('data-y');
                            if (x && y) {
                                return element;
                            } else {
                                return getCellCoords(element.parentNode as HTMLElement);
                            }
                        }
                    };
                    const cell = getCellCoords(e.target);
                    if (cell && cell.classList.contains('highlight')) {
                        openEditor.call(libraryBase.jspreadsheet.current, cell as HTMLTableCellElement, undefined, e);
                    }
                }
            }
        }
    }
};

const pasteControls = function (e?: TargetEvent<ClipboardEvent>) {
    // A paste on a control with the class jss_object outside the worksheet goes to that control
    if (e && e.target && typeof e.target.closest === 'function' && !getElement(e.target)[0] && e.target.closest('.jss_object')) {
        return;
    }

    if (libraryBase.jspreadsheet.current && libraryBase.jspreadsheet.current.selectedCell) {
        if (!libraryBase.jspreadsheet.current.edition) {
            if (libraryBase.jspreadsheet.current.options.editable != false) {
                if (e && e.clipboardData) {
                    paste.call(
                        libraryBase.jspreadsheet.current,
                        libraryBase.jspreadsheet.current.selectedCell[0],
                        libraryBase.jspreadsheet.current.selectedCell[1],
                        e.clipboardData.getData('text')
                    );
                    e.preventDefault();
                } else if ((window as Window & { clipboardData?: DataTransfer }).clipboardData) {
                    // Internet Explorer has the clipboard on the window
                    paste.call(
                        libraryBase.jspreadsheet.current,
                        libraryBase.jspreadsheet.current.selectedCell[0],
                        libraryBase.jspreadsheet.current.selectedCell[1],
                        (window as Window & { clipboardData?: DataTransfer }).clipboardData!.getData('text')
                    );
                }
            }
        }
    }
};

const getRole = function (element: HTMLElement) {
    if (element.classList.contains('jss_selectall')) {
        return 'select-all';
    }

    if (element.classList.contains('jss_corner')) {
        return 'fill-handle';
    }

    let tempElement = element;

    while (!tempElement.classList.contains('jss_spreadsheet')) {
        if (tempElement.classList.contains('jss_row')) {
            return 'row';
        }

        if (tempElement.classList.contains('jss_nested')) {
            return 'nested';
        }

        if (tempElement.classList.contains('jtabs-headers')) {
            return 'tabs';
        }

        if (tempElement.classList.contains('jtoolbar')) {
            return 'toolbar';
        }

        if (tempElement.classList.contains('jss_pagination')) {
            return 'pagination';
        }

        if (tempElement.tagName === 'TBODY') {
            return 'cell';
        }

        if (tempElement.tagName === 'TFOOT') {
            return getElementIndex(element) === 0 ? 'grid' : 'footer';
        }

        if (tempElement.tagName === 'THEAD') {
            return 'header';
        }

        // The element is in a worksheet, so the loop ends at the element of the spreadsheet
        tempElement = tempElement.parentElement!;
    }

    return 'applications';
};

const defaultContextMenu = function (worksheet: Worksheet, x: number, y: number, role: jspreadsheet.ContextMenuRole): jspreadsheet.ContextMenuItem[] {
    const items: jspreadsheet.ContextMenuItem[] = [];

    if (role === 'header') {
        // Insert a new column
        if (worksheet.options.allowInsertColumn != false) {
            items.push({
                title: jSuites.translate('Insert a new column before'),
                onclick: function () {
                    worksheet.insertColumn(1, parseInt(x), 1);
                },
            });
        }

        if (worksheet.options.allowInsertColumn != false) {
            items.push({
                title: jSuites.translate('Insert a new column after'),
                onclick: function () {
                    worksheet.insertColumn(1, parseInt(x), 0);
                },
            });
        }

        // Delete a column
        if (worksheet.options.allowDeleteColumn != false) {
            items.push({
                title: jSuites.translate('Delete selected columns'),
                onclick: function () {
                    worksheet.deleteColumn(worksheet.getSelectedColumns().length ? undefined : parseInt(x));
                },
            });
        }

        // Rename column
        if (worksheet.options.allowRenameColumn != false) {
            items.push({
                title: jSuites.translate('Rename this column'),
                onclick: function () {
                    const oldValue = worksheet.getHeader(x);

                    const newValue = prompt(jSuites.translate('Column name'), oldValue);

                    // A cancelled prompt gives null, and setHeader then uses the default name (issue #42)
                    worksheet.setHeader(x, newValue!);
                },
            });
        }

        // Sorting
        if (worksheet.options.columnSorting != false) {
            // Line
            items.push({ type: 'line' });

            items.push({
                title: jSuites.translate('Order ascending'),
                onclick: function () {
                    worksheet.orderBy(x, 0);
                },
            });
            items.push({
                title: jSuites.translate('Order descending'),
                onclick: function () {
                    worksheet.orderBy(x, 1);
                },
            });
        }
    }

    if (role === 'row' || role === 'cell') {
        // Insert new row
        if (worksheet.options.allowInsertRow != false) {
            items.push({
                title: jSuites.translate('Insert a new row before'),
                onclick: function () {
                    worksheet.insertRow(1, parseInt(y), 1);
                },
            });

            items.push({
                title: jSuites.translate('Insert a new row after'),
                onclick: function () {
                    worksheet.insertRow(1, parseInt(y));
                },
            });
        }

        if (worksheet.options.allowDeleteRow != false) {
            items.push({
                title: jSuites.translate('Delete selected rows'),
                onclick: function () {
                    worksheet.deleteRow(worksheet.getSelectedRows().length ? undefined : parseInt(y));
                },
            });
        }
    }

    if (role === 'cell') {
        if (worksheet.options.allowComments != false) {
            items.push({ type: 'line' });

            const title = worksheet.records[y][x].element.getAttribute('title') || '';

            items.push({
                title: jSuites.translate(title ? 'Edit comments' : 'Add comments'),
                onclick: function () {
                    const comment = prompt(jSuites.translate('Comments'), title);
                    if (comment) {
                        worksheet.setComments(getCellNameFromCoords(x, y), comment);
                    }
                },
            });

            if (title) {
                items.push({
                    title: jSuites.translate('Clear comments'),
                    onclick: function () {
                        worksheet.setComments(getCellNameFromCoords(x, y), '');
                    },
                });
            }
        }
    }

    // Line
    if (items.length !== 0) {
        items.push({ type: 'line' });
    }

    // Copy
    if (role === 'header' || role === 'row' || role === 'cell') {
        items.push({
            title: jSuites.translate('Copy') + '...',
            shortcut: 'Ctrl + C',
            onclick: function () {
                copy.call(worksheet, true);
            },
        });

        // Paste
        if (navigator && navigator.clipboard) {
            items.push({
                title: jSuites.translate('Paste') + '...',
                shortcut: 'Ctrl + V',
                onclick: function () {
                    if (worksheet.selectedCell) {
                        navigator.clipboard.readText().then(function (text) {
                            if (text) {
                                paste.call(worksheet, worksheet.selectedCell![0], worksheet.selectedCell![1], text);
                            }
                        });
                    }
                },
            });
        }
    }

    // Save
    if (worksheet.parent.config.allowExport != false) {
        items.push({
            title: jSuites.translate('Save as') + '...',
            shortcut: 'Ctrl + S',
            onclick: function () {
                worksheet.download();
            },
        });
    }

    // About
    if (worksheet.parent.config.about != false) {
        items.push({
            title: jSuites.translate('About'),
            onclick: function () {
                if (typeof worksheet.parent.config.about === 'undefined' || worksheet.parent.config.about === true) {
                    alert(version.print());
                } else {
                    alert(worksheet.parent.config.about);
                }
            },
        });
    }

    return items;
};

const getElementIndex = function (element: HTMLElement) {
    const parentChildren = element.parentElement!.children;

    for (let i = 0; i < parentChildren.length; i++) {
        const currentElement = parentChildren[i];

        if (element === currentElement) {
            return i;
        }
    }

    return -1;
};

const contextMenuControls = function (e: TargetEvent<PointerEvent>) {
    e = e || window.event;

    let mouseButton;

    if ('buttons' in e) {
        mouseButton = e.buttons;
    } else {
        mouseButton = (e as MouseEvent).which || (e as MouseEvent).button;
    }

    if (libraryBase.jspreadsheet.current) {
        const spreadsheet = libraryBase.jspreadsheet.current.parent;

        if (libraryBase.jspreadsheet.current.edition) {
            e.preventDefault();
        } else {
            spreadsheet.contextMenu.contextmenu.close();

            if (libraryBase.jspreadsheet.current) {
                const role = getRole(e.target);

                // Text from data-x and data-y, or an index. Number() converts them as a comparison does: null is 0.
                let x: string | number | null = null,
                    y: string | number | null = null;

                if (role === 'cell') {
                    let cellElement = e.target;
                    while (cellElement.tagName !== 'TD') {
                        cellElement = cellElement.parentNode as HTMLElement;
                    }

                    y = cellElement.getAttribute('data-y');
                    x = cellElement.getAttribute('data-x');

                    if (
                        !libraryBase.jspreadsheet.current.selectedCell ||
                        Number(x) < parseInt(libraryBase.jspreadsheet.current.selectedCell[0]) ||
                        Number(x) > parseInt(libraryBase.jspreadsheet.current.selectedCell[2]) ||
                        Number(y) < parseInt(libraryBase.jspreadsheet.current.selectedCell[1]) ||
                        Number(y) > parseInt(libraryBase.jspreadsheet.current.selectedCell[3])
                    ) {
                        updateSelectionFromCoords.call(libraryBase.jspreadsheet.current, parseInt(x), parseInt(y), parseInt(x), parseInt(y), e);
                    }
                } else if (role === 'row' || role === 'header') {
                    if (role === 'row') {
                        y = e.target.getAttribute('data-y');
                    } else {
                        x = e.target.getAttribute('data-x');
                    }

                    if (
                        !libraryBase.jspreadsheet.current.selectedCell ||
                        Number(x) < parseInt(libraryBase.jspreadsheet.current.selectedCell[0]) ||
                        Number(x) > parseInt(libraryBase.jspreadsheet.current.selectedCell[2]) ||
                        Number(y) < parseInt(libraryBase.jspreadsheet.current.selectedCell[1]) ||
                        Number(y) > parseInt(libraryBase.jspreadsheet.current.selectedCell[3])
                    ) {
                        // A null index selects the whole row or column
                        const column = x == null ? null : parseInt(x);
                        const row = y == null ? null : parseInt(y);
                        updateSelectionFromCoords.call(libraryBase.jspreadsheet.current, column, row, column, row, e);
                    }
                } else if (role === 'nested') {
                    // The indexes of the columns under the nested header
                    const columns = e.target
                        .getAttribute('data-column')!
                        .split(',')
                        .map((column) => parseInt(column));

                    x = getElementIndex(e.target) - 1;
                    y = getElementIndex(e.target.parentElement!);

                    if (
                        !libraryBase.jspreadsheet.current.selectedCell ||
                        columns[0] != parseInt(libraryBase.jspreadsheet.current.selectedCell[0]) ||
                        columns[columns.length - 1] != parseInt(libraryBase.jspreadsheet.current.selectedCell[2]) ||
                        libraryBase.jspreadsheet.current.selectedCell[1] != null ||
                        libraryBase.jspreadsheet.current.selectedCell[3] != null
                    ) {
                        updateSelectionFromCoords.call(libraryBase.jspreadsheet.current, columns[0], null, columns[columns.length - 1], null, e);
                    }
                } else if (role === 'select-all') {
                    selectAll.call(libraryBase.jspreadsheet.current);
                } else if (role === 'tabs') {
                    x = getElementIndex(e.target);
                } else if (role === 'footer') {
                    x = getElementIndex(e.target) - 1;
                    y = getElementIndex(e.target.parentElement!);
                }

                // Table found
                let items = defaultContextMenu(libraryBase.jspreadsheet.current, parseInt(x), parseInt(y), role);

                if (typeof spreadsheet.config.contextMenu === 'function') {
                    const result = spreadsheet.config.contextMenu(libraryBase.jspreadsheet.current, x, y, e, items, role, x, y);

                    if (result) {
                        items = result;
                    } else if (result === false) {
                        return;
                    }
                }

                if (typeof spreadsheet.plugins === 'object') {
                    Object.entries(spreadsheet.plugins).forEach(function ([, plugin]) {
                        if (typeof plugin.contextMenu === 'function') {
                            const result = plugin.contextMenu(
                                libraryBase.jspreadsheet.current!,
                                x !== null ? parseInt(x) : null,
                                y !== null ? parseInt(y) : null,
                                e,
                                items,
                                role,
                                x !== null ? parseInt(x) : null,
                                y !== null ? parseInt(y) : null
                            );

                            if (result) {
                                items = result;
                            }
                        }
                    });
                }

                // The id is depending on header and body
                spreadsheet.contextMenu.contextmenu.open(e, items);
                // Avoid the real one
                e.preventDefault();
            }
        }
    }
};

const touchStartControls = function (e: TargetEvent<TouchEvent>) {
    const jssTable = getElement(e.target);

    if (jssTable[0]) {
        if (libraryBase.jspreadsheet.current != jssTable[0].jssWorksheet) {
            if (libraryBase.jspreadsheet.current) {
                libraryBase.jspreadsheet.current.resetSelection();
            }
            libraryBase.jspreadsheet.current = jssTable[0].jssWorksheet;
        }
    } else {
        if (libraryBase.jspreadsheet.current && !e.target.closest('.jss_object')) {
            libraryBase.jspreadsheet.current.resetSelection();
            libraryBase.jspreadsheet.current = null;
        }
    }

    if (libraryBase.jspreadsheet.current && jssTable[0]) {
        if (!libraryBase.jspreadsheet.current.edition) {
            const columnId = getIndex(e.target, 'data-x');
            const rowId = getIndex(e.target, 'data-y');

            if (columnId != null && rowId != null) {
                updateSelectionFromCoords.call(libraryBase.jspreadsheet.current, columnId, rowId, undefined, undefined, e);

                if (libraryBase.jspreadsheet.current.options.editable == false) {
                    return;
                }

                libraryBase.jspreadsheet.timeControl = setTimeout(function () {
                    // Keep temporary reference to the element
                    const column =
                        libraryBase.jspreadsheet.current!.options.columns[getCellColumnIndex.call(libraryBase.jspreadsheet.current!, columnId, rowId)];
                    if (column.type == 'color') {
                        libraryBase.jspreadsheet.tmpElement = null;
                    } else {
                        libraryBase.jspreadsheet.tmpElement = e.target;
                    }
                    openEditor.call(libraryBase.jspreadsheet.current!, e.target as HTMLTableCellElement, false, e);
                }, 500);
            }
        }
    }
};

const touchEndControls = function (e: TargetEvent<TouchEvent>) {
    // Clear any time control
    if (libraryBase.jspreadsheet.timeControl) {
        clearTimeout(libraryBase.jspreadsheet.timeControl);
        libraryBase.jspreadsheet.timeControl = null;
        // Element
        if (libraryBase.jspreadsheet.tmpElement && libraryBase.jspreadsheet.tmpElement.children[0].tagName == 'INPUT') {
            libraryBase.jspreadsheet.tmpElement.children[0].focus();
        }
        libraryBase.jspreadsheet.tmpElement = null;
    }
};

export const cutControls = function (e?: TargetEvent<ClipboardEvent>) {
    if (libraryBase.jspreadsheet.current) {
        if (!libraryBase.jspreadsheet.current.edition) {
            copy.call(libraryBase.jspreadsheet.current, true, undefined, undefined, undefined, undefined, true);
            if (libraryBase.jspreadsheet.current.options.editable != false) {
                libraryBase.jspreadsheet.current.setValue(
                    libraryBase.jspreadsheet.current.highlighted.map(function (record) {
                        return record.element;
                    }),
                    ''
                );
            }
        }
    }
};

const copyControls = function (e?: TargetEvent<ClipboardEvent>) {
    if (libraryBase.jspreadsheet.current) {
        if (!libraryBase.jspreadsheet.current.edition) {
            copy.call(libraryBase.jspreadsheet.current, true);
        }
    }
};

const isMac = function () {
    return navigator.platform.toUpperCase().indexOf('MAC') >= 0;
};

const isCtrl = function (e: TargetEvent<KeyboardEvent | MouseEvent>) {
    if (isMac()) {
        return e.metaKey;
    } else {
        return e.ctrlKey;
    }
};

const keyDownControls = function (e: TargetEvent<KeyboardEvent>) {
    // Keys on the worksheet tabs and on the button to add a worksheet do not go to the grid
    const keyPath = typeof e.composedPath === 'function' ? e.composedPath() : [];
    // The first node of the path can be a node without closest, such as the document. The code checks for closest.
    const keyTarget = (keyPath.length ? keyPath[0] : e.target) as HTMLElement;
    if (keyTarget && typeof keyTarget.closest === 'function' && keyTarget.closest('.jss_spreadsheet > .jtabs-headers-container')) {
        return;
    }

    if (libraryBase.jspreadsheet.current) {
        // Keys on a control with the class jss_object outside the worksheet, for example a button next to the grid, also in
        // an open shadow root, do not go to the grid. An editor can have a jss_object control, so this applies only when no
        // editor is open.
        if (!libraryBase.jspreadsheet.current.edition) {
            if (keyTarget && typeof keyTarget.closest === 'function' && !getElement(keyTarget)[0] && keyTarget.closest('.jss_object')) {
                return;
            }
        }

        if (libraryBase.jspreadsheet.current.edition) {
            const current = libraryBase.jspreadsheet.current;
            const column =
                current.options.columns &&
                current.options.columns[getCellColumnIndex.call(current, parseInt(current.edition![2]), parseInt(current.edition![3]))];

            if (e.which == 27) {
                // Escape
                if (libraryBase.jspreadsheet.current.edition) {
                    // Exit without saving
                    closeEditor.call(libraryBase.jspreadsheet.current, libraryBase.jspreadsheet.current.edition[0], false);
                }
                e.preventDefault();
            } else if (e.which == 13) {
                // Enter
                if (column && column.type == 'calendar') {
                    closeEditor.call(libraryBase.jspreadsheet.current, libraryBase.jspreadsheet.current.edition[0], true);
                } else if (column && column.type == 'dropdown') {
                    // Do nothing
                } else {
                    // Alt enter -> do not close editor
                    if (
                        (libraryBase.jspreadsheet.current.options.wordWrap == true ||
                            (column && column.wordWrap == true) ||
                            (libraryBase.jspreadsheet.current.options.data[parseInt(libraryBase.jspreadsheet.current.edition[3])][
                                parseInt(libraryBase.jspreadsheet.current.edition[2])
                            ] &&
                                (
                                    libraryBase.jspreadsheet.current.options.data[parseInt(libraryBase.jspreadsheet.current.edition[3])][
                                        parseInt(libraryBase.jspreadsheet.current.edition[2])
                                    ] as string
                                ).length > 200)) &&
                        e.altKey
                    ) {
                        // Add new line to the editor
                        const editorTextarea = libraryBase.jspreadsheet.current.edition[0].children[0] as HTMLTextAreaElement;
                        let editorValue = (libraryBase.jspreadsheet.current.edition[0].children[0] as HTMLTextAreaElement).value;
                        const editorIndexOf = editorTextarea.selectionStart;
                        editorValue = editorValue.slice(0, editorIndexOf) + '\n' + editorValue.slice(editorIndexOf);
                        editorTextarea.value = editorValue;
                        editorTextarea.focus();
                        editorTextarea.selectionStart = editorIndexOf + 1;
                        editorTextarea.selectionEnd = editorIndexOf + 1;
                    } else {
                        (libraryBase.jspreadsheet.current.edition[0].children[0] as HTMLElement).blur();
                    }
                }
            } else if (e.which == 9) {
                // Tab
                if (column && ['calendar', 'html'].includes(column.type as string)) {
                    closeEditor.call(libraryBase.jspreadsheet.current, libraryBase.jspreadsheet.current.edition[0], true);
                } else {
                    (libraryBase.jspreadsheet.current.edition[0].children[0] as HTMLElement).blur();
                }
            }
        }

        if (!libraryBase.jspreadsheet.current.edition && libraryBase.jspreadsheet.current.selectedCell) {
            // Which key
            if (e.which == 37) {
                left.call(libraryBase.jspreadsheet.current, e.shiftKey, e.ctrlKey);
                e.preventDefault();
            } else if (e.which == 39) {
                right.call(libraryBase.jspreadsheet.current, e.shiftKey, e.ctrlKey);
                e.preventDefault();
            } else if (e.which == 38) {
                up.call(libraryBase.jspreadsheet.current, e.shiftKey, e.ctrlKey);
                e.preventDefault();
            } else if (e.which == 40) {
                down.call(libraryBase.jspreadsheet.current, e.shiftKey, e.ctrlKey);
                e.preventDefault();
            } else if (e.which == 36) {
                first.call(libraryBase.jspreadsheet.current, e.shiftKey, e.ctrlKey);
                e.preventDefault();
            } else if (e.which == 35) {
                last.call(libraryBase.jspreadsheet.current, e.shiftKey, e.ctrlKey);
                e.preventDefault();
            } else if (e.which == 46 || e.which == 8) {
                // Delete
                if (libraryBase.jspreadsheet.current.options.editable != false) {
                    if (libraryBase.jspreadsheet.current.selectedRow != null) {
                        if (libraryBase.jspreadsheet.current.options.allowDeleteRow != false) {
                            if (confirm(jSuites.translate('Are you sure to delete the selected rows?'))) {
                                libraryBase.jspreadsheet.current.deleteRow();
                            }
                        }
                    } else if (libraryBase.jspreadsheet.current.selectedHeader != null) {
                        if (libraryBase.jspreadsheet.current.options.allowDeleteColumn != false) {
                            if (confirm(jSuites.translate('Are you sure to delete the selected columns?'))) {
                                libraryBase.jspreadsheet.current.deleteColumn();
                            }
                        }
                    } else {
                        // Change value
                        libraryBase.jspreadsheet.current.setValue(
                            libraryBase.jspreadsheet.current.highlighted.map(function (record) {
                                return record.element;
                            }),
                            ''
                        );
                    }
                }
            } else if (e.which == 13) {
                // Move cursor
                if (e.shiftKey) {
                    up.call(libraryBase.jspreadsheet.current);
                } else {
                    if (libraryBase.jspreadsheet.current.options.allowInsertRow != false) {
                        if (libraryBase.jspreadsheet.current.options.allowManualInsertRow != false) {
                            if (libraryBase.jspreadsheet.current.selectedCell[1] == libraryBase.jspreadsheet.current.options.data.length - 1) {
                                // New record in case selectedCell in the last row
                                libraryBase.jspreadsheet.current.insertRow();
                            }
                        }
                    }

                    down.call(libraryBase.jspreadsheet.current);
                }
                e.preventDefault();
            } else if (e.which == 9) {
                // Tab
                if (e.shiftKey) {
                    left.call(libraryBase.jspreadsheet.current);
                } else {
                    if (libraryBase.jspreadsheet.current.options.allowInsertColumn != false) {
                        if (libraryBase.jspreadsheet.current.options.allowManualInsertColumn != false) {
                            if (libraryBase.jspreadsheet.current.selectedCell[0] == libraryBase.jspreadsheet.current.options.data[0].length - 1) {
                                // New record in case selectedCell in the last column
                                libraryBase.jspreadsheet.current.insertColumn();
                            }
                        }
                    }

                    right.call(libraryBase.jspreadsheet.current);
                }
                e.preventDefault();
            } else {
                if ((e.ctrlKey || e.metaKey) && !e.shiftKey) {
                    if (e.which == 65) {
                        // Ctrl + A
                        selectAll.call(libraryBase.jspreadsheet.current);
                        e.preventDefault();
                    } else if (e.which == 83) {
                        // Ctrl + S
                        libraryBase.jspreadsheet.current.download();
                        e.preventDefault();
                    } else if (e.which == 89) {
                        // Ctrl + Y
                        libraryBase.jspreadsheet.current.redo();
                        e.preventDefault();
                    } else if (e.which == 90) {
                        // Ctrl + Z
                        libraryBase.jspreadsheet.current.undo();
                        e.preventDefault();
                    } else if (e.which == 67) {
                        // Ctrl + C
                        copy.call(libraryBase.jspreadsheet.current, true);
                        e.preventDefault();
                    } else if (e.which == 88) {
                        // Ctrl + X
                        if (libraryBase.jspreadsheet.current.options.editable != false) {
                            cutControls();
                        } else {
                            copyControls();
                        }
                        e.preventDefault();
                    } else if (e.which == 86) {
                        // Ctrl + V
                        pasteControls();
                    }
                } else {
                    if (libraryBase.jspreadsheet.current.selectedCell) {
                        if (libraryBase.jspreadsheet.current.options.editable != false) {
                            const rowId = libraryBase.jspreadsheet.current.selectedCell[1];
                            const columnId = libraryBase.jspreadsheet.current.selectedCell[0];
                            const column =
                                libraryBase.jspreadsheet.current.options.columns[getCellColumnIndex.call(libraryBase.jspreadsheet.current, columnId, rowId)];

                            // Characters able to start a edition
                            if (e.keyCode == 32) {
                                // Space
                                e.preventDefault();
                                if (column.type == 'checkbox' || column.type == 'radio') {
                                    setCheckRadioValue.call(libraryBase.jspreadsheet.current);
                                } else {
                                    // Start edition
                                    openEditor.call(
                                        libraryBase.jspreadsheet.current,
                                        libraryBase.jspreadsheet.current.records[rowId][columnId].element,
                                        true,
                                        e
                                    );
                                }
                            } else if (e.keyCode == 113) {
                                // Start edition with current content F2
                                openEditor.call(libraryBase.jspreadsheet.current, libraryBase.jspreadsheet.current.records[rowId][columnId].element, false, e);
                            } else if ((e.key.length === 1 || e.key === 'Process') && !(e.altKey || isCtrl(e))) {
                                // Start edition
                                openEditor.call(libraryBase.jspreadsheet.current, libraryBase.jspreadsheet.current.records[rowId][columnId].element, true, e);
                                // Prevent entries in the calendar
                                if (column && column.type == 'calendar') {
                                    e.preventDefault();
                                }
                            }
                        }
                    }
                }
            }
        } else {
            if (e.target.classList.contains('jss_search')) {
                if (libraryBase.jspreadsheet.timeControl) {
                    clearTimeout(libraryBase.jspreadsheet.timeControl);
                }

                libraryBase.jspreadsheet.timeControl = setTimeout(function () {
                    libraryBase.jspreadsheet.current!.search((e.target as HTMLInputElement).value);
                }, 200);
            }
        }
    }
};

export const wheelControls = function (this: Worksheet, e?: WheelEvent) {
    const obj = this;

    if (obj.options.lazyLoading == true) {
        if (libraryBase.jspreadsheet.timeControlLoading == null) {
            libraryBase.jspreadsheet.timeControlLoading = setTimeout(function () {
                if (obj.content.scrollTop + obj.content.clientHeight >= obj.content.scrollHeight - 10) {
                    if (loadDown.call(obj)) {
                        if (obj.content.scrollTop + obj.content.clientHeight > obj.content.scrollHeight - 10) {
                            obj.content.scrollTop = obj.content.scrollTop - obj.content.clientHeight;
                        }
                        updateCornerPosition.call(obj);
                    }
                } else if (obj.content.scrollTop <= obj.content.clientHeight) {
                    if (loadUp.call(obj)) {
                        if (obj.content.scrollTop < 10) {
                            obj.content.scrollTop = obj.content.scrollTop + obj.content.clientHeight;
                        }
                        updateCornerPosition.call(obj);
                    }
                }

                libraryBase.jspreadsheet.timeControlLoading = null;
            }, 100);
        }
    }
};

let scrollLeft = 0;

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

export const scrollControls = function (this: Worksheet, e: Event) {
    const obj = this;

    wheelControls.call(obj);

    if (obj.options.freezeColumns! > 0 && obj.content.scrollLeft != scrollLeft) {
        scrollLeft = obj.content.scrollLeft;
        updateFreezePosition.call(obj);
    }

    // Close a floating editor. The text editor is inside the cell and scrolls with it.
    if (obj.options.lazyLoading == true || obj.options.tableOverflow == true) {
        if (obj.edition && hasFloatingEditor.call(obj, parseInt(obj.edition[2])) && (e.target as HTMLElement).className.substr(0, 9) != 'jdropdown') {
            closeEditor.call(obj, obj.edition[0], true);
        }
    }
};

/**
 * Check if the editor of a column floats over the page. It does not move with the cell.
 */
const hasFloatingEditor = function (this: Worksheet, x: number) {
    const obj = this;

    const column = obj.options.columns && obj.options.columns[x];
    if (!column || !column.type) {
        return false;
    }

    return typeof column.type === 'object' || ['dropdown', 'calendar', 'color', 'html', 'image'].indexOf(column.type) >= 0;
};

export const setEvents = function (root: Document | HTMLElement | ShadowRoot) {
    destroyEvents(root);
    // The handlers read the target as an element, which the type of a listener does not promise
    root.addEventListener('mouseup', mouseUpControls as EventListener);
    root.addEventListener('mousedown', mouseDownControls as EventListener);
    root.addEventListener('mousemove', mouseMoveControls as EventListener);
    root.addEventListener('mouseover', mouseOverControls as EventListener);
    root.addEventListener('dblclick', doubleClickControls as EventListener);
    root.addEventListener('paste', pasteControls as EventListener);
    root.addEventListener('contextmenu', contextMenuControls as EventListener);
    root.addEventListener('touchstart', touchStartControls as EventListener);
    root.addEventListener('touchend', touchEndControls as EventListener);
    root.addEventListener('touchcancel', touchEndControls as EventListener);
    root.addEventListener('touchmove', touchEndControls as EventListener);
    document.addEventListener('keydown', keyDownControls as EventListener);
};

export const destroyEvents = function (root: Document | HTMLElement | ShadowRoot) {
    // The handlers read the target as an element, which the type of a listener does not promise
    root.removeEventListener('mouseup', mouseUpControls as EventListener);
    root.removeEventListener('mousedown', mouseDownControls as EventListener);
    root.removeEventListener('mousemove', mouseMoveControls as EventListener);
    root.removeEventListener('mouseover', mouseOverControls as EventListener);
    root.removeEventListener('dblclick', doubleClickControls as EventListener);
    root.removeEventListener('paste', pasteControls as EventListener);
    root.removeEventListener('contextmenu', contextMenuControls as EventListener);
    root.removeEventListener('touchstart', touchStartControls as EventListener);
    root.removeEventListener('touchend', touchEndControls as EventListener);
    root.removeEventListener('touchcancel', touchEndControls as EventListener);
    root.removeEventListener('touchmove', touchEndControls as EventListener);
    document.removeEventListener('keydown', keyDownControls as EventListener);
};
