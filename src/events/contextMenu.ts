import jSuites from 'jsuites';
import libraryBase from '../utils/libraryBase.js';
import { selectAll, updateSelectionFromCoords } from '../utils/selection.js';
import { copy, paste } from '../utils/copyPaste.js';
import version from '../utils/version.js';
import { getCellNameFromCoords } from '../utils/helpers.js';
import { getElementIndex, getRole } from './helpers.js';
import type { TargetEvent, Worksheet } from '../types/internal';

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
                    if (worksheet.selection.range) {
                        navigator.clipboard.readText().then(function (text) {
                            if (text) {
                                paste.call(worksheet, worksheet.selection.range![0], worksheet.selection.range![1], text);
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

export const contextMenuControls = function (e: TargetEvent<PointerEvent>) {
    e = e || window.event;

    const current = libraryBase.jspreadsheet.current;

    if (current) {
        const spreadsheet = current.parent;

        if (current.editor.isOpen()) {
            e.preventDefault();
        } else {
            spreadsheet.contextMenu.contextmenu.close();

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
                    !current.selection.range ||
                    Number(x) < parseInt(current.selection.range[0]) ||
                    Number(x) > parseInt(current.selection.range[2]) ||
                    Number(y) < parseInt(current.selection.range[1]) ||
                    Number(y) > parseInt(current.selection.range[3])
                ) {
                    updateSelectionFromCoords.call(current, parseInt(x), parseInt(y), parseInt(x), parseInt(y), e);
                }
            } else if (role === 'row' || role === 'header') {
                if (role === 'row') {
                    y = e.target.getAttribute('data-y');
                } else {
                    x = e.target.getAttribute('data-x');
                }

                if (
                    !current.selection.range ||
                    Number(x) < parseInt(current.selection.range[0]) ||
                    Number(x) > parseInt(current.selection.range[2]) ||
                    Number(y) < parseInt(current.selection.range[1]) ||
                    Number(y) > parseInt(current.selection.range[3])
                ) {
                    // A null index selects the whole row or column
                    const column = x == null ? null : parseInt(x);
                    const row = y == null ? null : parseInt(y);
                    updateSelectionFromCoords.call(current, column, row, column, row, e);
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
                    !current.selection.range ||
                    columns[0] != parseInt(current.selection.range[0]) ||
                    columns[columns.length - 1] != parseInt(current.selection.range[2]) ||
                    current.selection.range[1] != null ||
                    current.selection.range[3] != null
                ) {
                    updateSelectionFromCoords.call(current, columns[0], null, columns[columns.length - 1], null, e);
                }
            } else if (role === 'select-all') {
                selectAll.call(current);
            } else if (role === 'tabs') {
                x = getElementIndex(e.target);
            } else if (role === 'footer') {
                x = getElementIndex(e.target) - 1;
                y = getElementIndex(e.target.parentElement!);
            }

            // Table found
            let items = defaultContextMenu(current, parseInt(x), parseInt(y), role);

            if (typeof spreadsheet.config.contextMenu === 'function') {
                const result = spreadsheet.config.contextMenu(current, x, y, e, items, role, x, y);

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
                            current,
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
};
