import jSuites from 'jsuites';

import dispatch from './dispatch.js';
import { updateCell } from './render.js';
import { getCellColumnIndex, getMask, isFormula } from './values.js';
import { getColumnNameFromId } from './internalHelpers.js';
import { getLabelText, restoreActiveCell } from './accessibility.js';
import type { Worksheet, WorksheetMethod } from '../types/internal';
import type { CellChange } from '../jspreadsheet';
import type { MaskOptions } from '../types/jsuites';
import type { CalendarOptions, ColorOptions, DropdownOptions } from '../types/jsuites-module';

/**
 * The editor in a cell that is being edited, read as an input because the text editors are inputs or text areas.
 * jsuites keeps its control on the element, and a masked input keeps its mask options.
 * The typings of the calendar and the color picker say that close returns nothing, but it returns the value.
 */
type EditorElement = HTMLInputElement & {
    dropdown: { close(ignoreEvents: boolean): void; getValue(asArray: true): string[]; getValue(): string };
    calendar: { close(ignoreEvents: boolean): string };
    color: { close(ignoreEvents: boolean): string };
    mask?: MaskOptions;
};

/**
 * Open the editor
 *
 * @param object cell
 * @return void
 */
export const openEditor: WorksheetMethod<'openEditor'> = function (cell, empty, e) {
    const obj = this;

    // Get cell position. The filter of a dropdown gets the text of data-x and data-y.
    const yText = cell.getAttribute('data-y')!;
    const xText = cell.getAttribute('data-x')!;
    const y = parseInt(yText);
    const x = parseInt(xText);

    // Column options
    const column = obj.options.columns && obj.options.columns[getCellColumnIndex.call(obj, x, y)];

    // On edition start
    dispatch.call(obj, 'oneditionstart', obj, cell, x, y);

    // Accessible name of the editor
    const label = getLabelText.call(obj.parent, 'editCell', getColumnNameFromId([x, y]));

    // Overflow
    if (x > 0) {
        obj.records[y][x - 1].element.style.overflow = 'hidden';
    }

    // Create editor
    const createEditor = function <K extends keyof HTMLElementTagNameMap>(type: K) {
        // Cell information
        const info = cell.getBoundingClientRect();

        // Create dropdown
        const editor = document.createElement(type);
        editor.style.width = info.width + 'px';
        editor.style.height = info.height - 2 + 'px';
        editor.style.minHeight = info.height - 2 + 'px';

        // Accessible name of a text editor
        if (type === 'input' || type === 'textarea') {
            editor.setAttribute('aria-label', label);
        }

        // The editor is the focus target, and it has its own name
        cell.removeAttribute('tabindex');
        cell.removeAttribute('aria-label');

        // Edit cell
        cell.classList.add('editor');
        cell.innerHTML = '';
        cell.appendChild(editor);

        return editor;
    };

    // Readonly
    if (cell.classList.contains('readonly') == true) {
        // Do nothing
    } else {
        // Holder
        obj.editor.start(obj.records[y][x].element, x, y);

        // If there is a custom editor for it
        if (column && typeof column.type === 'object') {
            if (typeof column.type.openEditor === 'function') {
                // The editor is the focus target, and it has its own name
                cell.removeAttribute('tabindex');
                cell.removeAttribute('aria-label');

                // Custom editors
                column.type.openEditor(cell, obj.options.data[y][x], x, y, obj, column, e);

                // On edition start
                dispatch.call(obj, 'oncreateeditor', obj, cell, x, y, null, column);
            } else {
                // A custom type without openEditor has no editor: do not keep edition open
                obj.editor.end();
            }
        } else {
            // Native functions
            if (column && column.type == 'hidden') {
                // A hidden column has no editor: do not keep edition open
                obj.editor.end();
            } else if (column && (column.type == 'checkbox' || column.type == 'radio')) {
                // Get value
                const value = (cell.children[0] as HTMLInputElement).checked ? false : true;
                // Toogle value
                obj.setValue(cell, value);
                // Do not keep edition open
                obj.editor.end();
            } else if (column && column.type == 'dropdown') {
                // Get current value: the id of an item, or the ids of the items with multiple
                let value = obj.options.data[y][x] as string | number | string[];
                if (column.multiple && !Array.isArray(value)) {
                    // A number has no split
                    value = ('' + value).split(column.delimiter || ';');
                }

                // Create dropdown
                let source;

                if (typeof column.filter == 'function') {
                    source = column.filter(obj.element, cell, xText, yText, column.source);
                } else {
                    source = column.source;
                }

                // Do not change the original source
                const data = [];
                if (source) {
                    for (let j = 0; j < source.length; j++) {
                        data.push(source[j]);
                    }
                }

                // Create editor
                const editor = createEditor('div');

                // On edition start
                dispatch.call(obj, 'oncreateeditor', obj, cell, x, y, null, column);

                const options: DropdownOptions = {
                    data: data,
                    multiple: column.multiple ? true : false,
                    autocomplete: column.autocomplete ? true : false,
                    opened: true,
                    value: value,
                    width: '100%',
                    position: obj.options.tableOverflow == true || obj.parent.config.fullscreen == true ? true : false,
                    onclose: function () {
                        closeEditor.call(obj, cell, true);
                    },
                };
                if (column.options && column.options.type) {
                    options.type = column.options.type;
                }
                const dropdown = jSuites.dropdown(editor, options);
                if (dropdown && dropdown.header) {
                    dropdown.header.setAttribute('aria-label', label);
                    // The editor is the focus target, as the other editors
                    if (!editor.contains(document.activeElement)) {
                        dropdown.header.focus({ preventScroll: true });
                    }
                }
            } else if (column && (column.type == 'calendar' || column.type == 'color')) {
                // Value
                const value = obj.options.data[y][x];
                // Create editor
                const editor = createEditor('input');

                dispatch.call(obj, 'oncreateeditor', obj, cell, x, y, null, column);

                // The input converts the value to text
                editor.value = value as string;

                // The options of the calendar or the color picker
                const options: CalendarOptions & ColorOptions = column.options ? { ...column.options } : {};

                if (obj.options.tableOverflow == true || obj.parent.config.fullscreen == true) {
                    options.position = true;
                }
                // The value of a calendar or a color cell is text
                options.value = obj.options.data[y][x] as string;
                options.opened = true;
                // The color picker gives its instance as the second argument, and the calendar nothing
                options.onclose = function (el: HTMLElement, value?: unknown) {
                    closeEditor.call(obj, cell, true);
                };
                // Current value
                if (column.type == 'color') {
                    jSuites.color(editor, options);

                    const rect = cell.getBoundingClientRect();

                    if (options.position) {
                        // The color picker adds its container after the input
                        ((editor.nextSibling as HTMLElement).children[1] as HTMLElement).style.top = rect.top + rect.height + 'px';
                        ((editor.nextSibling as HTMLElement).children[1] as HTMLElement).style.left = rect.left + 'px';
                    }
                } else {
                    if (!options.format) {
                        options.format = 'YYYY-MM-DD';
                    }

                    jSuites.calendar(editor, options);
                }
                // Focus on editor
                editor.focus();
            } else if (column && column.type == 'html') {
                const value = obj.options.data[y][x];
                // Create editor
                const editor = createEditor('div');

                dispatch.call(obj, 'oncreateeditor', obj, cell, x, y, null, column);

                editor.style.position = 'relative';
                const div = document.createElement('div');
                div.classList.add('jss_richtext');
                editor.appendChild(div);
                // The jsuites typings miss the editor property
                const richText = jSuites.editor(div, {
                    focus: true,
                    value: value as string,
                }) as ReturnType<typeof jSuites.editor> & { editor?: HTMLElement };
                // Accessible role and name of the rich text editor
                if (richText && richText.editor) {
                    richText.editor.setAttribute('role', 'textbox');
                    richText.editor.setAttribute('aria-multiline', 'true');
                    richText.editor.setAttribute('aria-label', label);
                }
                const rect = cell.getBoundingClientRect();
                const rectContent = div.getBoundingClientRect();
                if (window.innerHeight < rect.bottom + rectContent.height) {
                    div.style.top = rect.bottom - (rectContent.height + 2) + 'px';
                } else {
                    div.style.top = rect.top + 'px';
                }

                if (window.innerWidth < rect.left + rectContent.width) {
                    div.style.left = rect.right - (rectContent.width + 2) + 'px';
                } else {
                    div.style.left = rect.left + 'px';
                }
            } else if (column && column.type == 'image') {
                // Value
                const img = cell.children[0] as HTMLImageElement;
                // Create editor
                const editor = createEditor('div');

                dispatch.call(obj, 'oncreateeditor', obj, cell, x, y, null, column);

                editor.style.position = 'relative';
                const div = document.createElement('div');
                div.classList.add('jclose');
                if (img && img.src) {
                    div.appendChild(img);
                }
                editor.appendChild(div);
                jSuites.image(div, column);
                const rect = cell.getBoundingClientRect();
                const rectContent = div.getBoundingClientRect();
                if (window.innerHeight < rect.bottom + rectContent.height) {
                    div.style.top = rect.top - (rectContent.height + 2) + 'px';
                } else {
                    div.style.top = rect.top + 'px';
                }

                div.style.left = rect.left + 'px';
            } else {
                // Value
                const value = empty == true ? '' : obj.options.data[y][x];

                // Basic editor
                let editor;

                if ((!column || column.wordWrap != false) && (obj.options.wordWrap == true || (column && column.wordWrap == true))) {
                    editor = createEditor('textarea');
                } else {
                    editor = createEditor('input');
                }

                dispatch.call(obj, 'oncreateeditor', obj, cell, x, y, null, column);

                editor.focus();
                // The input converts the value to text
                editor.value = value as string;

                // Column options
                const options = column;

                // Apply format when is not a formula
                if (!isFormula(value)) {
                    if (options) {
                        // Format
                        const opt = getMask(options);

                        if (opt) {
                            // Masking
                            if (!options.disabledMaskOnEdition) {
                                if (options.mask) {
                                    const m = options.mask.split(';');
                                    editor.setAttribute('data-mask', m[0]);
                                } else if (options.locale) {
                                    editor.setAttribute('data-locale', options.locale);
                                }
                            }
                            // Input
                            opt.input = editor;
                            // Configuration
                            (editor as EditorElement).mask = opt;
                            // Do not treat the decimals
                            // The jsuites typings miss render and extract
                            jSuites.mask.render(value, opt, false);
                        }
                    }
                }

                editor.onblur = function () {
                    closeEditor.call(obj, cell, true);
                };
                editor.scrollLeft = editor.scrollWidth;
            }
        }
    }
};

/**
 * Close the editor and save the information
 *
 * @param object cell
 * @param boolean save
 * @return void
 */
export const closeEditor: WorksheetMethod<'closeEditor'> = function (cell, save) {
    const obj = this;

    // Only the cell that is being edited has an editor to close
    if (!obj.editor.isOpen() || obj.editor.cell !== cell) {
        return;
    }

    const x = parseInt(cell.getAttribute('data-x'));
    const y = parseInt(cell.getAttribute('data-y'));

    // Column options
    const column = obj.options.columns && obj.options.columns[getCellColumnIndex.call(obj, x, y)];

    // The editor element has the jsuites control as a property
    let value;

    // Get cell properties
    if (save == true) {
        // If custom editor
        if (column && typeof column.type === 'object') {
            if (typeof column.type.closeEditor === 'function') {
                // Custom editor
                value = column.type.closeEditor(cell, save, parseInt(x), parseInt(y), obj, column);
            } else {
                // A custom type without closeEditor gives no new value: the cell keeps its value
                value = obj.options.data[y][x];
            }
        } else {
            // Native functions
            if (column && (column.type == 'checkbox' || column.type == 'radio' || column.type == 'hidden')) {
                // Do nothing
            } else if (column && column.type == 'dropdown') {
                const dropdown = (cell.children[0] as EditorElement).dropdown;
                const delimiter = column.delimiter;
                dropdown.close(true);
                value = delimiter && column.multiple ? dropdown.getValue(true).join(delimiter) : dropdown.getValue();
            } else if (column && column.type == 'calendar') {
                value = (cell.children[0] as EditorElement).calendar.close(true);
            } else if (column && column.type == 'color') {
                value = (cell.children[0] as EditorElement).color.close(true);
            } else if (column && column.type == 'html') {
                // The rich text editor keeps its jsuites control on its element
                value = (cell.children[0].children[0] as HTMLElement & { editor: { getData(): string } }).editor.getData();
            } else if (column && column.type == 'image') {
                const img = cell.children[0].children[0].children[0] as HTMLImageElement | undefined;
                value = img && img.tagName == 'IMG' ? img.src : '';
            } else if (column && column.type == 'numeric') {
                value = (cell.children[0] as EditorElement).value;
                if (('' + value).substr(0, 1) != '=') {
                    if (value == '') {
                        value = column.allowEmpty ? '' : 0;
                    }
                }
                (cell.children[0] as EditorElement).onblur = null;
            } else {
                value = (cell.children[0] as EditorElement).value;
                (cell.children[0] as EditorElement).onblur = null;

                // Column options
                const options = column;

                if (options) {
                    // Format
                    const opt = getMask(options);
                    if (opt) {
                        // Keep numeric in the raw data
                        if (value !== '' && !isFormula(value) && typeof value !== 'number') {
                            const t = jSuites.mask.extract(value, opt, true);
                            if (t && t.value !== '') {
                                value = t.value;
                            }
                        }
                    }
                }
            }
        }

        // Ignore changes if the value is the same
        if (obj.options.data[y][x] == value) {
            cell.innerHTML = obj.editor.html;
        } else {
            obj.setValue(cell, value);
        }
    } else {
        if (column && typeof column.type === 'object') {
            // Custom editor
            if (typeof column.type.closeEditor === 'function') {
                column.type.closeEditor(cell, save, parseInt(x), parseInt(y), obj, column);
            }
        } else {
            if (column && column.type == 'dropdown') {
                (cell.children[0] as EditorElement).dropdown.close(true);
            } else if (column && column.type == 'calendar') {
                (cell.children[0] as EditorElement).calendar.close(true);
            } else if (column && column.type == 'color') {
                (cell.children[0] as EditorElement).color.close(true);
            } else {
                (cell.children[0] as EditorElement).onblur = null;
            }
        }

        // Restore value
        cell.innerHTML = obj.editor.html;
    }

    // On edition end
    dispatch.call(obj, 'oneditionend', obj, cell, x, y, value, save);

    // Remove editor class
    cell.classList.remove('editor');

    // Finish edition
    obj.editor.end();

    // The cell is a focus target again
    restoreActiveCell.call(obj, cell);
};

/**
 * Toogle
 */
export const setCheckRadioValue = function (this: Worksheet) {
    const obj = this;

    const records = [];
    const keys = Object.keys(obj.selection.highlighted);
    for (let i = 0; i < keys.length; i++) {
        const x = parseInt(obj.selection.highlighted[i].element.getAttribute('data-x'));
        const y = parseInt(obj.selection.highlighted[i].element.getAttribute('data-y'));

        const column = obj.options.columns[getCellColumnIndex.call(obj, x, y)];
        if (column.type == 'checkbox' || column.type == 'radio') {
            // Update cell
            records.push(updateCell.call(obj, x, y, !obj.options.data[y][x]));
        }
    }

    if (records.length) {
        // Update history
        obj.history.add({
            action: 'setValue',
            records: records,
            selection: obj.selection.range && [...obj.selection.range],
        });

        // On after changes
        const onafterchangesRecords = records.map(function (record) {
            return {
                x: record.x,
                y: record.y,
                value: record.value,
                oldValue: record.oldValue,
            };
        });

        // The coordinates can be text (issue #52), and a read-only cell has no value (issue #58)
        dispatch.call(obj, 'onafterchanges', obj, onafterchangesRecords as unknown as CellChange[]);
    }
};
