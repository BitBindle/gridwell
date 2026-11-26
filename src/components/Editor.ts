import { getMask, isFormula } from '../utils/values.js';
import { extract, toEditor } from '../utils/mask.js';
import { getColumnNameFromId } from '../utils/internalHelpers.js';
import { ColorPicker, type ColorPickerOptions } from '../utils/color.js';
import { Dropdown } from '../utils/dropdown.js';
import { Calendar, type CalendarOptions } from '../utils/calendar.js';
import { createRichText, getRichText } from '../utils/richText.js';
import { ImageUpload, type ImageUploadElement } from '../utils/imageUpload.js';
import type { Worksheet } from '../types/internal';
import type jspreadsheet from '../jspreadsheet';

/**
 * The editor in a cell that is being edited, read as an input because the text editors are inputs or text areas.
 * The dropdown, the calendar and the color picker keep their control on the element.
 */
type EditorElement = HTMLInputElement & {
    dropdown: Dropdown;
    calendar: Calendar;
    color: ColorPicker;
};

/** The options of a column, as the worksheet keeps them. */
type ColumnOptions = Worksheet['options']['columns'][number];

/**
 * Put a new editor of the given tag in a cell, with the size of the cell. The editor is the focus target in place of the
 * cell, and a text editor gets the accessible name.
 */
function createEditor<K extends keyof HTMLElementTagNameMap>(cell: HTMLTableCellElement, type: K, label: string) {
    // Cell information
    const info = cell.getBoundingClientRect();

    // Create dropdown
    const editor = document.createElement(type);
    editor.style.width = `${info.width}px`;
    editor.style.height = `${info.height - 2}px`;
    editor.style.minHeight = `${info.height - 2}px`;

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
}

/**
 * The cell editor of a worksheet: which cell is being edited, where it is, and its HTML before the editor opened. It
 * opens the editor of the column type, writes the value to the cell when the editor closes, and toggles checkboxes.
 *
 * Each worksheet has one, as worksheet.editor. openEditor and closeEditor call it.
 */
export class Editor {
    readonly worksheet: Worksheet;

    /** The cell that is being edited, or null when no editor is open. */
    cell: HTMLTableCellElement | null = null;

    /** The HTML of the cell before the editor opened. Closing restores it when the value does not change. */
    html = '';

    /** The column of the cell, or -1 when no editor is open. */
    x = -1;

    /** The row of the cell, or -1 when no editor is open. */
    y = -1;

    /** The text of the editValue hook that the text editor opened with, or null. */
    text: string | null = null;

    constructor(worksheet: Worksheet) {
        this.worksheet = worksheet;
    }

    /** Whether an editor is open. */
    isOpen(): this is Editor & { cell: HTMLTableCellElement } {
        return this.cell !== null;
    }

    /** Keep the cell that an editor opens on. */
    start(cell: HTMLTableCellElement, x: number, y: number) {
        this.cell = cell;
        this.html = cell.innerHTML;
        this.x = x;
        this.y = y;
        this.text = null;
    }

    /** Forget the cell when its editor closes, or when the column type has no editor to keep open. */
    end() {
        this.cell = null;
        this.html = '';
        this.x = -1;
        this.y = -1;
        this.text = null;
    }

    /**
     * The editor of a column with a custom type. A custom type without openEditor has no editor.
     */
    private openCustom(
        cell: HTMLTableCellElement,
        x: number,
        y: number,
        column: ColumnOptions & { type: jspreadsheet.CustomEditor },
        e: Parameters<jspreadsheet.WorksheetInstance['openEditor']>[2]
    ) {
        const worksheet = this.worksheet;

        if (typeof column.type.openEditor === 'function') {
            // The editor is the focus target, and it has its own name
            cell.removeAttribute('tabindex');
            cell.removeAttribute('aria-label');

            // Custom editors
            column.type.openEditor(cell, worksheet.options.data[y][x], x, y, worksheet, column, e);

            // On edition start
            worksheet.dispatch('oncreateeditor', worksheet, cell, x, y, null, column);
        } else {
            // A custom type without openEditor has no editor: do not keep edition open
            worksheet.editor.end();
        }
    }

    /**
     * A checkbox or a radio has no editor: opening it toggles the value.
     */
    private toggleCheckbox(cell: HTMLTableCellElement) {
        const worksheet = this.worksheet;

        // Get value
        const value = (cell.children[0] as HTMLInputElement).checked ? false : true;
        // Toogle value
        worksheet.setValue(cell, value);
        // Do not keep edition open
        worksheet.editor.end();
    }

    /**
     * The editor of a dropdown column: an open dropdown with the items of the source, or of the filter.
     * xText and yText are data-x and data-y of the cell, which the filter gets as text.
     */
    private openDropdown(cell: HTMLTableCellElement, x: number, y: number, xText: string, yText: string, column: ColumnOptions, label: string) {
        const worksheet = this.worksheet;

        // Get current value: the id of an item, or the ids of the items with multiple
        let value = worksheet.options.data[y][x] as string | number | string[];
        if (column.multiple && !Array.isArray(value)) {
            // A number has no split
            value = `${value}`.split(column.delimiter || ';');
        }

        // Create dropdown
        let source;

        if (typeof column.filter === 'function') {
            source = column.filter(worksheet.element, cell, xText, yText, column.source);
        } else {
            source = column.source;
        }

        // Do not change the original source
        const data = [];
        if (source) {
            for (const option of source) {
                data.push(option);
            }
        }

        // Create editor
        const editor = createEditor(cell, 'div', label);

        // On edition start
        worksheet.dispatch('oncreateeditor', worksheet, cell, x, y, null, column);

        // The dropdown takes the focus. Closing it, also with a choice, saves the value.
        const dropdown = new Dropdown(editor, {
            data,
            multiple: !!column.multiple,
            autocomplete: !!column.autocomplete,
            value,
            width: '100%',
            position: worksheet.options.tableOverflow === true || worksheet.parent.config.fullscreen === true,
            onclose: () => {
                worksheet.editor.close(cell, true);
            },
        });
        dropdown.header.setAttribute('aria-label', label);
        dropdown.open();
    }

    /**
     * The editor of a calendar or a color column: an input with an open calendar, or a color picker.
     */
    private openPicker(cell: HTMLTableCellElement, x: number, y: number, column: ColumnOptions, label: string) {
        const worksheet = this.worksheet;

        // Value
        const value = worksheet.options.data[y][x];
        // Create editor
        const editor = createEditor(cell, 'input', label);

        worksheet.dispatch('oncreateeditor', worksheet, cell, x, y, null, column);

        // The input converts the value to text
        editor.value = value as string;

        const fixed = worksheet.options.tableOverflow === true || worksheet.parent.config.fullscreen === true;

        if (column.type === 'color') {
            // The value of a color cell is text. Closing the picker, also with a color, saves it.
            const picker = new ColorPicker(editor, {
                ...(column.options as ColorPickerOptions | undefined),
                ...(fixed ? { position: true } : {}),
                value: worksheet.options.data[y][x] as string,
                onclose: () => {
                    worksheet.editor.close(cell, true);
                },
            });

            if (fixed) {
                const rect = cell.getBoundingClientRect();
                picker.content.style.top = `${rect.top + rect.height}px`;
                picker.content.style.left = `${rect.left}px`;
            }

            // The picker takes the focus
            picker.open();
            return;
        }

        // The value of a calendar cell is text. Closing the calendar, also with a day, saves it.
        const calendar = new Calendar(editor, {
            ...(column.options as CalendarOptions | undefined),
            ...(fixed ? { position: true } : {}),
            value: worksheet.options.data[y][x] as string,
            onclose: () => {
                worksheet.editor.close(cell, true);
            },
        });

        // The input has the focus
        calendar.open();
    }

    /**
     * The editor of an html column: a rich text editor over the cell, kept inside the window.
     */
    private openHtml(cell: HTMLTableCellElement, x: number, y: number, column: ColumnOptions, label: string) {
        const worksheet = this.worksheet;

        const value = worksheet.options.data[y][x];
        // Create editor
        const editor = createEditor(cell, 'div', label);

        worksheet.dispatch('oncreateeditor', worksheet, cell, x, y, null, column);

        editor.style.position = 'relative';
        const div = document.createElement('div');
        div.classList.add('jss_richtext');
        editor.appendChild(div);
        createRichText(div, value == null ? '' : `${value}`, label);

        const rect = cell.getBoundingClientRect();
        const rectContent = div.getBoundingClientRect();
        if (window.innerHeight < rect.bottom + rectContent.height) {
            div.style.top = `${rect.bottom - (rectContent.height + 2)}px`;
        } else {
            div.style.top = `${rect.top}px`;
        }

        if (window.innerWidth < rect.left + rectContent.width) {
            div.style.left = `${rect.right - (rectContent.width + 2)}px`;
        } else {
            div.style.left = `${rect.left}px`;
        }
    }

    /**
     * The editor of an image column: an image upload over the cell, with the image of the cell.
     */
    private openImage(cell: HTMLTableCellElement, x: number, y: number, column: ColumnOptions, label: string) {
        const worksheet = this.worksheet;

        // Value
        const img = cell.children[0] as HTMLImageElement | undefined;
        // Create editor
        const editor = createEditor(cell, 'div', label);

        worksheet.dispatch('oncreateeditor', worksheet, cell, x, y, null, column);

        editor.style.position = 'relative';
        const div = document.createElement('div');
        editor.appendChild(div);
        const upload = new ImageUpload(div, {
            value: img?.tagName === 'IMG' ? img.src : '',
            label,
            ondone: () => {
                worksheet.editor.close(cell, true);
            },
        });

        const rect = cell.getBoundingClientRect();
        const rectContent = div.getBoundingClientRect();
        if (window.innerHeight < rect.bottom + rectContent.height) {
            div.style.top = `${rect.top - (rectContent.height + 2)}px`;
        } else {
            div.style.top = `${rect.top}px`;
        }

        div.style.left = `${rect.left}px`;
        upload.focus();
    }

    /**
     * The text editor: an input, or a text area when the column wraps its text, with the mask of the column. empty opens it
     * without the value of the cell.
     */
    private openText(
        cell: HTMLTableCellElement,
        x: number,
        y: number,
        // The type of the column options, as openEditor reads them. The code still checks for a column.
        column: ColumnOptions,
        label: string,
        empty: boolean | undefined
    ) {
        const worksheet = this.worksheet;

        // Value
        const value = empty === true ? '' : worksheet.options.data[y][x];
        const editValue = worksheet.parent.config.editValue;
        const text = empty !== true && typeof editValue === 'function' ? editValue(worksheet, x, y, value) : undefined;

        // Basic editor
        let editor;

        if (column?.wordWrap !== false && (worksheet.options.wordWrap === true || column?.wordWrap === true)) {
            editor = createEditor(cell, 'textarea', label);
        } else {
            editor = createEditor(cell, 'input', label);
        }

        worksheet.dispatch('oncreateeditor', worksheet, cell, x, y, null, column);

        editor.focus();
        // A value with a mask shows as a plain number, such as 1234.5 for "$ 1,234.50", or as a date in the format of the
        // mask. The input converts the value to text.
        // The text of the editValue hook shows as it is.
        const mask = column && !isFormula(value) ? getMask(column) : null;
        if (text !== undefined) {
            // A hook in JavaScript can return null
            editor.value = text == null ? '' : `${text}`;
            this.text = editor.value;
        } else {
            editor.value = mask ? toEditor(value, mask) : (value as string);
        }

        editor.onblur = () => {
            worksheet.editor.close(cell, true);
        };
        editor.scrollLeft = editor.scrollWidth;
    }

    /**
     * Open the editor of a cell: the editor of its column type, or of the custom type. A read-only cell, a hidden column and
     * a custom type without openEditor have no editor, and a checkbox or a radio toggles.
     * @param cell - The td of the cell.
     * @param empty - Open the editor without the value of the cell.
     * @param e - The event that opens the editor.
     */
    open(cell: HTMLTableCellElement, empty?: boolean, e?: KeyboardEvent | MouseEvent | TouchEvent) {
        const worksheet = this.worksheet;

        // Get cell position. The filter of a dropdown gets the text of data-x and data-y.
        const yText = cell.getAttribute('data-y')!;
        const xText = cell.getAttribute('data-x')!;
        const y = parseInt(yText);
        const x = parseInt(xText);

        // Column options
        const column = worksheet.options.columns?.[worksheet.cells.getColumnIndex(x, y)];

        // On edition start
        worksheet.dispatch('oneditionstart', worksheet, cell, x, y);

        // Accessible name of the editor
        const label = worksheet.accessibility.getLabelText('editCell', getColumnNameFromId([x, y]));

        // Overflow
        if (x > 0) {
            worksheet.records[y][x - 1].element.style.overflow = 'hidden';
        }

        // Readonly
        if (cell.classList.contains('readonly') === true) {
            // Do nothing
        } else {
            // Holder
            worksheet.editor.start(worksheet.records[y][x].element, x, y);

            // If there is a custom editor for it
            if (column && typeof column.type === 'object') {
                // The test of typeof narrows the type, not the column
                this.openCustom(cell, x, y, column as ColumnOptions & { type: jspreadsheet.CustomEditor }, e);
            } else {
                // Native functions
                if (column?.type === 'hidden') {
                    // A hidden column has no editor: do not keep edition open
                    worksheet.editor.end();
                } else if (column && (column.type === 'checkbox' || column.type === 'radio')) {
                    this.toggleCheckbox(cell);
                } else if (column?.type === 'dropdown') {
                    this.openDropdown(cell, x, y, xText, yText, column, label);
                } else if (column && (column.type === 'calendar' || column.type === 'color')) {
                    this.openPicker(cell, x, y, column, label);
                } else if (column?.type === 'html') {
                    this.openHtml(cell, x, y, column, label);
                } else if (column?.type === 'image') {
                    this.openImage(cell, x, y, column, label);
                } else {
                    this.openText(cell, x, y, column, label, empty);
                }
            }
        }
    }

    /**
     * Close the editor when lazy loading removes its row. The application can take the edit in oneditordetach: when the
     * handler returns true, the editor closes without a save. Else the editor saves and closes, as before.
     */
    detach() {
        if (!this.isOpen()) {
            return;
        }
        const worksheet = this.worksheet;
        const cell = this.cell;
        // The editor input leaves the table, so its blur must not save it. A handler can move the focus.
        const input = cell.children[0] as HTMLElement | undefined;
        if (input) {
            input.onblur = null;
        }

        // The cell has its position now. The position of the open, kept in x and y, can be old after an insert above it.
        const x = parseInt(cell.getAttribute('data-x'));
        const y = parseInt(cell.getAttribute('data-y'));
        const taken = worksheet.dispatch('oneditordetach', worksheet, cell, x, y) === true;
        // The handler can close the editor itself
        if (this.isOpen() && this.cell === cell) {
            this.close(cell, !taken);
        }
    }

    /**
     * Close the editor of a cell.
     * @param cell - The td of the cell.
     * @param save - Save the value of the editor in the cell.
     */
    close(cell: HTMLTableCellElement, save: boolean) {
        const worksheet = this.worksheet;

        // Only the cell that is being edited has an editor to close
        if (!worksheet.editor.isOpen() || worksheet.editor.cell !== cell) {
            return;
        }

        const x = parseInt(cell.getAttribute('data-x'));
        const y = parseInt(cell.getAttribute('data-y'));

        // Column options
        const column = worksheet.options.columns?.[worksheet.cells.getColumnIndex(x, y)];

        let value;

        // Get cell properties
        if (save === true) {
            // If custom editor
            if (column && typeof column.type === 'object') {
                if (typeof column.type.closeEditor === 'function') {
                    // Custom editor
                    value = column.type.closeEditor(cell, save, parseInt(x), parseInt(y), worksheet, column);
                } else {
                    // A custom type without closeEditor gives no new value: the cell keeps its value
                    value = worksheet.options.data[y][x];
                }
            } else {
                // Native functions
                if (column && (column.type === 'checkbox' || column.type === 'radio' || column.type === 'hidden')) {
                    // Do nothing
                } else if (column?.type === 'dropdown') {
                    const dropdown = (cell.children[0] as EditorElement).dropdown;
                    const delimiter = column.delimiter;
                    dropdown.close(true);
                    value = delimiter && column.multiple ? dropdown.getValue(true).join(delimiter) : dropdown.getValue();
                } else if (column?.type === 'calendar') {
                    value = (cell.children[0] as EditorElement).calendar.close(true);
                } else if (column?.type === 'color') {
                    value = (cell.children[0] as EditorElement).color.close(true);
                } else if (column?.type === 'html') {
                    value = getRichText(cell.children[0].children[0] as HTMLElement);
                } else if (column?.type === 'image') {
                    value = (cell.children[0].children[0] as ImageUploadElement).upload!.getValue();
                } else {
                    value = (cell.children[0] as EditorElement).value;
                    (cell.children[0] as EditorElement).onblur = null;

                    if (worksheet.editor.text !== null && value === worksheet.editor.text) {
                        // The text of the editValue hook that the user did not change keeps the value of the cell
                        value = worksheet.options.data[y][x];
                    } else {
                        // A number that is typed with the symbols of the mask, such as "$ 1,234.50", is kept as a number.
                        // A text that is not a number is kept as it is typed.
                        const mask = column ? getMask(column) : null;
                        if (mask && value !== '' && !isFormula(value)) {
                            value = extract(value, mask);
                        }

                        if (column?.type === 'numeric' && value === '') {
                            value = column.allowEmpty ? '' : 0;
                        }
                    }
                }
            }

            // Ignore changes if the value is the same
            // eslint-disable-next-line eqeqeq -- a value equals its text, for example 1 and "1"
            if (worksheet.options.data[y][x] == value) {
                cell.innerHTML = worksheet.editor.html;
            } else {
                // A built-in editor goes first, so that a change that onbeforechange cancels leaves the cell as it was
                if (!(column && typeof column.type === 'object')) {
                    cell.innerHTML = worksheet.editor.html;
                }
                worksheet.setValue(cell, value);
            }
        } else {
            if (column && typeof column.type === 'object') {
                // Custom editor
                if (typeof column.type.closeEditor === 'function') {
                    column.type.closeEditor(cell, save, parseInt(x), parseInt(y), worksheet, column);
                }
            } else {
                if (column?.type === 'dropdown') {
                    (cell.children[0] as EditorElement).dropdown.close(true);
                } else if (column?.type === 'calendar') {
                    (cell.children[0] as EditorElement).calendar.close(true);
                } else if (column?.type === 'color') {
                    (cell.children[0] as EditorElement).color.close(true);
                } else {
                    (cell.children[0] as EditorElement).onblur = null;
                }
            }

            // Restore value
            cell.innerHTML = worksheet.editor.html;
        }

        // On edition end
        worksheet.dispatch('oneditionend', worksheet, cell, x, y, value, save);

        // Remove editor class
        cell.classList.remove('editor');

        // Finish edition
        worksheet.editor.end();

        // The cell is a focus target again
        worksheet.selection.restoreActiveCell(cell);
    }

    /**
     * Toogle
     */
    setCheckRadioValue() {
        const worksheet = this.worksheet;

        const records = [];
        for (const cell of worksheet.selection.highlighted) {
            const x = parseInt(cell.element.getAttribute('data-x'));
            const y = parseInt(cell.element.getAttribute('data-y'));

            const column = worksheet.options.columns[worksheet.cells.getColumnIndex(x, y)];
            if (column.type === 'checkbox' || column.type === 'radio') {
                // Update cell
                const record = worksheet.cells.update(x, y, !worksheet.options.data[y][x]);
                if (record) {
                    records.push(record);
                }
            }
        }

        if (records.length) {
            // Update history
            worksheet.history.add({
                action: 'setValue',
                records,
                selection: worksheet.selection.range && [...worksheet.selection.range],
            });

            // On after changes
            const onafterchangesRecords = records.map((record) => {
                return {
                    x: record.x,
                    y: record.y,
                    value: record.value,
                    oldValue: record.oldValue,
                };
            });

            worksheet.dispatch('onafterchanges', worksheet, onafterchangesRecords);
        }
    }
}
