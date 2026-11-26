import jSuites from 'jsuites';
import { getMask, isFormula } from '../utils/values.js';
import { getColumnNameFromId } from '../utils/internalHelpers.js';
import type { Worksheet } from '../types/internal';
import type jspreadsheet from '../jspreadsheet';
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
    }

    /** Forget the cell when its editor closes, or when the column type has no editor to keep open. */
    end() {
        this.cell = null;
        this.html = '';
        this.x = -1;
        this.y = -1;
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
     * The editor of a dropdown column: an open jSuites dropdown with the items of the source, or of the filter.
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

        const options: DropdownOptions = {
            data,
            multiple: column.multiple ? true : false,
            autocomplete: column.autocomplete ? true : false,
            opened: true,
            value,
            width: '100%',
            position: worksheet.options.tableOverflow === true || worksheet.parent.config.fullscreen === true ? true : false,
            onclose() {
                worksheet.editor.close(cell, true);
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
    }

    /**
     * The editor of a calendar or a color column: an input with an open jSuites calendar or color picker.
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

        // The options of the calendar or the color picker
        const options: CalendarOptions & ColorOptions = column.options ? { ...column.options } : {};

        if (worksheet.options.tableOverflow === true || worksheet.parent.config.fullscreen === true) {
            options.position = true;
        }
        // The value of a calendar or a color cell is text
        options.value = worksheet.options.data[y][x] as string;
        options.opened = true;
        // The color picker gives its instance as the second argument, and the calendar nothing
        options.onclose = (el: HTMLElement, value?: unknown) => {
            worksheet.editor.close(cell, true);
        };
        // Current value
        if (column.type === 'color') {
            jSuites.color(editor, options);

            const rect = cell.getBoundingClientRect();

            if (options.position) {
                // The color picker adds its container after the input
                ((editor.nextSibling as HTMLElement).children[1] as HTMLElement).style.top = `${rect.top + rect.height}px`;
                ((editor.nextSibling as HTMLElement).children[1] as HTMLElement).style.left = `${rect.left}px`;
            }
        } else {
            if (!options.format) {
                options.format = 'YYYY-MM-DD';
            }

            jSuites.calendar(editor, options);
        }
        // Focus on editor
        editor.focus();
    }

    /**
     * The editor of an html column: a jSuites rich text editor over the cell, kept inside the window.
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
     * The editor of an image column: a jSuites image upload over the cell, with the image of the cell.
     */
    private openImage(cell: HTMLTableCellElement, x: number, y: number, column: ColumnOptions, label: string) {
        const worksheet = this.worksheet;

        // Value
        const img = cell.children[0] as HTMLImageElement;
        // Create editor
        const editor = createEditor(cell, 'div', label);

        worksheet.dispatch('oncreateeditor', worksheet, cell, x, y, null, column);

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
            div.style.top = `${rect.top - (rectContent.height + 2)}px`;
        } else {
            div.style.top = `${rect.top}px`;
        }

        div.style.left = `${rect.left}px`;
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

        // Basic editor
        let editor;

        if ((!column || column.wordWrap !== false) && (worksheet.options.wordWrap === true || (column && column.wordWrap === true))) {
            editor = createEditor(cell, 'textarea', label);
        } else {
            editor = createEditor(cell, 'input', label);
        }

        worksheet.dispatch('oncreateeditor', worksheet, cell, x, y, null, column);

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

        editor.onblur = () => {
            worksheet.editor.close(cell, true);
        };
        editor.scrollLeft = editor.scrollWidth;
    }

    /**
     * Open the editor of a cell: the editor of its column type, or of the custom type. A read-only cell, a hidden column and
     * a custom type without openEditor have no editor, and a checkbox or a radio toggles.
     *
     * @param object cell
     * @return void
     */
    open(cell: HTMLTableCellElement, empty?: boolean, e?: KeyboardEvent | MouseEvent | TouchEvent) {
        const worksheet = this.worksheet;

        // Get cell position. The filter of a dropdown gets the text of data-x and data-y.
        const yText = cell.getAttribute('data-y')!;
        const xText = cell.getAttribute('data-x')!;
        const y = parseInt(yText);
        const x = parseInt(xText);

        // Column options
        const column = worksheet.options.columns && worksheet.options.columns[worksheet.cells.getColumnIndex(x, y)];

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
                if (column && column.type === 'hidden') {
                    // A hidden column has no editor: do not keep edition open
                    worksheet.editor.end();
                } else if (column && (column.type === 'checkbox' || column.type === 'radio')) {
                    this.toggleCheckbox(cell);
                } else if (column && column.type === 'dropdown') {
                    this.openDropdown(cell, x, y, xText, yText, column, label);
                } else if (column && (column.type === 'calendar' || column.type === 'color')) {
                    this.openPicker(cell, x, y, column, label);
                } else if (column && column.type === 'html') {
                    this.openHtml(cell, x, y, column, label);
                } else if (column && column.type === 'image') {
                    this.openImage(cell, x, y, column, label);
                } else {
                    this.openText(cell, x, y, column, label, empty);
                }
            }
        }
    }

    /**
     * Close the editor and save the information
     *
     * @param object cell
     * @param boolean save
     * @return void
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
        const column = worksheet.options.columns && worksheet.options.columns[worksheet.cells.getColumnIndex(x, y)];

        // The editor element has the jsuites control as a property
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
                } else if (column && column.type === 'dropdown') {
                    const dropdown = (cell.children[0] as EditorElement).dropdown;
                    const delimiter = column.delimiter;
                    dropdown.close(true);
                    value = delimiter && column.multiple ? dropdown.getValue(true).join(delimiter) : dropdown.getValue();
                } else if (column && column.type === 'calendar') {
                    value = (cell.children[0] as EditorElement).calendar.close(true);
                } else if (column && column.type === 'color') {
                    value = (cell.children[0] as EditorElement).color.close(true);
                } else if (column && column.type === 'html') {
                    // The rich text editor keeps its jsuites control on its element
                    value = (cell.children[0].children[0] as HTMLElement & { editor: { getData(): string } }).editor.getData();
                } else if (column && column.type === 'image') {
                    const img = cell.children[0].children[0].children[0] as HTMLImageElement | undefined;
                    value = img && img.tagName === 'IMG' ? img.src : '';
                } else if (column && column.type === 'numeric') {
                    value = (cell.children[0] as EditorElement).value;
                    if (`${value}`.substr(0, 1) !== '=') {
                        if (value === '') {
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
            // eslint-disable-next-line eqeqeq -- a value equals its text, for example 1 and "1"
            if (worksheet.options.data[y][x] == value) {
                cell.innerHTML = worksheet.editor.html;
            } else {
                worksheet.setValue(cell, value);
            }
        } else {
            if (column && typeof column.type === 'object') {
                // Custom editor
                if (typeof column.type.closeEditor === 'function') {
                    column.type.closeEditor(cell, save, parseInt(x), parseInt(y), worksheet, column);
                }
            } else {
                if (column && column.type === 'dropdown') {
                    (cell.children[0] as EditorElement).dropdown.close(true);
                } else if (column && column.type === 'calendar') {
                    (cell.children[0] as EditorElement).calendar.close(true);
                } else if (column && column.type === 'color') {
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
                records.push(worksheet.cells.update(x, y, !worksheet.options.data[y][x]));
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

            // The coordinates can be text (issue #52), and a read-only cell has no value (issue #58)
            worksheet.dispatch('onafterchanges', worksheet, onafterchangesRecords as unknown as CellChange[]);
        }
    }
}
