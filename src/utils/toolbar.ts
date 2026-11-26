import jSuites from 'jsuites';
import { translate } from './dictionary.js';
import { getCellNameFromCoords } from './helpers.js';
import type jspreadsheet from '../jspreadsheet';
import type { Spreadsheet, Worksheet } from '../types/internal';
import type { Color } from 'jsuites';

/** The picker of the borders, which keeps the thickness, the color and the style that the user chose. */
type BorderPicker = jspreadsheet.PickerInstance & { thickness?: number; color?: string; style?: string };

/** The element of a picker, on which jsuites keeps the picker. */
type PickerElement = HTMLElement & { picker?: jspreadsheet.PickerInstance };

function setItemStatus(toolbarItem: HTMLElement, worksheet: Worksheet) {
    const disabled = worksheet.options.editable === false;

    toolbarItem.classList.toggle('jtoolbar-disabled', disabled);
    // A disabled item keeps the focus, as the WAI-ARIA toolbar pattern allows, and says that it is disabled
    toolbarItem.setAttribute('aria-disabled', disabled ? 'true' : 'false');
}

/**
 * The updateState of an item that changes cells: the item is disabled while the worksheet that the spreadsheet shows is
 * not editable.
 */
function getUpdateState(spreadsheet: Spreadsheet) {
    return (a: HTMLElement, b: object, toolbarItem: HTMLElement) => {
        setItemStatus(toolbarItem, spreadsheet.getWorksheetInstance());
    };
}

/**
 * Give the selected cells of the worksheet a style. setStyle toggles off a value that a cell has already, so a select
 * item, which sets the value the user chose, leaves those cells out, and only a button such as bold toggles.
 * @param property - A CSS property, such as 'font-size'.
 * @param value - Its value, such as 'large'.
 * @param toggle - Toggle the value off in the cells that have it.
 */
function setSelectedStyle(worksheet: Worksheet, property: string, value: string, toggle?: boolean) {
    const cells = worksheet.getSelected(true);
    if (cells) {
        const changed = toggle ? cells : cells.filter((cellName: string) => worksheet.getStyle(cellName, property) !== value);
        if (changed.length) {
            worksheet.setStyle(
                Object.fromEntries(
                    changed.map((cellName: string) => {
                        return [cellName, `${property}:${value}`];
                    })
                )
            );
        }
    }
}

/**
 * The selected range, from its first column and row to its last column and row.
 */
function getSelectedRange(range: [number, number, number, number]) {
    return [Math.min(range[0], range[2]), Math.min(range[1], range[3]), Math.max(range[0], range[2]), Math.max(range[1], range[3])];
}

/**
 * Undo, redo and save.
 */
function getHistoryItems(spreadsheet: Spreadsheet): jspreadsheet.ToolbarItem[] {
    return [
        {
            content: 'undo',
            tooltip: translate('Undo'),
            onclick() {
                const worksheet = spreadsheet.getWorksheetInstance();

                worksheet.undo();
            },
        },
        {
            content: 'redo',
            tooltip: translate('Redo'),
            onclick() {
                const worksheet = spreadsheet.getWorksheetInstance();

                worksheet.redo();
            },
        },
        {
            content: 'save',
            tooltip: translate('Save'),
            onclick() {
                const worksheet = spreadsheet.getWorksheetInstance();

                if (worksheet) {
                    worksheet.download();
                }
            },
        },
    ];
}

/**
 * The font, the size, the alignment, bold, the color of the text and the color of the background.
 */
function getFontItems(spreadsheet: Spreadsheet): jspreadsheet.ToolbarItem[] {
    return [
        {
            type: 'select',
            width: '120px',
            tooltip: translate('Font'),
            options: ['Default', 'Verdana', 'Arial', 'Courier New'],
            render(e: string) {
                return `<span style="font-family:${e}">${e}</span>`;
            },
            onchange(a: HTMLElement, b: object, c: string, d: string, e: string) {
                // e is the key of the option as text, and the first option is Default
                const value = e === '0' ? '' : d;

                setSelectedStyle(spreadsheet.getWorksheetInstance(), 'font-family', value);
            },
            updateState: getUpdateState(spreadsheet),
        },
        {
            type: 'select',
            width: '48px',
            content: 'format_size',
            tooltip: translate('Font size'),
            options: ['x-small', 'small', 'medium', 'large', 'x-large'],
            render(e: string) {
                return `<span style="font-size:${e}">${e}</span>`;
            },
            onchange(a: HTMLElement, b: object, c: string, value: string) {
                setSelectedStyle(spreadsheet.getWorksheetInstance(), 'font-size', value);
            },
            updateState: getUpdateState(spreadsheet),
        },
        {
            type: 'select',
            tooltip: translate('Text alignment'),
            options: ['left', 'center', 'right', 'justify'],
            render(e: string) {
                return `<i class="material-icons">format_align_${e}</i>`;
            },
            onchange(a: HTMLElement, b: object, c: string, value: string) {
                setSelectedStyle(spreadsheet.getWorksheetInstance(), 'text-align', value);
            },
            updateState: getUpdateState(spreadsheet),
        },
        {
            content: 'format_bold',
            tooltip: translate('Bold'),
            onclick() {
                setSelectedStyle(spreadsheet.getWorksheetInstance(), 'font-weight', 'bold', true);
            },
            updateState: getUpdateState(spreadsheet),
        },
        {
            type: 'color',
            content: 'format_color_text',
            tooltip: translate('Text color'),
            k: 'color',
            updateState: getUpdateState(spreadsheet),
        },
        {
            type: 'color',
            content: 'format_color_fill',
            tooltip: translate('Background color'),
            k: 'background-color',
            updateState: getUpdateState(spreadsheet),
        },
    ];
}

/**
 * The vertical alignment.
 */
function getVerticalAlignItem(spreadsheet: Spreadsheet): jspreadsheet.ToolbarItem {
    const verticalAlign = ['top', 'middle', 'bottom'];

    return {
        type: 'select',
        tooltip: translate('Vertical alignment'),
        options: ['vertical_align_top', 'vertical_align_center', 'vertical_align_bottom'],
        render(e: string) {
            return `<i class="material-icons">${e}</i>`;
        },
        value: 1,
        onchange(a: HTMLElement, b: object, c: string, d: string, value: string) {
            // The key is the index of the option as text
            setSelectedStyle(spreadsheet.getWorksheetInstance(), 'vertical-align', (verticalAlign as unknown as Record<string, string>)[value]);
        },
        updateState: getUpdateState(spreadsheet),
    };
}

/**
 * Merge the selected cells, or remove the merge of the selected merged cells.
 */
function getMergeItem(spreadsheet: Spreadsheet): jspreadsheet.ToolbarItem {
    return {
        content: 'web',
        tooltip: translate('Merge the selected cells'),
        onclick() {
            const worksheet = spreadsheet.getWorksheetInstance();

            if (worksheet.selection.range && confirm(translate('The merged cells will retain the value of the top-left cell only. Are you sure?'))) {
                const selectedRange = getSelectedRange(worksheet.selection.range);

                const cell = getCellNameFromCoords(selectedRange[0], selectedRange[1]);
                if (worksheet.records[selectedRange[1]][selectedRange[0]].element.getAttribute('data-merged')) {
                    worksheet.removeMerge(cell);
                } else {
                    const colspan = selectedRange[2] - selectedRange[0] + 1;
                    const rowspan = selectedRange[3] - selectedRange[1] + 1;

                    if (colspan !== 1 || rowspan !== 1) {
                        worksheet.setMerge(cell, colspan, rowspan);
                    }
                }
            }
        },
        updateState: getUpdateState(spreadsheet),
    };
}

/**
 * Set the borders of the selected cells.
 * @param picker - The picker of the borders, with the thickness, the color and the style that the user chose.
 * @param type - The option of the picker, such as 'border_outer'.
 */
function setBorders(worksheet: Worksheet, picker: BorderPicker, type: string) {
    if (worksheet.selection.range) {
        const selectedRange = getSelectedRange(worksheet.selection.range);

        // Default options
        let thickness = picker.thickness || 1;
        const color = picker.color || 'black';
        const borderStyle = picker.style || 'solid';

        if (borderStyle === 'double') {
            thickness += 2;
        }

        const style: Record<string, string> = {};

        // Matrix
        const px = selectedRange[0];
        const py = selectedRange[1];
        const ux = selectedRange[2];
        const uy = selectedRange[3];

        function setBorder(columnName: string, i: number, j: number) {
            const border = ['', '', '', ''];

            if (
                ((type === 'border_top' || type === 'border_outer') && j === py) ||
                ((type === 'border_inner' || type === 'border_horizontal') && j > py) ||
                type === 'border_all'
            ) {
                border[0] = `border-top: ${thickness}px ${borderStyle} ${color}`;
            } else {
                border[0] = 'border-top: ';
            }

            if ((type === 'border_all' || type === 'border_right' || type === 'border_outer') && i === ux) {
                border[1] = `border-right: ${thickness}px ${borderStyle} ${color}`;
            } else {
                border[1] = 'border-right: ';
            }

            if ((type === 'border_all' || type === 'border_bottom' || type === 'border_outer') && j === uy) {
                border[2] = `border-bottom: ${thickness}px ${borderStyle} ${color}`;
            } else {
                border[2] = 'border-bottom: ';
            }

            if (
                ((type === 'border_left' || type === 'border_outer') && i === px) ||
                ((type === 'border_inner' || type === 'border_vertical') && i > px) ||
                type === 'border_all'
            ) {
                border[3] = `border-left: ${thickness}px ${borderStyle} ${color}`;
            } else {
                border[3] = 'border-left: ';
            }

            style[columnName] = border.join(';');
        }

        for (let j = selectedRange[1]; j <= selectedRange[3]; j++) {
            // Row - py - uy
            for (let i = selectedRange[0]; i <= selectedRange[2]; i++) {
                // Col - px - ux
                setBorder(getCellNameFromCoords(i, j), i, j);

                if (worksheet.records[j][i].element.getAttribute('data-merged')) {
                    setBorder(getCellNameFromCoords(selectedRange[0], selectedRange[1]), i, j);
                }
            }
        }

        if (Object.keys(style)) {
            worksheet.setStyle(style);
        }
    }
}

/**
 * Add the controls for the color, the thickness and the style of the borders to the picker of the borders.
 * @param a - The element of the picker.
 * @param b - The picker, which keeps what the user chooses.
 */
function addBorderControls(a: HTMLElement, b: BorderPicker) {
    // Border color
    const container = document.createElement('div');
    let div = document.createElement('div');
    container.appendChild(div);

    const colorPicker = jSuites.color(div, {
        closeOnChange: false,
        onchange(o: HTMLElement, v: string) {
            (o.parentNode!.children[1] as HTMLElement).style.color = v;
            b.color = v;
        },
    });

    const i = document.createElement('i');
    i.classList.add('material-icons');
    i.innerHTML = 'color_lens';
    i.onclick = () => {
        colorPicker.open();
    };
    container.appendChild(i);
    a.children[1].appendChild(container);

    div = document.createElement('div');
    jSuites.picker(div, {
        data: [1, 2, 3, 4, 5],
        render(e: number) {
            return `<div style="height: ${e}px; width: 30px; background-color: black;"></div>`;
        },
        onchange(a: HTMLElement, k: object, c: number, d: number) {
            b.thickness = d;
        },
        width: '50px',
    });
    a.children[1].appendChild(div);

    const borderStylePicker = document.createElement('div');
    jSuites.picker(borderStylePicker, {
        data: ['solid', 'dotted', 'dashed', 'double'],
        render(e: string) {
            if (e === 'double') {
                return `<div style="width: 30px; border-top: 3px ${e} black;"></div>`;
            }
            return `<div style="width: 30px; border-top: 2px ${e} black;"></div>`;
        },
        onchange(a: HTMLElement, k: object, c: string, d: string) {
            b.style = d;
        },
        width: '50px',
    });
    a.children[1].appendChild(borderStylePicker);

    div = document.createElement('div');
    div.style.flex = '1';
    a.children[1].appendChild(div);
}

/**
 * The borders.
 */
function getBordersItem(spreadsheet: Spreadsheet): jspreadsheet.ToolbarItem {
    return {
        type: 'select',
        tooltip: translate('Borders'),
        options: [
            'border_all',
            'border_outer',
            'border_inner',
            'border_horizontal',
            'border_vertical',
            'border_left',
            'border_top',
            'border_right',
            'border_bottom',
            'border_clear',
        ],
        columns: 5,
        render(e: string) {
            return `<i class="material-icons">${e}</i>`;
        },
        right: true,
        onchange(a: HTMLElement, b: BorderPicker, c: string, d: string) {
            setBorders(spreadsheet.getWorksheetInstance(), b, d);
        },
        onload: addBorderControls,
        updateState: getUpdateState(spreadsheet),
    };
}

/**
 * Toggle the fullscreen of the spreadsheet.
 */
function getFullscreenItem(spreadsheet: Spreadsheet): jspreadsheet.ToolbarItem {
    return {
        content: 'fullscreen',
        tooltip: translate('Toggle Fullscreen'),
        onclick(a: HTMLElement, b: object, c: HTMLElement) {
            if (c.children[0].textContent === 'fullscreen') {
                spreadsheet.fullscreen(true);
                c.children[0].textContent = 'fullscreen_exit';
            } else {
                spreadsheet.fullscreen(false);
                c.children[0].textContent = 'fullscreen';
            }
        },
        updateState(a: HTMLElement, b: object, c: HTMLElement, d: jspreadsheet.WorksheetInstance) {
            if (d.parent.config.fullscreen === true) {
                c.children[0].textContent = 'fullscreen_exit';
            } else {
                c.children[0].textContent = 'fullscreen';
            }
        },
    };
}

/**
 * The toolbar of a spreadsheet: it builds the default items, creates the jSuites toolbar, shows and hides it, and updates
 * the state of its items for a worksheet.
 *
 * Each spreadsheet has one, as spreadsheet.toolbarManager, which the factory creates. spreadsheet.toolbar is the element
 * of the toolbar while it is shown, which applications use.
 */
export class Toolbar {
    readonly spreadsheet: Spreadsheet;

    constructor(spreadsheet: Spreadsheet) {
        this.spreadsheet = spreadsheet;
    }

    /**
     * The items of the default toolbar.
     */
    getDefault(): jspreadsheet.ToolbarItem[] {
        const spreadsheet = this.spreadsheet;

        return [
            ...getHistoryItems(spreadsheet),
            { type: 'divisor' },
            ...getFontItems(spreadsheet),
            getVerticalAlignItem(spreadsheet),
            getMergeItem(spreadsheet),
            getBordersItem(spreadsheet),
            { type: 'divisor' },
            getFullscreenItem(spreadsheet),
        ];
    }

    private adjustForJSuites(toolbar: jspreadsheet.ToolbarOptions) {
        const spreadsheet = this.spreadsheet;

        // The items are changed in place into items of jSuites.toolbar. A divisor has no other properties.
        const items = toolbar.items as jspreadsheet.ToolbarItemBase[];

        for (const item of items) {
            // jSuites.toolbar gives an item with a tooltip its title and its accessible name

            if (item.type === 'select') {
                if (item.options) {
                    item.data = item.options;
                    delete item.options;
                } else {
                    item.data = item.v;
                    delete item.v;

                    if (item.k && !item.onchange) {
                        item.onchange = (el: HTMLElement, config: object, value: string) => {
                            // The CSS property of a select item with v
                            setSelectedStyle(spreadsheet.getWorksheetInstance(), item.k as string, value);
                        };
                    }
                }
            } else if (item.type === 'color') {
                item.type = 'i';

                item.onclick = (a: HTMLElement, b: object, c: HTMLElement & { color?: Color }) => {
                    if (!c.color) {
                        jSuites.color(c, {
                            onchange(o: HTMLElement, v: string) {
                                // The CSS property of a color item
                                setSelectedStyle(spreadsheet.getWorksheetInstance(), item.k as string, v);
                            },
                            onopen(o: HTMLElement & { color: Color }) {
                                o.color.select('');
                            },
                        });

                        c.color!.open();
                    }
                };
            }
        }
    }

    /**
     * Create toolbar
     */
    create(toolbar: jspreadsheet.ToolbarOptions) {
        const spreadsheet = this.spreadsheet;

        const toolbarElement = document.createElement('div');
        toolbarElement.classList.add('jss_toolbar');

        this.adjustForJSuites(toolbar);

        if (typeof spreadsheet.plugins === 'object') {
            for (const plugin of Object.values(spreadsheet.plugins)) {
                if (typeof plugin.toolbar === 'function') {
                    const result = plugin.toolbar(toolbar);

                    if (result) {
                        toolbar = result;
                    }
                }
            }
        }

        jSuites.toolbar(toolbarElement, toolbar);

        this.addPickerKeys(toolbarElement);
        this.addToolbarKeys(toolbarElement);

        return toolbarElement;
    }

    /**
     * The items of the toolbar that take the focus: the buttons and the pickers.
     * @param toolbarElement - The element of the toolbar.
     */
    private getFocusableItems(toolbarElement: HTMLElement) {
        return Array.from(toolbarElement.querySelectorAll<HTMLElement>('.jtoolbar-item[role="button"], .jtoolbar-item.jpicker'));
    }

    /**
     * Make the toolbar work with the keyboard, with the WAI-ARIA toolbar pattern. The toolbar is one tab stop: Tab goes to the
     * item that had the focus last, the arrows, Home and End move between the items, and Enter and Space press a button.
     * @param toolbarElement - The element of the toolbar.
     */
    private addToolbarKeys(toolbarElement: HTMLElement) {
        toolbarElement.setAttribute('role', 'toolbar');
        toolbarElement.setAttribute('aria-label', translate('Toolbar'));

        const items = this.getFocusableItems(toolbarElement);
        for (const [i, item] of items.entries()) {
            item.setAttribute('tabindex', i === 0 ? '0' : '-1');
        }

        // The item that gets the focus, also from a press, is the tab stop
        toolbarElement.addEventListener('focusin', (e) => {
            const target = e.target as HTMLElement;
            const items = this.getFocusableItems(toolbarElement);

            if (items.includes(target)) {
                for (const item of items) {
                    item.setAttribute('tabindex', item === target ? '0' : '-1');
                }
            }
        });

        toolbarElement.addEventListener('keydown', (e) => this.toolbarKeys(toolbarElement, e));
    }

    /**
     * Handle a key on an item of the toolbar.
     * @param toolbarElement - The element of the toolbar.
     * @param e - The keydown.
     */
    private toolbarKeys(toolbarElement: HTMLElement, e: KeyboardEvent) {
        const target = e.target as HTMLElement;
        const items = this.getFocusableItems(toolbarElement);
        let index = items.indexOf(target);

        // An open picker uses the keys for its options
        if (index === -1 || target.classList.contains('jpicker-focus')) {
            return;
        }

        switch (e.key) {
            case 'ArrowRight':
                index = (index + 1) % items.length;
                break;
            case 'ArrowLeft':
                index = (index - 1 + items.length) % items.length;
                break;
            case 'Home':
                index = 0;
                break;
            case 'End':
                index = items.length - 1;
                break;
            case 'Enter':
            case ' ':
                // A picker opens with these keys itself
                if (target.getAttribute('role') === 'button') {
                    e.preventDefault();
                    if (!target.classList.contains('jtoolbar-disabled')) {
                        target.click();
                    }
                }
                return;
            default:
                return;
        }

        e.preventDefault();
        items[index].focus();
    }

    /**
     * Let the pickers of the toolbar be used from the keyboard, which jsuites does not do. Enter, Space and the arrows open a
     * picker, the arrows, Home and End move through its options, Enter and Space choose one, and Escape closes it. The picker
     * tells whether it is open in aria-expanded, and names the option that the keys are on in aria-activedescendant.
     * @param toolbarElement - The element of the toolbar.
     */
    private addPickerKeys(toolbarElement: HTMLElement) {
        for (const element of toolbarElement.querySelectorAll<PickerElement>('.jpicker')) {
            const picker = element.picker;

            if (picker) {
                // jsuites opens and closes a picker with these methods, also after a press, but leaves aria-expanded as it is
                const { open, close } = picker;

                picker.open = () => {
                    open();
                    element.setAttribute('aria-expanded', 'true');
                };

                picker.close = () => {
                    close();
                    element.setAttribute('aria-expanded', 'false');
                    element.removeAttribute('aria-activedescendant');
                    for (const option of element.querySelectorAll('.jss_picker_active')) {
                        option.classList.remove('jss_picker_active');
                    }
                };
            }
        }

        toolbarElement.addEventListener('keydown', (e) => this.pickerKeys(e));
    }

    /**
     * Handle a key on a picker of the toolbar.
     * @param e - The keydown.
     */
    private pickerKeys(e: KeyboardEvent) {
        const element = e.target as PickerElement;

        // Only the keys on the picker itself, not on the controls in its options, such as those of the borders
        if (!element.picker || !element.classList.contains('jpicker')) {
            return;
        }

        const picker = element.picker;
        const options = Array.from(element.querySelectorAll<HTMLElement & { k?: string }>(':scope > .jpicker-content > .jpicker-item'));

        if (!options.length) {
            return;
        }

        if (!element.classList.contains('jpicker-focus')) {
            if (['Enter', ' ', 'ArrowDown', 'ArrowUp'].includes(e.key)) {
                e.preventDefault();
                picker.open();

                // Start on the chosen option
                const index = options.findIndex((option) => `${option.k}` === `${picker.getValue()}`);
                this.setActiveOption(element, options, index === -1 ? 0 : index);
            }

            return;
        }

        const active = element.getAttribute('aria-activedescendant');
        let index = options.findIndex((option) => option.id !== '' && option.id === active);

        switch (e.key) {
            case 'ArrowDown':
            case 'ArrowRight':
                index = Math.min(index + 1, options.length - 1);
                break;
            case 'ArrowUp':
            case 'ArrowLeft':
                index = Math.max(index - 1, 0);
                break;
            case 'Home':
                index = 0;
                break;
            case 'End':
                index = options.length - 1;
                break;
            case 'Enter':
            case ' ':
                e.preventDefault();
                // jsuites chooses an option on a click, and closes the picker unless the option has controls
                options[index]?.click();
                return;
            case 'Escape':
                e.preventDefault();
                picker.close();
                return;
            case 'Tab':
                picker.close();
                return;
            default:
                return;
        }

        e.preventDefault();
        this.setActiveOption(element, options, index);
    }

    /**
     * Mark the option of a picker that the keys are on.
     * @param element - The element of the picker.
     * @param options - The options of the picker.
     * @param index - The index of the option.
     */
    private setActiveOption(element: HTMLElement, options: HTMLElement[], index: number) {
        for (const [i, option] of options.entries()) {
            if (!option.id) {
                // The list of options has an id, which the picker names in aria-controls
                option.id = `${element.getAttribute('aria-controls')}-${i}`;
            }
            option.classList.toggle('jss_picker_active', i === index);
        }

        element.setAttribute('aria-activedescendant', options[index].id);
    }

    update(worksheet: Worksheet) {
        if (worksheet.parent.toolbar) {
            // jsuites keeps the toolbar instance on the element
            worksheet.parent.toolbar.toolbar.update(worksheet);
        }
    }

    show() {
        const spreadsheet = this.spreadsheet;

        if (spreadsheet.config.toolbar && !spreadsheet.toolbar) {
            let toolbar: jspreadsheet.ToolbarOptions;

            if (Array.isArray(spreadsheet.config.toolbar)) {
                toolbar = {
                    items: spreadsheet.config.toolbar,
                };
            } else if (typeof spreadsheet.config.toolbar === 'object') {
                toolbar = spreadsheet.config.toolbar;
            } else {
                toolbar = {
                    items: this.getDefault(),
                };

                if (typeof spreadsheet.config.toolbar === 'function') {
                    toolbar = spreadsheet.config.toolbar(toolbar);
                }
            }

            // jSuites.toolbar in createToolbar added the toolbar instance
            spreadsheet.toolbar = spreadsheet.element.insertBefore(this.create(toolbar), spreadsheet.element.children[1]) as Spreadsheet['toolbar'];
        }
    }

    hide() {
        const spreadsheet = this.spreadsheet;

        if (spreadsheet.toolbar) {
            spreadsheet.toolbar.parentNode!.removeChild(spreadsheet.toolbar);

            // The public typings declare the toolbar as always there
            delete (spreadsheet as Partial<Spreadsheet>).toolbar;
        }
    }
}
