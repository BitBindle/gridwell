import { translate } from './dictionary.js';
import { getCellNameFromCoords } from './helpers.js';
import { ColorPicker, type ColorElement } from './color.js';
import { BordersDialog, type BorderOptions } from './borders.js';
import { ToolbarControl } from './toolbarControl.js';
import type jspreadsheet from '../jspreadsheet';
import type { Spreadsheet, Worksheet } from '../types/internal';
import { askConfirm } from './dialogs.js';

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

            if (
                worksheet.selection.range &&
                askConfirm(worksheet.parent.config, translate('The merged cells will retain the value of the top-left cell only. Are you sure?'))
            ) {
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
 * @param borders - The thickness, the color and the style that the user chose.
 * @param type - The type of borders, such as 'border_outer'.
 */
function setBorders(worksheet: Worksheet, borders: BorderOptions, type: string) {
    if (worksheet.selection.range) {
        const selectedRange = getSelectedRange(worksheet.selection.range);

        // Default options
        let thickness = borders.thickness || 1;
        const color = borders.color || 'black';
        const borderStyle = borders.style || 'solid';

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
 * The borders: a button that opens a dialog with the types of borders, and their color, thickness and style.
 */
function getBordersItem(spreadsheet: Spreadsheet): jspreadsheet.ToolbarItem {
    return {
        content: 'border_all',
        tooltip: translate('Borders'),
        render(element: HTMLElement) {
            new BordersDialog(element, (type, borders) => {
                setBorders(spreadsheet.getWorksheetInstance(), borders, type);
            });
        },
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
 * The toolbar of a spreadsheet: it builds the default items, creates the ToolbarControl, shows and hides it, and updates
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

    /**
     * Change the items in place into the items that the toolbar creates: a select item gets its options as data, and a
     * color item becomes a button that opens a color picker.
     * @param toolbar - The settings of the toolbar.
     */
    private adjustItems(toolbar: jspreadsheet.ToolbarOptions) {
        const spreadsheet = this.spreadsheet;

        // A divisor has no other properties
        const items = toolbar.items as jspreadsheet.ToolbarItemBase[];

        for (const item of items) {
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

                item.onclick = (a: HTMLElement, b: object, c: ColorElement, e: Event) => {
                    const picker =
                        c.color ||
                        new ColorPicker(c, {
                            onchange(o: HTMLElement, v: string) {
                                // The CSS property of a color item
                                setSelectedStyle(spreadsheet.getWorksheetInstance(), item.k as string, v);
                            },
                            onopen(o: ColorElement, picker: ColorPicker) {
                                picker.select('');
                            },
                        });

                    // A click in the picker, which is in the item, does not open it again
                    if (!picker.content.contains(e.target as Node)) {
                        picker.open();
                    }
                };
            }
        }
    }

    /**
     * Create the element of the toolbar.
     * @param toolbar - The settings of the toolbar.
     */
    create(toolbar: jspreadsheet.ToolbarOptions) {
        const spreadsheet = this.spreadsheet;

        const toolbarElement = document.createElement('div');
        toolbarElement.classList.add('jss_toolbar');

        this.adjustItems(toolbar);

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

        new ToolbarControl(toolbarElement, toolbar);

        return toolbarElement;
    }

    update(worksheet: Worksheet) {
        if (worksheet.parent.toolbar) {
            // The element keeps its ToolbarControl
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

            // create gives the element with its ToolbarControl
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
