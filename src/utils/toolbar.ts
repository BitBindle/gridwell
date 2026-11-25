import jSuites from 'jsuites';
import { getCellNameFromCoords } from './helpers.js';
import { getWorksheetInstance } from './internal.js';
import type { Spreadsheet, Worksheet } from '../types/internal';

/** A jsuites color picker. jsuites also keeps it on its element as `color`. */
type ColorPicker = ReturnType<typeof jSuites.color>;

/** The border select keeps the thickness, the color and the style that the user chose on its picker instance. */
type BorderPicker = { thickness?: number; color?: string; style?: string };

const setItemStatus = function (toolbarItem: HTMLElement, worksheet: Worksheet) {
    if (worksheet.options.editable != false) {
        toolbarItem.classList.remove('jtoolbar-disabled');
    } else {
        toolbarItem.classList.add('jtoolbar-disabled');
    }
};

export const getDefault = function (this: Spreadsheet) {
    // The public ToolbarItem type requires content, which the select items here do not have
    const items: any[] = [];
    const spreadsheet = this;

    const getActive = function () {
        return getWorksheetInstance.call(spreadsheet);
    };

    items.push({
        content: 'undo',
        onclick: function () {
            const worksheet = getActive();

            worksheet.undo();
        },
    });

    items.push({
        content: 'redo',
        onclick: function () {
            const worksheet = getActive();

            worksheet.redo();
        },
    });

    items.push({
        content: 'save',
        onclick: function () {
            const worksheet = getActive();

            if (worksheet) {
                worksheet.download();
            }
        },
    });

    items.push({
        type: 'divisor',
    });

    items.push({
        type: 'select',
        width: '120px',
        options: ['Default', 'Verdana', 'Arial', 'Courier New'],
        render: function (e: string) {
            return '<span style="font-family:' + e + '">' + e + '</span>';
        },
        onchange: function (a: HTMLElement, b: object, c: string, d: string, e: string) {
            const worksheet = getActive();

            const cells = worksheet.getSelected(true);
            if (cells) {
                // e is the key of the option as text, and the first option is Default
                const value = e === '0' ? '' : d;

                worksheet.setStyle(
                    Object.fromEntries(
                        cells.map(function (cellName: string) {
                            return [cellName, 'font-family: ' + value];
                        })
                    )
                );
            }
        },
        updateState: function (a: HTMLElement, b: object, toolbarItem: HTMLElement) {
            setItemStatus(toolbarItem, getActive());
        },
    });

    items.push({
        type: 'select',
        width: '48px',
        content: 'format_size',
        options: ['x-small', 'small', 'medium', 'large', 'x-large'],
        render: function (e: string) {
            return '<span style="font-size:' + e + '">' + e + '</span>';
        },
        onchange: function (a: HTMLElement, b: object, c: string, value: string) {
            const worksheet = getActive();

            const cells = worksheet.getSelected(true);
            if (cells) {
                worksheet.setStyle(
                    Object.fromEntries(
                        cells.map(function (cellName: string) {
                            return [cellName, 'font-size: ' + value];
                        })
                    )
                );
            }
        },
        updateState: function (a: HTMLElement, b: object, toolbarItem: HTMLElement) {
            setItemStatus(toolbarItem, getActive());
        },
    });

    items.push({
        type: 'select',
        options: ['left', 'center', 'right', 'justify'],
        render: function (e: string) {
            return '<i class="material-icons">format_align_' + e + '</i>';
        },
        onchange: function (a: HTMLElement, b: object, c: string, value: string) {
            const worksheet = getActive();

            const cells = worksheet.getSelected(true);
            if (cells) {
                worksheet.setStyle(
                    Object.fromEntries(
                        cells.map(function (cellName: string) {
                            return [cellName, 'text-align: ' + value];
                        })
                    )
                );
            }
        },
        updateState: function (a: HTMLElement, b: object, toolbarItem: HTMLElement) {
            setItemStatus(toolbarItem, getActive());
        },
    });

    items.push({
        content: 'format_bold',
        onclick: function () {
            const worksheet = getActive();

            const cells = worksheet.getSelected(true);
            if (cells) {
                worksheet.setStyle(
                    Object.fromEntries(
                        cells.map(function (cellName: string) {
                            return [cellName, 'font-weight:bold'];
                        })
                    )
                );
            }
        },
        updateState: function (a: HTMLElement, b: object, toolbarItem: HTMLElement) {
            setItemStatus(toolbarItem, getActive());
        },
    });

    items.push({
        type: 'color',
        content: 'format_color_text',
        k: 'color',
        updateState: function (a: HTMLElement, b: object, toolbarItem: HTMLElement) {
            setItemStatus(toolbarItem, getActive());
        },
    });

    items.push({
        type: 'color',
        content: 'format_color_fill',
        k: 'background-color',
        updateState: function (a: HTMLElement, b: object, toolbarItem: HTMLElement, d: Worksheet) {
            setItemStatus(toolbarItem, getActive());
        },
    });

    const verticalAlign = ['top', 'middle', 'bottom'];

    items.push({
        type: 'select',
        options: ['vertical_align_top', 'vertical_align_center', 'vertical_align_bottom'],
        render: function (e: string) {
            return '<i class="material-icons">' + e + '</i>';
        },
        value: 1,
        onchange: function (a: HTMLElement, b: object, c: string, d: string, value: string) {
            const worksheet = getActive();

            const cells = worksheet.getSelected(true);
            if (cells) {
                worksheet.setStyle(
                    Object.fromEntries(
                        cells.map(function (cellName: string) {
                            // The key is the index of the option as text
                            return [cellName, 'vertical-align: ' + (verticalAlign as unknown as Record<string, string>)[value]];
                        })
                    )
                );
            }
        },
        updateState: function (a: HTMLElement, b: object, toolbarItem: HTMLElement) {
            setItemStatus(toolbarItem, getActive());
        },
    });

    items.push({
        content: 'web',
        tooltip: jSuites.translate('Merge the selected cells'),
        onclick: function () {
            const worksheet = getActive();

            if (worksheet.selectedCell && confirm(jSuites.translate('The merged cells will retain the value of the top-left cell only. Are you sure?'))) {
                const selectedRange = [
                    Math.min(worksheet.selectedCell[0], worksheet.selectedCell[2]),
                    Math.min(worksheet.selectedCell[1], worksheet.selectedCell[3]),
                    Math.max(worksheet.selectedCell[0], worksheet.selectedCell[2]),
                    Math.max(worksheet.selectedCell[1], worksheet.selectedCell[3]),
                ];

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
        updateState: function (a: HTMLElement, b: object, toolbarItem: HTMLElement) {
            setItemStatus(toolbarItem, getActive());
        },
    });

    items.push({
        type: 'select',
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
        render: function (e: string) {
            return '<i class="material-icons">' + e + '</i>';
        },
        right: true,
        onchange: function (a: HTMLElement, b: BorderPicker, c: string, d: string) {
            const worksheet = getActive();

            if (worksheet.selectedCell) {
                const selectedRange = [
                    Math.min(worksheet.selectedCell[0], worksheet.selectedCell[2]),
                    Math.min(worksheet.selectedCell[1], worksheet.selectedCell[3]),
                    Math.max(worksheet.selectedCell[0], worksheet.selectedCell[2]),
                    Math.max(worksheet.selectedCell[1], worksheet.selectedCell[3]),
                ];

                const type = d;

                if (selectedRange) {
                    // Default options
                    let thickness = b.thickness || 1;
                    const color = b.color || 'black';
                    const borderStyle = b.style || 'solid';

                    if (borderStyle === 'double') {
                        thickness += 2;
                    }

                    const style: Record<string, string> = {};

                    // Matrix
                    const px = selectedRange[0];
                    const py = selectedRange[1];
                    const ux = selectedRange[2];
                    const uy = selectedRange[3];

                    const setBorder = function (columnName: string, i: number, j: number) {
                        const border = ['', '', '', ''];

                        if (
                            ((type === 'border_top' || type === 'border_outer') && j === py) ||
                            ((type === 'border_inner' || type === 'border_horizontal') && j > py) ||
                            type === 'border_all'
                        ) {
                            border[0] = 'border-top: ' + thickness + 'px ' + borderStyle + ' ' + color;
                        } else {
                            border[0] = 'border-top: ';
                        }

                        if ((type === 'border_all' || type === 'border_right' || type === 'border_outer') && i === ux) {
                            border[1] = 'border-right: ' + thickness + 'px ' + borderStyle + ' ' + color;
                        } else {
                            border[1] = 'border-right: ';
                        }

                        if ((type === 'border_all' || type === 'border_bottom' || type === 'border_outer') && j === uy) {
                            border[2] = 'border-bottom: ' + thickness + 'px ' + borderStyle + ' ' + color;
                        } else {
                            border[2] = 'border-bottom: ';
                        }

                        if (
                            ((type === 'border_left' || type === 'border_outer') && i === px) ||
                            ((type === 'border_inner' || type === 'border_vertical') && i > px) ||
                            type === 'border_all'
                        ) {
                            border[3] = 'border-left: ' + thickness + 'px ' + borderStyle + ' ' + color;
                        } else {
                            border[3] = 'border-left: ';
                        }

                        style[columnName] = border.join(';');
                    };

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
        },
        onload: function (a: HTMLElement, b: BorderPicker) {
            // Border color
            const container = document.createElement('div');
            let div = document.createElement('div');
            container.appendChild(div);

            const colorPicker = jSuites.color(div, {
                closeOnChange: false,
                onchange: function (o: HTMLElement, v: string) {
                    (o.parentNode!.children[1] as HTMLElement).style.color = v;
                    b.color = v;
                },
            } as any);

            const i = document.createElement('i');
            i.classList.add('material-icons');
            i.innerHTML = 'color_lens';
            i.onclick = function () {
                colorPicker.open();
            };
            container.appendChild(i);
            a.children[1].appendChild(container);

            div = document.createElement('div');
            jSuites.picker(div, {
                type: 'select',
                data: [1, 2, 3, 4, 5],
                render: function (e: number) {
                    return '<div style="height: ' + e + 'px; width: 30px; background-color: black;"></div>';
                },
                onchange: function (a: HTMLElement, k: object, c: number, d: number) {
                    b.thickness = d;
                },
                width: '50px',
            } as any);
            a.children[1].appendChild(div);

            const borderStylePicker = document.createElement('div');
            jSuites.picker(borderStylePicker, {
                type: 'select',
                data: ['solid', 'dotted', 'dashed', 'double'],
                render: function (e: string) {
                    if (e === 'double') {
                        return '<div style="width: 30px; border-top: 3px ' + e + ' black;"></div>';
                    }
                    return '<div style="width: 30px; border-top: 2px ' + e + ' black;"></div>';
                },
                onchange: function (a: HTMLElement, k: object, c: string, d: string) {
                    b.style = d;
                },
                width: '50px',
            } as any);
            a.children[1].appendChild(borderStylePicker);

            div = document.createElement('div');
            div.style.flex = '1';
            a.children[1].appendChild(div);
        },
        updateState: function (a: HTMLElement, b: object, toolbarItem: HTMLElement) {
            setItemStatus(toolbarItem, getActive());
        },
    });

    items.push({
        type: 'divisor',
    });

    items.push({
        content: 'fullscreen',
        tooltip: 'Toggle Fullscreen',
        onclick: function (a: HTMLElement, b: object, c: HTMLElement) {
            if (c.children[0].textContent === 'fullscreen') {
                spreadsheet.fullscreen(true);
                c.children[0].textContent = 'fullscreen_exit';
            } else {
                spreadsheet.fullscreen(false);
                c.children[0].textContent = 'fullscreen';
            }
        },
        updateState: function (a: HTMLElement, b: object, c: HTMLElement, d: Worksheet) {
            if (d.parent.config.fullscreen === true) {
                c.children[0].textContent = 'fullscreen_exit';
            } else {
                c.children[0].textContent = 'fullscreen';
            }
        },
    });

    return items;
};

const adjustToolbarSettingsForJSuites = function (this: Spreadsheet, toolbar: any) {
    const spreadsheet = this;

    const items = toolbar.items;

    for (let i = 0; i < items.length; i++) {
        // Tooltip
        if (items[i].tooltip) {
            items[i].title = items[i].tooltip;

            delete items[i].tooltip;
        }

        if (items[i].type == 'select') {
            if (items[i].options) {
                items[i].data = items[i].options;
                delete items[i].options;
            } else {
                items[i].data = items[i].v;
                delete items[i].v;

                if (items[i].k && !items[i].onchange) {
                    items[i].onchange = function (el: HTMLElement, config: object, value: string) {
                        const worksheet = getWorksheetInstance.call(spreadsheet);

                        const cells = worksheet.getSelected(true);

                        worksheet.setStyle(
                            Object.fromEntries(
                                cells.map(function (cellName: string) {
                                    return [cellName, items[i].k + ': ' + value];
                                })
                            )
                        );
                    };
                }
            }
        } else if (items[i].type == 'color') {
            items[i].type = 'i';

            items[i].onclick = function (a: HTMLElement, b: object, c: HTMLElement & { color?: ColorPicker }) {
                if (!c.color) {
                    jSuites.color(c, {
                        onchange: function (o: HTMLElement, v: string) {
                            const worksheet = getWorksheetInstance.call(spreadsheet);

                            const cells = worksheet.getSelected(true);

                            worksheet.setStyle(
                                Object.fromEntries(
                                    cells.map(function (cellName: string) {
                                        return [cellName, items[i].k + ': ' + v];
                                    })
                                )
                            );
                        },
                        onopen: function (o: HTMLElement & { color: ColorPicker }) {
                            o.color.select('');
                        },
                    } as any);

                    c.color!.open();
                }
            };
        }
    }
};

/**
 * Create toolbar
 */
export const createToolbar = function (this: Spreadsheet, toolbar: any) {
    const spreadsheet = this;

    const toolbarElement = document.createElement('div');
    toolbarElement.classList.add('jss_toolbar');

    adjustToolbarSettingsForJSuites.call(spreadsheet, toolbar);

    if (typeof spreadsheet.plugins === 'object') {
        Object.entries(spreadsheet.plugins).forEach(function ([, plugin]) {
            if (typeof plugin.toolbar === 'function') {
                const result = plugin.toolbar(toolbar);

                if (result) {
                    toolbar = result;
                }
            }
        });
    }

    jSuites.toolbar(toolbarElement, toolbar);

    return toolbarElement;
};

export const updateToolbar = function (worksheet: Worksheet) {
    if (worksheet.parent.toolbar) {
        // jsuites keeps the toolbar instance on the element
        worksheet.parent.toolbar.toolbar.update(worksheet);
    }
};

export const showToolbar = function (this: Spreadsheet) {
    const spreadsheet = this;

    if (spreadsheet.config.toolbar && !spreadsheet.toolbar) {
        let toolbar: any;

        if (Array.isArray(spreadsheet.config.toolbar)) {
            toolbar = {
                items: spreadsheet.config.toolbar,
            };
        } else if (typeof spreadsheet.config.toolbar === 'object') {
            toolbar = spreadsheet.config.toolbar;
        } else {
            toolbar = {
                items: getDefault.call(spreadsheet),
            };

            if (typeof spreadsheet.config.toolbar === 'function') {
                toolbar = spreadsheet.config.toolbar(toolbar);
            }
        }

        // jSuites.toolbar in createToolbar added the toolbar instance
        spreadsheet.toolbar = spreadsheet.element.insertBefore(
            createToolbar.call(spreadsheet, toolbar),
            spreadsheet.element.children[1]
        ) as Spreadsheet['toolbar'];
    }
};

export const hideToolbar = function (this: Spreadsheet) {
    const spreadsheet = this;

    if (spreadsheet.toolbar) {
        spreadsheet.toolbar.parentNode!.removeChild(spreadsheet.toolbar);

        // The public typings declare the toolbar as always there
        delete (spreadsheet as Partial<Spreadsheet>).toolbar;
    }
};
