import { translate } from './dictionary.js';

/** The colors that the color picker shows when it is given none: the colors of Material Design, by shade. */
const materialPalette = [
    [
        '#ffebee',
        '#fce4ec',
        '#f3e5f5',
        '#e8eaf6',
        '#e3f2fd',
        '#e0f7fa',
        '#e0f2f1',
        '#e8f5e9',
        '#f1f8e9',
        '#f9fbe7',
        '#fffde7',
        '#fff8e1',
        '#fff3e0',
        '#fbe9e7',
        '#efebe9',
        '#fafafa',
        '#eceff1',
    ],
    [
        '#ffcdd2',
        '#f8bbd0',
        '#e1bee7',
        '#c5cae9',
        '#bbdefb',
        '#b2ebf2',
        '#b2dfdb',
        '#c8e6c9',
        '#dcedc8',
        '#f0f4c3',
        '#fff9c4',
        '#ffecb3',
        '#ffe0b2',
        '#ffccbc',
        '#d7ccc8',
        '#f5f5f5',
        '#cfd8dc',
    ],
    [
        '#ef9a9a',
        '#f48fb1',
        '#ce93d8',
        '#9fa8da',
        '#90caf9',
        '#80deea',
        '#80cbc4',
        '#a5d6a7',
        '#c5e1a5',
        '#e6ee9c',
        '#fff59d',
        '#ffe082',
        '#ffcc80',
        '#ffab91',
        '#bcaaa4',
        '#eeeeee',
        '#b0bec5',
    ],
    [
        '#e57373',
        '#f06292',
        '#ba68c8',
        '#7986cb',
        '#64b5f6',
        '#4dd0e1',
        '#4db6ac',
        '#81c784',
        '#aed581',
        '#dce775',
        '#fff176',
        '#ffd54f',
        '#ffb74d',
        '#ff8a65',
        '#a1887f',
        '#e0e0e0',
        '#90a4ae',
    ],
    [
        '#ef5350',
        '#ec407a',
        '#ab47bc',
        '#5c6bc0',
        '#42a5f5',
        '#26c6da',
        '#26a69a',
        '#66bb6a',
        '#9ccc65',
        '#d4e157',
        '#ffee58',
        '#ffca28',
        '#ffa726',
        '#ff7043',
        '#8d6e63',
        '#bdbdbd',
        '#78909c',
    ],
    [
        '#f44336',
        '#e91e63',
        '#9c27b0',
        '#3f51b5',
        '#2196f3',
        '#00bcd4',
        '#009688',
        '#4caf50',
        '#8bc34a',
        '#cddc39',
        '#ffeb3b',
        '#ffc107',
        '#ff9800',
        '#ff5722',
        '#795548',
        '#9e9e9e',
        '#607d8b',
    ],
    [
        '#e53935',
        '#d81b60',
        '#8e24aa',
        '#3949ab',
        '#1e88e5',
        '#00acc1',
        '#00897b',
        '#43a047',
        '#7cb342',
        '#c0ca33',
        '#fdd835',
        '#ffb300',
        '#fb8c00',
        '#f4511e',
        '#6d4c41',
        '#757575',
        '#546e7a',
    ],
    [
        '#d32f2f',
        '#c2185b',
        '#7b1fa2',
        '#303f9f',
        '#1976d2',
        '#0097a7',
        '#00796b',
        '#388e3c',
        '#689f38',
        '#afb42b',
        '#fbc02d',
        '#ffa000',
        '#f57c00',
        '#e64a19',
        '#5d4037',
        '#616161',
        '#455a64',
    ],
    [
        '#c62828',
        '#ad1457',
        '#6a1b9a',
        '#283593',
        '#1565c0',
        '#00838f',
        '#00695c',
        '#2e7d32',
        '#558b2f',
        '#9e9d24',
        '#f9a825',
        '#ff8f00',
        '#ef6c00',
        '#d84315',
        '#4e342e',
        '#424242',
        '#37474f',
    ],
    [
        '#b71c1c',
        '#880e4f',
        '#4a148c',
        '#1a237e',
        '#0d47a1',
        '#006064',
        '#004d40',
        '#1b5e20',
        '#33691e',
        '#827717',
        '#f57f17',
        '#ff6f00',
        '#e65100',
        '#bf360c',
        '#3e2723',
        '#212121',
        '#263238',
    ],
];

/** The element of a color picker, which keeps its color picker. */
export type ColorElement = HTMLElement & { color?: ColorPicker };

/** The options of a color picker. */
export type ColorPickerOptions = {
    /** The color that the picker starts with, such as '#ff0000'. */
    value?: string | null;
    /** The placeholder of an input. */
    placeholder?: string;
    /** The rows of colors, such as [['#ff0000', '#00ff00']]. A color without # gets one. */
    palette?: string[][] | null;
    /** Close the picker when a color is chosen. */
    closeOnChange?: boolean;
    /** The text of the button that closes the picker. */
    doneLabel?: string;
    /** The text of the button that removes the color. */
    resetLabel?: string;
    /** Give the picker a fixed position, for a picker in an element that scrolls. */
    position?: boolean | null;
    /** Open the picker when it is created. */
    opened?: boolean;
    onopen?: ((element: ColorElement, picker: ColorPicker) => void) | null;
    onclose?: ((element: ColorElement, picker: ColorPicker) => void) | null;
    /**
     * Called when the color changes.
     * @param element - The element of the picker.
     * @param color - The color, or an empty text when it is removed.
     * @param picker - The picker.
     */
    onchange?: ((element: ColorElement, color: string, picker: ColorPicker) => void) | null;
};

/**
 * A color picker: a grid of colors, an input for any other color, and buttons that remove the color and close the picker.
 * The grid has the ARIA grid pattern: it is one tab stop, the arrows, Home and End move between the colors, and Enter
 * and Space choose one. Escape closes the picker, and so do a press outside it and the focus leaving it.
 *
 * The toolbar uses it for the colors of the text, of the background and of the borders, and the editor of a color
 * column for the cell. On an input, it shows after the input; on another element, it shows in the element. The element
 * keeps the picker as its color property, and the picker has the classes of the jSuites color picker.
 */
export class ColorPicker {
    /** The type, as jSuites names its components. */
    readonly type = 'color';

    /** The element of the picker. */
    readonly element: ColorElement;

    /** The options of the picker, with their defaults. */
    readonly options: Required<Omit<ColorPickerOptions, 'value' | 'palette'>> & { value: string; palette: string[][] };

    /** The element of the picker, which holds the grid, the input and the buttons. Its position can be set. */
    readonly content: HTMLElement;

    // The element that shows the picker while it has the class jcolor-focus
    private readonly container: HTMLElement & { color?: ColorPicker };

    // The cells of the grid, by color
    private readonly cells = new Map<string, HTMLTableCellElement>();

    // The input for any other color
    private readonly input: HTMLInputElement;

    constructor(element: ColorElement, options: ColorPickerOptions = {}) {
        this.element = element;
        this.options = {
            value: options.value || '',
            placeholder: options.placeholder ?? '',
            palette: options.palette || materialPalette,
            closeOnChange: options.closeOnChange ?? true,
            doneLabel: options.doneLabel ?? translate('Done'),
            resetLabel: options.resetLabel ?? translate('Reset'),
            position: options.position ?? null,
            opened: options.opened ?? false,
            onopen: options.onopen ?? null,
            onclose: options.onclose ?? null,
            onchange: options.onchange ?? null,
        };

        if (element instanceof HTMLInputElement) {
            element.classList.add('jcolor-input');
            element.readOnly = true;
            element.value = this.options.value;
            if (this.options.placeholder) {
                element.placeholder = this.options.placeholder;
            }
        }

        this.container = document.createElement('div');
        this.container.className = 'jcolor';

        this.content = document.createElement('div');
        this.content.className = 'jcolor-content';
        this.content.setAttribute('role', 'dialog');
        this.content.setAttribute('aria-label', translate('Color'));
        this.container.appendChild(this.content);

        const controls = document.createElement('div');
        controls.className = 'jcolor-controls';
        controls.appendChild(this.createButton('jcolor-reset', this.options.resetLabel, () => this.reset()));
        controls.appendChild(this.createButton('jcolor-close', this.options.doneLabel, () => this.close()));
        this.content.appendChild(controls);

        this.content.appendChild(this.createGrid());

        const custom = document.createElement('label');
        custom.className = 'jcolor-custom';
        const customText = document.createElement('span');
        customText.textContent = translate('Custom color');
        this.input = document.createElement('input');
        this.input.type = 'color';
        this.input.addEventListener('change', () => this.setValue(this.input.value));
        custom.appendChild(customText);
        custom.appendChild(this.input);
        this.content.appendChild(custom);

        this.content.addEventListener('keydown', this.keydown);
        this.container.addEventListener('focusout', this.focusOut);

        if (element instanceof HTMLInputElement) {
            element.parentNode!.insertBefore(this.container, element.nextSibling);
            element.addEventListener('mouseup', () => this.open());
        } else {
            element.appendChild(this.container);
        }

        element.color = this;
        this.container.color = this;

        if (this.options.opened) {
            this.open();
        }
    }

    /** Tell whether the picker is open. */
    isOpened() {
        return this.container.classList.contains('jcolor-focus');
    }

    /** Open the picker inside the window, and move the focus to the chosen color, or the first. */
    open() {
        if (this.isOpened()) {
            return;
        }

        this.container.classList.add('jcolor-focus');
        document.addEventListener('mousedown', this.pressOutside, true);
        document.addEventListener('touchstart', this.pressOutside, true);

        this.select(this.options.value);
        if (/^#[0-9a-f]{6}$/i.test(this.options.value)) {
            this.input.value = this.options.value;
        }

        this.content.style.marginTop = '';
        this.content.style.marginLeft = '';
        this.content.style.position = this.options.position ? 'fixed' : '';

        const rect = this.content.getBoundingClientRect();
        if (rect.left + rect.width > window.innerWidth) {
            this.content.style.marginLeft = `${-1 * (rect.left + rect.width - (window.innerWidth - 20))}px`;
        }
        if (rect.top + rect.height > window.innerHeight) {
            this.content.style.marginTop = `${-1 * (rect.top + rect.height - (window.innerHeight - 20))}px`;
        }

        this.options.onopen?.(this.element, this);

        this.getTabStop().focus({ preventScroll: true });
    }

    /**
     * Close the picker. When the focus is in the picker, it goes back to the element.
     * @param ignoreEvents - Do not call onclose.
     * @returns The color.
     */
    close(ignoreEvents?: boolean) {
        document.removeEventListener('mousedown', this.pressOutside, true);
        document.removeEventListener('touchstart', this.pressOutside, true);

        if (this.isOpened()) {
            const hadFocus = this.container.contains(document.activeElement);
            this.container.classList.remove('jcolor-focus');
            if (hadFocus && this.element.isConnected) {
                this.element.focus({ preventScroll: true });
            }

            if (!ignoreEvents) {
                this.options.onclose?.(this.element, this);
            }
        }

        return this.options.value;
    }

    /** The color, or an empty text when there is none. */
    getValue() {
        return this.options.value;
    }

    /**
     * Change the color, and call onchange. The picker closes unless closeOnChange is false.
     * @param color - The color, or an empty text to remove it.
     */
    setValue(color: string | null | undefined) {
        color = color || '';

        if (color !== this.options.value) {
            this.options.value = color;
            this.select(color);

            this.options.onchange?.(this.element, color, this);

            if (this.element instanceof HTMLInputElement) {
                this.element.value = color;
            }

            if (this.options.closeOnChange) {
                this.close();
            }
        }
    }

    /**
     * Mark a color of the grid as chosen, without calling onchange.
     * @param color - The color.
     */
    select(color: string) {
        this.options.value = color;
        for (const [value, cell] of this.cells) {
            const selected = value.toLowerCase() === color.toLowerCase();
            cell.classList.toggle('jcolor-selected', selected);
            cell.setAttribute('aria-selected', selected ? 'true' : 'false');
        }
        this.setTabStop(this.getTabStop());
    }

    /** Remove the color and close the picker. */
    private reset() {
        this.setValue('');
        this.close();
    }

    /**
     * A button of the picker.
     * @param className - The class of the button.
     * @param text - Its text.
     * @param onclick - What it does.
     */
    private createButton(className: string, text: string, onclick: () => void) {
        const button = document.createElement('button');
        button.type = 'button';
        button.className = className;
        button.textContent = text;
        button.addEventListener('click', onclick);
        return button;
    }

    /** The grid of the colors of the palette. */
    private createGrid() {
        const grid = document.createElement('div');
        grid.className = 'jcolor-grid';

        const table = document.createElement('table');
        table.setAttribute('role', 'grid');
        table.setAttribute('aria-label', translate('Colors'));

        for (const colors of this.options.palette) {
            const tr = document.createElement('tr');
            for (let color of colors) {
                if (color.length < 7 && color[0] !== '#') {
                    color = `#${color}`;
                }
                const td = document.createElement('td');
                td.style.backgroundColor = color;
                td.setAttribute('data-value', color);
                td.setAttribute('role', 'gridcell');
                td.setAttribute('aria-label', color);
                td.setAttribute('aria-selected', 'false');
                td.setAttribute('tabindex', '-1');
                tr.appendChild(td);
                this.cells.set(color, td);
            }
            table.appendChild(tr);
        }

        table.addEventListener('click', (e) => {
            const td = (e.target as Element).closest('td');
            if (td) {
                this.choose(td);
            }
        });

        grid.appendChild(table);
        return grid;
    }

    /**
     * Choose the color of a cell of the grid.
     * @param td - The cell.
     */
    private choose(td: HTMLTableCellElement) {
        this.setTabStop(td);
        this.setValue(td.getAttribute('data-value'));
    }

    /** The cell of the grid that is the tab stop: the chosen color, or the first. */
    private getTabStop() {
        const cells = Array.from(this.cells.values());
        return cells.find((cell) => cell.classList.contains('jcolor-selected')) || cells[0];
    }

    /**
     * Make a cell of the grid the one that Tab goes to.
     * @param td - The cell.
     */
    private setTabStop(td: HTMLTableCellElement | undefined) {
        for (const cell of this.cells.values()) {
            cell.setAttribute('tabindex', cell === td ? '0' : '-1');
        }
    }

    /** The keys on the grid, and Escape on the picker. */
    private keydown = (e: KeyboardEvent) => {
        if (e.key === 'Escape') {
            // The editor of a cell closes without saving on Escape, so onclose is not called
            this.close(true);
            return;
        }

        const td = e.target as HTMLTableCellElement;
        if (td.tagName !== 'TD' || !this.content.contains(td)) {
            return;
        }

        const tr = td.parentElement as HTMLTableRowElement;
        const rows = Array.from(tr.parentElement!.children) as HTMLTableRowElement[];
        let row = rows.indexOf(tr);
        let column = td.cellIndex;

        switch (e.key) {
            case 'ArrowRight':
                column++;
                break;
            case 'ArrowLeft':
                column--;
                break;
            case 'ArrowDown':
                row++;
                break;
            case 'ArrowUp':
                row--;
                break;
            case 'Home':
                column = 0;
                break;
            case 'End':
                column = tr.cells.length - 1;
                break;
            case 'Enter':
            case ' ':
                this.choose(td);
                break;
            default:
                return;
        }

        e.preventDefault();
        e.stopPropagation();

        if (e.key !== 'Enter' && e.key !== ' ') {
            row = Math.max(0, Math.min(row, rows.length - 1));
            const cells = rows[row].cells;
            const next = cells[Math.max(0, Math.min(column, cells.length - 1))];
            this.setTabStop(next);
            next.focus();
        }
    };

    /** The focus leaving the picker and its element closes it. */
    private focusOut = (e: FocusEvent) => {
        const next = e.relatedTarget as Node | null;
        if (next && !this.container.contains(next) && next !== this.element) {
            this.close();
        }
    };

    /** A press outside the picker and its element closes it. */
    private pressOutside = (e: Event) => {
        const path = e.composedPath();
        if (!path.includes(this.container) && !path.includes(this.element)) {
            this.close();
        }
    };
}
