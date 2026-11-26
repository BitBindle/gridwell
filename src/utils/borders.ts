import { translate } from './dictionary.js';
import { createId } from './ids.js';
import { ColorPicker } from './color.js';
import { Picker } from './picker.js';

/** The color, the thickness and the style of the borders that the user chose. */
export type BorderOptions = { color?: string; thickness?: number; style?: string };

/** The number of columns of the grid of the types of borders. */
const columns = 5;

/** The types of borders, by their icons (from material icons), with their names. */
function getBorderTypes(): [string, string][] {
    return [
        ['border_all', translate('All borders')],
        ['border_outer', translate('Outer borders')],
        ['border_inner', translate('Inner borders')],
        ['border_horizontal', translate('Horizontal borders')],
        ['border_vertical', translate('Vertical borders')],
        ['border_left', translate('Left border')],
        ['border_top', translate('Top border')],
        ['border_right', translate('Right border')],
        ['border_bottom', translate('Bottom border')],
        ['border_clear', translate('No borders')],
    ];
}

/**
 * A button with an icon, which the text of the icon does not name.
 * @param icon - The icon, from material icons.
 * @param label - The name of the button.
 */
function createIconButton(icon: string, label: string) {
    const button = document.createElement('button');
    button.type = 'button';
    button.setAttribute('aria-label', label);
    button.setAttribute('title', label);

    const i = document.createElement('i');
    i.classList.add('material-icons');
    i.setAttribute('aria-hidden', 'true');
    i.textContent = icon;
    button.appendChild(i);

    return button;
}

/**
 * The borders item of the toolbar: a button that opens a dialog, which is not modal. The dialog has the types of borders,
 * then the border color, the border thickness and the border style.
 *
 * The types of borders are one tab stop: the arrows move between them in their grid, and Home and End go to the first and
 * the last. Tab and Shift+Tab move between the types, the color, the thickness and the style, and stay in the dialog.
 * Choosing a type of borders sets the borders of the selected cells and closes the dialog. Escape closes the dialog, or
 * only the color picker or the list of a picker when one is open, and so does a press outside it. The focus goes back to
 * the button when the dialog closes.
 */
export class BordersDialog {
    /** The button of the toolbar that opens the dialog. */
    readonly button: HTMLElement;

    /** The dialog. The button puts it after itself the first time it opens it. */
    readonly element: HTMLDivElement;

    /** The color, the thickness and the style that the user chose. */
    readonly borders: BorderOptions = {};

    // The buttons of the types of borders
    private readonly types: HTMLButtonElement[] = [];

    // The button that opens the color picker, the element of the color picker, and the color picker
    private readonly colorButton: HTMLButtonElement;
    private readonly colorElement: HTMLDivElement;
    private readonly colorPicker: ColorPicker;

    // The pickers of the thickness and the style
    private readonly thickness: Picker<number>;
    private readonly style: Picker<string>;

    // Called when the user chooses a type of borders
    private readonly onchange: (type: string, borders: BorderOptions) => void;

    /**
     * @param button - The button of the toolbar, with the icon of the type of borders.
     * @param onchange - Called when the user chooses a type of borders, with its icon, such as 'border_outer'.
     */
    constructor(button: HTMLElement, onchange: (type: string, borders: BorderOptions) => void) {
        this.button = button;
        this.onchange = onchange;
        button.classList.add('jss_borders_button');
        button.setAttribute('aria-haspopup', 'dialog');
        button.setAttribute('aria-expanded', 'false');
        button.addEventListener('click', () => {
            if (this.isOpened()) {
                this.close();
            } else {
                this.open();
            }
        });
        button.addEventListener('keydown', (e) => {
            if (e.key === 'ArrowDown' && !this.isOpened()) {
                e.preventDefault();
                this.open();
            }
        });

        this.element = document.createElement('div');
        this.element.id = createId('jss-borders');
        this.element.className = 'jss_borders';
        this.element.setAttribute('role', 'dialog');
        this.element.setAttribute('aria-label', translate('Borders'));
        this.element.hidden = true;
        this.element.addEventListener('keydown', this.keydown);

        const group = document.createElement('div');
        group.className = 'jss_borders_types';
        group.setAttribute('role', 'group');
        for (const [i, [type, label]] of getBorderTypes().entries()) {
            const item = createIconButton(type, label);
            item.setAttribute('tabindex', i === 0 ? '0' : '-1');
            item.addEventListener('click', () => this.choose(item, type));
            group.appendChild(item);
            this.types.push(item);
        }
        this.element.appendChild(group);

        const controls = document.createElement('div');
        controls.className = 'jss_borders_controls';

        // The color picker shows in its element, before the button that opens it
        const color = document.createElement('div');
        this.colorElement = document.createElement('div');
        this.colorButton = createIconButton('color_lens', translate('Border color'));
        this.colorButton.addEventListener('click', () => this.colorPicker.open());
        color.appendChild(this.colorElement);
        color.appendChild(this.colorButton);
        controls.appendChild(color);

        this.colorPicker = new ColorPicker(this.colorElement, {
            closeOnChange: false,
            onchange: (element: HTMLElement, value: string) => {
                this.colorButton.style.color = value;
                this.borders.color = value;
            },
            onclose: () => {
                // The element of the color picker does not take the focus, so the button that opened it takes it back
                if (this.colorElement.contains(document.activeElement)) {
                    this.colorButton.focus();
                }
            },
        });

        const thickness = document.createElement('div');
        thickness.setAttribute('aria-label', translate('Border thickness'));
        this.thickness = new Picker<number>(thickness, {
            data: [1, 2, 3, 4, 5],
            render(value: number) {
                return `<div style="height: ${value}px; width: 30px; background-color: black;"></div>`;
            },
            onchange: (element: HTMLElement, picker: Picker<number>, value: number) => {
                this.borders.thickness = value;
            },
            width: '50px',
        });
        controls.appendChild(thickness);

        const style = document.createElement('div');
        style.setAttribute('aria-label', translate('Border style'));
        this.style = new Picker<string>(style, {
            data: ['solid', 'dotted', 'dashed', 'double'],
            render(value: string) {
                return `<div style="width: 30px; border-top: ${value === 'double' ? 3 : 2}px ${value} black;"></div>`;
            },
            onchange: (element: HTMLElement, picker: Picker<string>, value: string) => {
                this.borders.style = value;
            },
            width: '50px',
        });
        controls.appendChild(style);

        this.element.appendChild(controls);
    }

    /** Tell whether the dialog is open. */
    isOpened() {
        return !this.element.hidden;
    }

    /** Open the dialog under the button, aligned with its right side and inside the window, and move the focus into it. */
    open() {
        if (this.isOpened() || this.button.getAttribute('aria-disabled') === 'true') {
            return;
        }

        if (this.element.previousElementSibling !== this.button) {
            this.button.after(this.element);
            this.button.setAttribute('aria-controls', this.element.id);
        }

        this.element.hidden = false;
        this.button.setAttribute('aria-expanded', 'true');
        document.addEventListener('mousedown', this.pressOutside, true);
        document.addEventListener('touchstart', this.pressOutside, true);

        let left = this.button.offsetLeft + this.button.offsetWidth - this.element.offsetWidth;
        this.element.style.left = `${left}px`;
        this.element.style.top = `${this.button.offsetTop + this.button.offsetHeight + 4}px`;
        const rect = this.element.getBoundingClientRect();
        if (rect.left < 0) {
            left += 10 - rect.left;
        } else if (rect.right > window.innerWidth) {
            left -= 10 + rect.right - window.innerWidth;
        }
        this.element.style.left = `${left}px`;

        this.getTypeTabStop().focus();
    }

    /** Close the dialog, with the color picker and the pickers in it. When the focus is in the dialog, it goes back to the button. */
    close() {
        document.removeEventListener('mousedown', this.pressOutside, true);
        document.removeEventListener('touchstart', this.pressOutside, true);

        if (!this.isOpened()) {
            return;
        }

        const hadFocus = this.element.contains(document.activeElement);

        this.colorPicker.close(true);
        this.thickness.close();
        this.style.close();

        this.element.hidden = true;
        this.button.setAttribute('aria-expanded', 'false');

        if (hadFocus) {
            this.button.focus();
        }
    }

    /**
     * Choose a type of borders: the button shows its icon, the dialog closes, and onchange sets the borders.
     * @param item - The button of the type.
     * @param type - Its icon, such as 'border_outer'.
     */
    private choose(item: HTMLButtonElement, type: string) {
        this.setTypeTabStop(item);
        const icon = this.button.querySelector('i');
        if (icon) {
            icon.textContent = type;
        }

        this.close();

        this.onchange(type, this.borders);
    }

    /** The button of the type of borders that is the tab stop. */
    private getTypeTabStop() {
        return this.types.find((item) => item.getAttribute('tabindex') === '0') || this.types[0];
    }

    /**
     * Make a button of a type of borders the one that Tab goes to.
     * @param item - The button.
     */
    private setTypeTabStop(item: HTMLButtonElement) {
        for (const type of this.types) {
            type.setAttribute('tabindex', type === item ? '0' : '-1');
        }
    }

    /** The elements that Tab goes to in the dialog, in order, without those of the color picker. */
    private getTabStops(): HTMLElement[] {
        return [this.getTypeTabStop(), this.colorButton, this.thickness.element, this.style.element];
    }

    /** The keys in the dialog. */
    private keydown = (e: KeyboardEvent) => {
        const target = e.target as HTMLElement;

        if (e.key === 'Escape') {
            // An open picker closed its list
            if (e.defaultPrevented) {
                return;
            }
            e.preventDefault();
            if (this.colorElement.contains(target)) {
                // The color picker closed itself
                this.colorButton.focus();
            } else {
                this.close();
            }
        } else if (e.key === 'Tab') {
            const stops = this.getTabStops();
            const index = stops.indexOf(target);
            if (e.shiftKey && index === 0) {
                e.preventDefault();
                stops[stops.length - 1].focus();
            } else if (!e.shiftKey && index === stops.length - 1) {
                e.preventDefault();
                stops[0].focus();
            }
        } else {
            this.moveInTypes(e, target);
        }
    };

    /**
     * Move between the types of borders with the arrows, Home and End.
     * @param e - The key.
     * @param target - The element that has the focus.
     */
    private moveInTypes(e: KeyboardEvent, target: HTMLElement) {
        const current = this.types.indexOf(target as HTMLButtonElement);
        if (current === -1) {
            return;
        }

        let index = current;

        switch (e.key) {
            case 'ArrowRight':
                index++;
                break;
            case 'ArrowLeft':
                index--;
                break;
            case 'ArrowDown':
                index += columns;
                break;
            case 'ArrowUp':
                index -= columns;
                break;
            case 'Home':
                index = 0;
                break;
            case 'End':
                index = this.types.length - 1;
                break;
            default:
                return;
        }

        e.preventDefault();
        // The focus stays on the first and the last type, and on the first and the last row
        const item = this.types[index >= 0 && index < this.types.length ? index : current];
        this.setTypeTabStop(item);
        item.focus();
    }

    /** A press outside the dialog and its button closes the dialog. */
    private pressOutside = (e: Event) => {
        const path = e.composedPath();
        if (!path.includes(this.element) && !path.includes(this.button)) {
            this.close();
        }
    };
}
