import { createId } from './ids.js';

/** The element of a picker, which keeps its picker. */
export type PickerElement = HTMLElement & { picker?: Picker<unknown> };

/** An element of an option of a picker, with the key and the value of the option. */
type OptionElement = HTMLElement & { k?: string; v?: unknown };

/**
 * The options of a picker with options of type T. A type, not an interface, so that it fits the options of the public
 * PickerInstance.
 */
export type PickerOptions<T> = {
    /** The options, as a list or as an object whose keys are the keys of the options. */
    data?: T[] | Record<string, T> | null;
    /** The key of the option that the picker shows first. */
    value?: number | string;
    /**
     * The HTML of an option, or an element.
     * @param option - The option.
     * @param element - The element of the option in the list, or null for the option that the picker shows.
     */
    render?: ((option: T, element: HTMLElement | null) => string | Element) | null;
    /**
     * Called when the user chooses an option.
     * @param element - The element of the picker.
     * @param picker - The picker.
     * @param value - The option.
     * @param value2 - The option again.
     * @param key - The key of the option, as text.
     * @param e - The click or the key that chose it.
     */
    onchange?: ((element: HTMLElement, picker: Picker<T>, value: T, value2: T, key: string, e: Event) => void) | null;
    onopen?: ((element: HTMLElement, picker: Picker<T>) => void) | null;
    onclose?: ((element: HTMLElement, picker: Picker<T>) => void) | null;
    /** Called once the picker is created, to add controls to its list. */
    onload?: ((element: HTMLElement, picker: Picker<T>) => void) | null;
    /** The width of the picker, such as '120px'. */
    width?: number | string | null;
    /** An icon (from material icons) that the picker shows in place of the chosen option. */
    content?: string | false;
    /** Show the options in this number of columns. */
    columns?: number | null;
    /** Align the list with the right side of the picker. */
    right?: boolean;
    /** Open the list above the picker. */
    bottom?: boolean;
};

/**
 * A picker: a button that shows the chosen option and opens a list of options, with the ARIA combobox pattern. A press
 * opens it, and so do Enter, Space and the up and down arrows. The arrows, Home and End move through the options, which
 * the picker names in aria-activedescendant, Enter and Space choose one, and Escape or Tab closes the list. A press
 * outside it closes it too.
 *
 * The toolbar uses it for its select items. The element keeps the picker as its picker property, and has the classes of
 * the jSuites picker.
 */
export class Picker<T = string> {
    /** The type, as jSuites names its components. */
    readonly type = 'picker';

    /** The element of the picker. */
    readonly element: HTMLElement & { picker?: Picker<T> };

    /** The options of the picker, with their defaults. */
    readonly options: PickerOptions<T> & { value: number | string };

    // The element that shows the chosen option
    private readonly header: HTMLElement;

    // The list of options
    private readonly content: HTMLElement;

    constructor(element: HTMLElement & { picker?: Picker<T> }, options: PickerOptions<T>) {
        this.element = element;
        this.options = {
            value: 0,
            data: null,
            render: null,
            onchange: null,
            onopen: null,
            onclose: null,
            onload: null,
            width: null,
            content: false,
            columns: null,
            right: false,
            bottom: false,
        };
        for (const key of Object.keys(this.options) as (keyof PickerOptions<T>)[]) {
            if (Object.prototype.hasOwnProperty.call(options, key)) {
                (this.options as Record<string, unknown>)[key] = options[key];
            }
        }
        if (isNaN(parseInt(this.options.value))) {
            this.options.value = 0;
        }

        const id = createId('jss-picker');

        element.classList.add('jpicker');
        element.setAttribute('role', 'combobox');
        element.setAttribute('aria-haspopup', 'listbox');
        element.setAttribute('aria-expanded', 'false');
        element.setAttribute('aria-controls', id);
        element.setAttribute('tabindex', '0');
        element.addEventListener('mousedown', () => {
            if (!this.isOpened()) {
                this.open();
            }
        });
        element.addEventListener('keydown', this.keydown);

        this.header = document.createElement('div');
        this.header.classList.add('jpicker-header');
        if (this.options.width) {
            this.header.style.width = `${parseInt(this.options.width)}px`;
        }

        this.content = document.createElement('div');
        this.content.id = id;
        this.content.setAttribute('role', 'listbox');
        this.content.classList.add('jpicker-content');
        if (this.options.columns! > 0) {
            this.content.classList.add('jpicker-columns');
            this.content.style.width = this.options.width ? `${this.options.width}` : `${36 * this.options.columns!}px`;
        }
        this.content.addEventListener('click', (e) => {
            const option = (e.target as Element).closest<OptionElement>('.jpicker-item');
            // The options of a picker in the list, such as those of the borders, are not options of this picker
            if (option?.parentNode === this.content) {
                this.setValue(option.k!, e);
            }
        });

        element.appendChild(this.header);
        element.appendChild(this.content);

        this.createOptions();
        this.setValue(this.options.value);

        element.picker = this;

        this.options.onload?.(element, this);
    }

    /** The key of the chosen option: the first value, or the key as text after the user chose one. */
    getValue() {
        return this.options.value;
    }

    /**
     * Choose the option with a key. With an event, the picker calls its onchange.
     * @param key - The key of the option.
     * @param e - The click or the key that chose it.
     */
    setValue(key: number | string, e?: Event) {
        this.setLabel(key);
        this.options.value = String(key);

        this.close();

        if (e) {
            const value = this.getData()[key];
            this.options.onchange?.(this.element, this, value, value, String(key), e);
        }
    }

    /** Tell whether the list is open. */
    isOpened() {
        return this.element.classList.contains('jpicker-focus');
    }

    /** Open the list, below the picker or above it when there is no room, and inside the window. */
    open() {
        if (this.isOpened()) {
            return;
        }

        this.element.classList.add('jpicker-focus');
        this.element.setAttribute('aria-expanded', 'true');
        this.element.focus();
        document.addEventListener('mousedown', this.pressOutside, true);
        document.addEventListener('touchstart', this.pressOutside, true);

        this.content.style.marginLeft = '';

        const rectHeader = this.header.getBoundingClientRect();
        const rectContent = this.content.getBoundingClientRect();

        let top: number;
        if (window.innerHeight < rectHeader.bottom + rectContent.height || this.options.bottom) {
            top = -1 * (rectContent.height + 4);
        } else {
            top = rectHeader.height + 4;
        }

        let left = 0;
        if (this.options.right === true) {
            left = -1 * rectContent.width + rectHeader.width;
        }
        if (rectContent.left + left < 0) {
            left = left + rectContent.left + 10;
        }
        if (rectContent.left + rectContent.width > window.innerWidth) {
            left = -1 * (10 + rectContent.left + rectContent.width - window.innerWidth);
        }

        this.content.style.marginTop = `${parseInt(top)}px`;
        this.content.style.marginLeft = `${parseInt(left)}px`;

        this.options.onopen?.(this.element, this);
    }

    /** Close the list. */
    close() {
        document.removeEventListener('mousedown', this.pressOutside, true);
        document.removeEventListener('touchstart', this.pressOutside, true);

        if (!this.isOpened()) {
            return;
        }

        this.element.classList.remove('jpicker-focus');
        this.element.setAttribute('aria-expanded', 'false');
        this.element.removeAttribute('aria-activedescendant');
        for (const option of this.getOptionElements()) {
            option.classList.remove('jss_picker_active');
        }

        this.options.onclose?.(this.element, this);
    }

    /** The options as an object, whose keys are the keys of the options. */
    private getData() {
        return (this.options.data || {}) as Record<string, T>;
    }

    /** The elements of the options. */
    private getOptionElements() {
        return Array.from(this.content.children).filter((child): child is OptionElement => child.classList.contains('jpicker-item'));
    }

    /**
     * The label of an option: what render gives, or the option as text.
     * @param key - The key of the option.
     * @param element - The element of the option in the list, or null for the option that the picker shows.
     */
    private getLabel(key: number | string, element: HTMLElement | null): string | Element {
        const option = this.getData()[key] ?? null;
        if (typeof this.options.render === 'function') {
            return this.options.render(option as T, element);
        }
        return option == null ? '' : String(option);
    }

    /**
     * Put a label in an element. render gives HTML or an element, and an option without render is text.
     * @param element - The element.
     * @param label - The label.
     * @param html - The label is HTML.
     */
    private static putLabel(element: HTMLElement, label: string | Element, html: boolean) {
        if (label instanceof Element) {
            element.textContent = '';
            element.appendChild(label);
        } else if (html) {
            element.innerHTML = label;
        } else {
            element.textContent = label;
        }
    }

    /** Create the elements of the options. */
    private createOptions() {
        this.content.textContent = '';

        const html = typeof this.options.render === 'function';
        for (const [i, [key, value]] of Object.entries(this.getData()).entries()) {
            const option: OptionElement = document.createElement('div');
            option.classList.add('jpicker-item');
            option.setAttribute('role', 'option');
            option.id = `${this.content.id}-${i}`;
            option.k = key;
            option.v = value;
            Picker.putLabel(option, this.getLabel(key, option), html);
            this.content.appendChild(option);
        }
    }

    /**
     * Show an option in the picker, or its icon.
     * @param key - The key of the option.
     */
    private setLabel(key: number | string) {
        if (this.options.content) {
            const icon = document.createElement('i');
            icon.textContent = this.options.content;
            icon.classList.add('material-icons');
            Picker.putLabel(this.header, icon, false);
        } else {
            Picker.putLabel(this.header, this.getLabel(key, null), typeof this.options.render === 'function');
        }

        for (const option of this.getOptionElements()) {
            option.setAttribute('aria-selected', option.k === String(key) ? 'true' : 'false');
        }
    }

    /**
     * Mark the option that the keys are on.
     * @param index - The index of the option.
     */
    private setActiveOption(index: number) {
        const options = this.getOptionElements();
        for (const [i, option] of options.entries()) {
            option.classList.toggle('jss_picker_active', i === index);
        }

        this.element.setAttribute('aria-activedescendant', options[index].id);
        options[index].scrollIntoView?.({ block: 'nearest' });
    }

    /** The keys on the picker. Only the keys on the picker itself, not on the controls in its list, such as those of the borders. */
    private keydown = (e: KeyboardEvent) => {
        if (e.target !== this.element) {
            return;
        }

        const options = this.getOptionElements();
        if (!options.length) {
            return;
        }

        if (!this.isOpened()) {
            if (['Enter', ' ', 'ArrowDown', 'ArrowUp'].includes(e.key)) {
                e.preventDefault();
                this.open();

                // Start on the chosen option
                const index = options.findIndex((option) => option.k === `${this.getValue()}`);
                this.setActiveOption(index === -1 ? 0 : index);
            }

            return;
        }

        const active = this.element.getAttribute('aria-activedescendant');
        let index = options.findIndex((option) => option.id === active);

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
                // A click chooses the option, as a press does, and gives onchange the click
                options[index]?.click();
                return;
            case 'Escape':
                e.preventDefault();
                this.close();
                return;
            case 'Tab':
                this.close();
                return;
            default:
                return;
        }

        e.preventDefault();
        this.setActiveOption(index);
    };

    /** A press outside the picker closes it. */
    private pressOutside = (e: Event) => {
        if (!e.composedPath().includes(this.element)) {
            this.close();
        }
    };
}
