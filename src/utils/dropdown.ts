import { createId } from './ids.js';

/** An item of a dropdown: a text or a number, which is its value and its label, or an object with an id and a name. */
export type DropdownItem =
    | string
    | number
    | {
          id: string | number;
          name: string;
          /** A second line under the name. */
          title?: string;
          /** The address of an image before the name. */
          image?: string;
          /** The name of the group that the item is in. */
          group?: string;
      };

/** The element of a dropdown, which keeps its dropdown. */
export type DropdownElement = HTMLElement & { dropdown?: Dropdown };

/** The options of a dropdown. */
export type DropdownOptions = {
    /** The items. */
    data?: DropdownItem[] | null;
    /** The value of the item that is chosen, or the values of the items, as a list or as a text with ; between them. */
    value?: string | number | (string | number)[] | null;
    /** Allow more than one item to be chosen. */
    multiple?: boolean;
    /** Filter the items with the text that is typed in the header. */
    autocomplete?: boolean;
    /** Open the dropdown when it is created. */
    opened?: boolean;
    /** The width of the dropdown, such as '100%'. */
    width?: string | number | null;
    /** Give the list a fixed position, for a dropdown in an element that scrolls. */
    position?: boolean;
    /** The text that the header shows when no item is chosen. */
    placeholder?: string;
    onopen?: ((element: DropdownElement, dropdown: Dropdown) => void) | null;
    /** Called when the list closes, after a press outside it, Enter, Tab or the choice of an item. Escape does not call it. */
    onclose?: ((element: DropdownElement, dropdown: Dropdown) => void) | null;
    /** Called after Escape, which puts back the value that the dropdown had when it opened and closes the list. */
    oncancel?: ((element: DropdownElement, dropdown: Dropdown) => void) | null;
};

/** An item of a dropdown, with its element. */
type Item = {
    value: string;
    text: string;
    element: HTMLElement;
};

/**
 * A dropdown: a header that shows the chosen items and opens a list of items, with the ARIA combobox pattern. The focus
 * stays in the header, and aria-activedescendant names the active item. A press on the header opens the list, and so do
 * the up and down arrows. The arrows, Page Up, Page Down, Home and End move through the items, a letter moves to the
 * next item that starts with it, Enter and Tab choose the active item and close the list, Escape closes it without a
 * change, and a press outside closes it. A press on an item chooses it, and a press on the chosen item chooses none.
 *
 * With multiple, a press, Enter or Space adds or removes an item, and the list stays open until Enter without an active
 * item, Tab or a press outside. With autocomplete, the header is an input whose text filters the items, and Space types
 * a space.
 *
 * The editor of a dropdown column and the filters of the columns use it. The element keeps the dropdown as its
 * dropdown property, and the dropdown has the classes of the jSuites dropdown.
 */
export class Dropdown {
    /** The type, as jSuites names its components. */
    readonly type = 'dropdown';

    /** The element of the dropdown. */
    readonly element: DropdownElement;

    /** The options of the dropdown, with their defaults. */
    readonly options: Required<Omit<DropdownOptions, 'data' | 'value' | 'width'>> & { data: DropdownItem[]; width: string | number | null };

    /** The input that shows the chosen items, and the text that filters them with autocomplete. */
    readonly header: HTMLInputElement;

    // The element that holds the list, and that has its position
    private readonly container: HTMLElement;

    // The list
    private readonly content: HTMLElement;

    // The items, in the order of the list
    private items: Item[] = [];

    // The chosen values, in the order they were chosen, with their labels. A value without an item is its own label.
    private readonly values = new Map<string, string>();

    // The values when the list opened, which Escape puts back
    private openValues: string[] = [];

    // The item that the keys are on
    private cursor: Item | null = null;

    // The text that filters the items
    private search = '';

    constructor(element: DropdownElement, options: DropdownOptions = {}) {
        this.element = element;
        this.options = {
            data: options.data ?? [],
            multiple: options.multiple ?? false,
            autocomplete: options.autocomplete ?? false,
            opened: options.opened ?? false,
            width: options.width ?? null,
            position: options.position ?? false,
            placeholder: options.placeholder ?? '',
            onopen: options.onopen ?? null,
            onclose: options.onclose ?? null,
            oncancel: options.oncancel ?? null,
        };

        const id = createId('jss-dropdown');

        element.classList.add('jdropdown', 'jdropdown-default');
        if (this.options.width) {
            element.style.width = `${this.options.width}`;
            element.style.minWidth = `${this.options.width}`;
        }

        const containerHeader = document.createElement('div');
        containerHeader.className = 'jdropdown-container-header';

        this.header = document.createElement('input');
        this.header.className = 'jdropdown-header jss_object';
        this.header.type = 'text';
        this.header.setAttribute('autocomplete', 'off');
        this.header.setAttribute('role', 'combobox');
        this.header.setAttribute('aria-haspopup', 'listbox');
        this.header.setAttribute('aria-expanded', 'false');
        this.header.setAttribute('aria-controls', id);
        if (this.options.autocomplete) {
            this.header.setAttribute('aria-autocomplete', 'list');
        } else {
            this.header.readOnly = true;
        }
        if (this.options.placeholder) {
            this.header.placeholder = this.options.placeholder;
        }
        this.header.addEventListener('click', this.pressHeader);
        this.header.addEventListener('keydown', this.keydown);
        this.header.addEventListener('input', () => {
            this.find(this.header.value);
            this.open();
        });
        containerHeader.appendChild(this.header);

        this.container = document.createElement('div');
        this.container.className = 'jdropdown-container';

        this.content = document.createElement('div');
        this.content.className = 'jdropdown-content';
        this.content.id = id;
        this.content.setAttribute('role', 'listbox');
        if (this.options.multiple) {
            this.content.setAttribute('aria-multiselectable', 'true');
        }
        // The focus stays in the header
        this.content.addEventListener('mousedown', (e) => e.preventDefault());
        this.content.addEventListener('click', (e) => {
            const option = (e.target as Element).closest('.jdropdown-item');
            const item = this.items.find((item) => item.element === option);
            if (item) {
                this.selectItem(item);
            }
        });
        // The list scrolls, not the worksheet
        this.content.addEventListener('wheel', (e) => e.stopPropagation());
        this.container.appendChild(this.content);

        element.appendChild(containerHeader);
        element.appendChild(this.container);

        this.createItems();
        this.setValue(options.value ?? null);

        element.dropdown = this;

        if (this.options.opened) {
            this.open();
        }
    }

    /**
     * The chosen values.
     * @param asArray - Give a list, not a text with ; between the values.
     */
    getValue(asArray: true): string[];
    getValue(asArray?: false): string;
    getValue(asArray?: boolean): string | string[];
    getValue(asArray?: boolean) {
        const values = Array.from(this.values.keys());
        return asArray ? values : values.join(';');
    }

    /**
     * The labels of the chosen items.
     * @param asArray - Give a list, not a text with ; and a space between the labels.
     */
    getText(asArray: true): string[];
    getText(asArray?: false): string;
    getText(asArray?: boolean): string | string[];
    getText(asArray?: boolean) {
        const labels = Array.from(this.values.values());
        return asArray ? labels : labels.join('; ');
    }

    /**
     * Choose the items with these values, in place of the chosen items.
     * @param value - A value, the values as a list, or a text with ; between them. Null or an empty text chooses none.
     */
    setValue(value: string | number | (string | number)[] | null | undefined) {
        this.values.clear();

        if (value != null && value !== '') {
            for (const v of Array.isArray(value) ? value : `${value}`.split(';')) {
                this.values.set(`${v}`, `${v}`);
            }
        }

        for (const item of this.items) {
            const selected = this.values.has(item.value);
            if (selected) {
                this.values.set(item.value, item.text);
            }
            this.markItem(item, selected);
        }

        if (!this.isOpened() || !this.options.autocomplete) {
            this.header.value = this.getText();
        }
    }

    /** Tell whether the list is open. */
    isOpened() {
        return this.element.classList.contains('jdropdown-focus');
    }

    /** Open the list below the header, or above it when there is no room, and move the keys to the first chosen item. */
    open() {
        if (this.isOpened()) {
            return;
        }

        this.element.classList.add('jdropdown-focus');
        this.header.setAttribute('aria-expanded', 'true');
        document.addEventListener('mousedown', this.pressOutside, true);
        document.addEventListener('touchstart', this.pressOutside, true);

        this.openValues = this.getValue(true);

        if (this.options.autocomplete) {
            this.header.value = this.search;
        }

        const first = this.items.find((item) => this.values.has(item.value));
        this.setCursor(first && !first.element.hidden ? first : null);

        const rect = this.element.getBoundingClientRect();
        const rectContainer = this.container.getBoundingClientRect();
        if (this.options.position) {
            this.container.style.position = 'fixed';
            if (window.innerHeight < rect.bottom + rectContainer.height) {
                this.container.style.top = '';
                this.container.style.bottom = `${window.innerHeight - rect.top + 1}px`;
            } else {
                this.container.style.top = `${rect.bottom}px`;
                this.container.style.bottom = '';
            }
            this.container.style.left = `${rect.left}px`;
        } else {
            this.container.style.top = '';
            this.container.style.bottom = window.innerHeight < rect.bottom + rectContainer.height ? `${rect.height + 1}px` : '';
        }
        this.container.style.minWidth = `${rect.width}px`;
        this.updateContent();

        this.header.focus({ preventScroll: true });

        this.options.onopen?.(this.element, this);
    }

    /**
     * Close the list.
     * @param ignoreEvents - Do not call onclose.
     * @returns The chosen values, as a text with ; between them.
     */
    close(ignoreEvents?: boolean) {
        document.removeEventListener('mousedown', this.pressOutside, true);
        document.removeEventListener('touchstart', this.pressOutside, true);

        if (this.isOpened()) {
            this.element.classList.remove('jdropdown-focus');
            this.header.setAttribute('aria-expanded', 'false');
            this.header.value = this.getText();
            this.setCursor(null);

            if (!ignoreEvents) {
                this.options.onclose?.(this.element, this);
            }
        }

        return this.getValue();
    }

    /** Choose no item. */
    reset() {
        this.setValue(null);
    }

    /**
     * Show only the items whose label, title or group has a text, and the chosen items.
     * @param text - The text. An empty text shows all the items.
     */
    find(text: string) {
        this.search = text;
        const query = text.trim().toLowerCase();

        for (const item of this.items) {
            const data = this.getItemData(item);
            const match =
                !query ||
                this.values.has(item.value) ||
                [item.text, data?.title, data?.group].some((label) => label != null && `${label}`.toLowerCase().includes(query));
            item.element.hidden = !match;
        }

        for (const group of this.content.querySelectorAll<HTMLElement>('.jdropdown-group')) {
            group.hidden = !group.querySelector('.jdropdown-item:not([hidden])');
        }

        this.setCursor(null);
        this.updateContent();
    }

    /**
     * Choose an item. Without multiple, it closes the list, and a press on the chosen item chooses none. With multiple, it
     * adds the item or removes it.
     * @param item - The item.
     */
    private selectItem(item: Item) {
        this.setCursor(item);

        if (this.options.multiple) {
            if (this.values.has(item.value)) {
                this.values.delete(item.value);
                this.markItem(item, false);
            } else {
                this.values.set(item.value, item.text);
                this.markItem(item, true);
            }
            if (!this.options.autocomplete) {
                this.header.value = this.getText();
            }
        } else {
            this.setValue(this.values.has(item.value) ? null : item.value);
            this.close();
        }
    }

    /** The item of the data that an element of the list shows, when it is an object. */
    private getItemData(item: Item) {
        const data = this.options.data[this.items.indexOf(item)];
        return typeof data === 'object' ? data : null;
    }

    /** Mark an item as chosen or not. */
    private markItem(item: Item, selected: boolean) {
        item.element.classList.toggle('jdropdown-selected', selected);
        item.element.setAttribute('aria-selected', selected ? 'true' : 'false');
    }

    /** Create the elements of the items: first the groups, with their items, then the items without a group. */
    private createItems() {
        const groups = new Map<string, HTMLElement>();
        const ungrouped: HTMLElement[] = [];
        const ordered: DropdownItem[] = [];

        // The items of the groups come first, as the list shows them
        for (const data of this.options.data) {
            if (typeof data === 'object' && data.group) {
                ordered.push(data);
            }
        }
        for (const data of this.options.data) {
            if (typeof data !== 'object' || !data.group) {
                ordered.push(data);
            }
        }
        this.options.data = ordered;

        for (const [i, data] of ordered.entries()) {
            const value = typeof data === 'object' ? `${data.id}` : `${data}`;
            const text = typeof data === 'object' ? `${data.name ?? ''}` : `${data}`;

            const element = document.createElement('div');
            element.className = 'jdropdown-item';
            element.id = `${this.content.id}-${i}`;
            element.setAttribute('role', 'option');
            element.setAttribute('aria-selected', 'false');

            if (typeof data === 'object' && data.image) {
                const image = document.createElement('img');
                image.className = 'jdropdown-image';
                image.src = data.image;
                image.alt = '';
                if (!data.title) {
                    image.classList.add('jdropdown-image-small');
                }
                element.appendChild(image);
            }

            const description = document.createElement('div');
            description.className = 'jdropdown-description';
            // An item without a label keeps its height
            description.textContent = text || ' ';
            if (typeof data === 'object' && data.title) {
                const title = document.createElement('div');
                title.className = 'jdropdown-title';
                title.textContent = data.title;
                description.appendChild(title);
            }
            element.appendChild(description);

            this.items.push({ value, text, element });

            if (typeof data === 'object' && data.group) {
                let group = groups.get(data.group);
                if (!group) {
                    const container = document.createElement('div');
                    container.className = 'jdropdown-group';
                    container.setAttribute('role', 'group');
                    const name = document.createElement('div');
                    name.className = 'jdropdown-group-name';
                    name.id = `${this.content.id}-group-${groups.size}`;
                    name.textContent = data.group;
                    container.setAttribute('aria-labelledby', name.id);
                    group = document.createElement('div');
                    group.className = 'jdropdown-group-items';
                    container.appendChild(name);
                    container.appendChild(group);
                    this.content.appendChild(container);
                    groups.set(data.group, group);
                }
                group.appendChild(element);
            } else {
                ungrouped.push(element);
            }
        }

        for (const element of ungrouped) {
            this.content.appendChild(element);
        }
    }

    /** Hide the list when no item is shown. */
    private updateContent() {
        this.content.style.display = this.getVisibleItems().length ? '' : 'none';
    }

    /** The items that the list shows, in its order. */
    private getVisibleItems() {
        return this.items.filter((item) => !item.element.hidden);
    }

    /**
     * Move the keys to an item.
     * @param item - The item, or null for none.
     */
    private setCursor(item: Item | null) {
        this.cursor?.element.classList.remove('jdropdown-cursor');
        this.cursor = item;

        if (item) {
            item.element.classList.add('jdropdown-cursor');
            this.header.setAttribute('aria-activedescendant', item.element.id);
            item.element.scrollIntoView?.({ block: 'nearest' });
        } else {
            this.header.removeAttribute('aria-activedescendant');
        }
    }

    /**
     * Move the keys by a number of items, from the active item. Without an active item, they go to the first item.
     * @param step - The number of items, negative to move up.
     */
    private moveCursor(step: number) {
        const items = this.getVisibleItems();
        if (!items.length) {
            return;
        }

        const index = this.cursor ? items.indexOf(this.cursor) : -1;
        if (index === -1) {
            this.setCursor(items[0]);
        } else {
            this.setCursor(items[Math.max(0, Math.min(items.length - 1, index + step))]);
        }
    }

    /**
     * Move the keys to the next item that starts with a letter.
     * @param letter - The letter.
     */
    private moveToLetter(letter: string) {
        const items = this.getVisibleItems();
        const start = this.cursor ? items.indexOf(this.cursor) + 1 : 0;
        const lower = letter.toLowerCase();

        for (const item of [...items.slice(start), ...items.slice(0, start)]) {
            if (item.text.toLowerCase().startsWith(lower)) {
                this.setCursor(item);
                return;
            }
        }
    }

    /** A press on the header opens the list, or closes it. With autocomplete, only a press on the arrow closes it. */
    private pressHeader = (e: MouseEvent) => {
        if (!this.isOpened()) {
            this.open();
        } else if (!this.options.autocomplete) {
            this.close();
        } else {
            const rect = this.header.getBoundingClientRect();
            if (rect.width - (e.clientX - rect.left) < 30) {
                this.close();
            }
        }
    };

    /** The keys on the header. */
    private keydown = (e: KeyboardEvent) => {
        if (!this.isOpened()) {
            if (e.key === 'ArrowDown' || e.key === 'ArrowUp') {
                e.preventDefault();
                this.open();
            }
            return;
        }

        switch (e.key) {
            case 'ArrowDown':
                this.moveCursor(1);
                break;
            case 'ArrowUp':
                this.moveCursor(-1);
                break;
            case 'PageDown':
                this.moveCursor(7);
                break;
            case 'PageUp':
                this.moveCursor(-7);
                break;
            case 'Home':
            case 'End':
                // With autocomplete, the keys also move in the text
                if (e.key === 'Home') {
                    this.setCursor(this.getVisibleItems()[0] ?? null);
                } else {
                    this.setCursor(this.getVisibleItems().at(-1) ?? null);
                }
                if (this.options.autocomplete) {
                    return;
                }
                break;
            case 'Enter':
                // The key chooses an item, and does not go to the worksheet, which would move the selection
                e.preventDefault();
                e.stopPropagation();
                // With multiple, it adds or removes the active item, and closes the list when no item is active. Without,
                // it chooses the active item, also when it is chosen, and closes the list.
                if (this.cursor && !this.cursor.element.hidden) {
                    if (this.options.multiple) {
                        this.selectItem(this.cursor);
                        return;
                    }
                    this.setValue(this.cursor.value);
                }
                this.close();
                return;
            case 'Escape':
                // The worksheet closes an editor without saving after it
                this.setValue(this.openValues);
                this.close(true);
                this.options.oncancel?.(this.element, this);
                return;
            case 'Tab':
                // Tab chooses the active item, as Enter, and moves the focus on
                if (!this.options.multiple && this.cursor && !this.cursor.element.hidden) {
                    this.setValue(this.cursor.value);
                }
                this.close();
                return;
            case ' ':
                // Space adds or removes the active item of a dropdown with multiple, and types a space with autocomplete
                if (this.options.autocomplete) {
                    return;
                }
                if (this.options.multiple && this.cursor && !this.cursor.element.hidden) {
                    this.selectItem(this.cursor);
                }
                break;
            default:
                if (!this.options.autocomplete && e.key.length === 1 && !e.ctrlKey && !e.metaKey && !e.altKey) {
                    this.moveToLetter(e.key);
                }
                return;
        }

        e.preventDefault();
    };

    /** A press outside the dropdown closes it. */
    private pressOutside = (e: Event) => {
        if (!e.composedPath().includes(this.element)) {
            this.close();
        }
    };
}
