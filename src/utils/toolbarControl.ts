import { translate } from './dictionary.js';
import { Picker } from './picker.js';
import type jspreadsheet from '../jspreadsheet';

/** The element of an item of the toolbar, with the updateState and the style property of its item. */
type ItemElement = HTMLDivElement & { updateState?: jspreadsheet.ToolbarItemBase['updateState']; k?: unknown };

/**
 * The toolbar of a spreadsheet, with the WAI-ARIA toolbar pattern. Its items are buttons with an icon, pickers, labels
 * and lines between them. The toolbar is one tab stop: Tab goes to the item that had the focus last, the arrows, Home and
 * End move between the items, and Enter and Space press a button.
 *
 * The element keeps the toolbar as its toolbar property, and the toolbar has the classes of the jSuites toolbar.
 */
export class ToolbarControl {
    /** The type, as jSuites names its components. */
    readonly type = 'toolbar';

    /** The element of the toolbar. */
    readonly element: HTMLDivElement & { toolbar?: ToolbarControl };

    /** The settings of the toolbar, with its items. A type, not an interface, so that it fits the public ToolbarInstance. */
    readonly options: { items: jspreadsheet.ToolbarItem[]; [property: string]: unknown };

    // The element that holds the items
    private readonly content: HTMLDivElement;

    constructor(element: HTMLDivElement & { toolbar?: ToolbarControl }, options: jspreadsheet.ToolbarOptions) {
        this.element = element;
        this.options = { ...options };

        element.classList.add('jtoolbar');
        element.setAttribute('role', 'toolbar');
        element.setAttribute('aria-label', translate('Toolbar'));
        element.textContent = '';

        this.content = document.createElement('div');
        element.appendChild(this.content);

        element.addEventListener('click', (e) => {
            const item = (e.target as Element).closest<HTMLElement>('.jtoolbar-item');
            if (item?.parentNode === this.content) {
                this.selectItem(item);
            }
        });

        // The item that gets the focus, also from a press, is the tab stop
        element.addEventListener('focusin', (e) => {
            const target = e.target as HTMLElement;
            const items = this.getFocusableItems();

            if (items.includes(target)) {
                for (const item of items) {
                    item.setAttribute('tabindex', item === target ? '0' : '-1');
                }
            }
        });

        element.addEventListener('keydown', this.keydown);

        this.create(this.options.items);

        element.toolbar = this;
    }

    /**
     * Replace the items of the toolbar.
     * @param items - The items.
     */
    create(items: jspreadsheet.ToolbarItem[]) {
        this.options.items = items;
        this.content.textContent = '';

        for (const item of items as jspreadsheet.ToolbarItemBase[]) {
            this.content.appendChild(this.createItem(item));
        }

        for (const [i, item] of this.getFocusableItems().entries()) {
            item.setAttribute('tabindex', i === 0 ? '0' : '-1');
        }
    }

    /**
     * Call updateState on each item.
     * @param a - The first argument of updateState after the item, the worksheet.
     * @param b - The second argument.
     */
    update(a?: unknown, b?: unknown) {
        for (const item of this.content.children as HTMLCollectionOf<ItemElement>) {
            if (typeof item.updateState === 'function') {
                // updateState gets the worksheet, which the toolbar does not know
                (item.updateState as (...args: unknown[]) => void)(this.element, this, item, a, b);
            }
        }
    }

    /**
     * Mark an item as selected.
     * @param element - The element of the item.
     */
    selectItem(element: HTMLElement) {
        for (const item of this.content.children) {
            if (item !== element) {
                item.classList.remove('jtoolbar-selected');
            }
        }
        element.classList.add('jtoolbar-selected');
    }

    /**
     * Disable or enable the toolbar.
     * @param state - Disable it.
     */
    setReadonly(state: boolean) {
        this.element.classList.toggle('jtoolbar-disabled', !!state);
    }

    /** The element of the toolbar. */
    get() {
        return this.element;
    }

    show() {
        this.element.style.display = '';
    }

    hide() {
        this.element.style.display = 'none';
    }

    /** Remove the items of the toolbar. */
    destroy() {
        this.content.textContent = '';
    }

    /**
     * The element of an item.
     * @param item - The item.
     */
    private createItem(item: jspreadsheet.ToolbarItemBase) {
        const element: ItemElement = document.createElement('div');
        element.classList.add('jtoolbar-item');

        if (item.width) {
            element.style.width = `${parseInt(item.width as string)}px`;
        }
        if (item.k) {
            element.k = item.k;
        }
        if (item.tooltip) {
            element.setAttribute('title', item.tooltip);
            element.setAttribute('aria-label', item.tooltip);
        }
        if (item.id) {
            element.id = item.id as string;
        }
        if (item.updateState) {
            element.updateState = item.updateState;
        }
        if (item.active) {
            element.classList.add('jtoolbar-active');
        }
        if (item.disabled) {
            element.classList.add('jtoolbar-disabled');
            element.setAttribute('aria-disabled', 'true');
        }

        if (item.type === 'select' || item.type === 'dropdown') {
            // The picker reads its options from the item
            new Picker(element, item as ConstructorParameters<typeof Picker>[1]);
        } else if (item.type === 'divisor') {
            element.classList.add('jtoolbar-divisor');
            element.setAttribute('role', 'separator');
        } else if (item.type === 'label') {
            element.classList.add('jtoolbar-label');
            element.textContent = (item.content as string) || '';
        } else {
            const icon = document.createElement('i');
            if (typeof item.class === 'string') {
                icon.classList.add(...item.class.split(' ').filter(Boolean));
            } else {
                icon.classList.add('material-icons');
            }
            // The text of a material icon is its name, which is not read when the item has a tooltip as its name
            if (item.tooltip) {
                icon.setAttribute('aria-hidden', 'true');
            }
            icon.textContent = (item.content as string) || '';
            element.setAttribute('role', 'button');
            element.appendChild(icon);

            if (item.title) {
                element.setAttribute('title', item.title as string);
            }

            if (typeof item.render === 'function') {
                item.render(element, item);
            }
        }

        if (typeof item.onclick === 'function') {
            const onclick = item.onclick as (...args: unknown[]) => void;
            element.addEventListener('click', (e) => onclick.call(item, this.element, this, element, e));
        }

        return element;
    }

    /** The items of the toolbar that take the focus: the buttons and the pickers. */
    private getFocusableItems() {
        return Array.from(this.content.querySelectorAll<HTMLElement>(':scope > .jtoolbar-item[role="button"], :scope > .jtoolbar-item.jpicker'));
    }

    /** Handle a key on an item of the toolbar. */
    private keydown = (e: KeyboardEvent) => {
        const target = e.target as HTMLElement;
        const items = this.getFocusableItems();
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
    };
}
