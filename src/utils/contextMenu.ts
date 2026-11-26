import { translate } from './dictionary.js';
import type gridwell from '../gridwell';

/** Where a menu opens: the event that opens it, or a point of the window. */
type MenuPosition = MouseEvent | TouchEvent | { x: number; y: number };

/** An element of an item of a menu, with its item, so that a key or a press finds what to do. */
type ItemElement = HTMLDivElement & { menuItem?: gridwell.ContextMenuOption };

/**
 * Tell whether an item is a line between the items. The type check does not narrow the union with the two types of a line.
 * @param item - The item.
 */
function isLine(item: gridwell.ContextMenuItem): item is Exclude<gridwell.ContextMenuItem, gridwell.ContextMenuOption> {
    return item.type === 'line' || item.type === 'divisor';
}

/**
 * The context menu of a spreadsheet: a menu with the ARIA menu pattern, which the keyboard and the mouse use. It opens at
 * the pointer, or under the focused element when a key opens it, and moves the focus to its first item. The arrow keys,
 * Home and End move between the items, Enter and Space choose one, the right arrow opens a submenu and the left arrow
 * closes it, and Escape or Tab closes the menu. A press outside it or the wheel closes it too. When it closes, the focus
 * goes back to the element that had it.
 *
 * The element keeps the menu as its contextmenu property, and has the classes of the jSuites menu, so that the styles
 * and the code of applications that read them keep working.
 */
export class ContextMenu {
    // The menu that is open: one menu of the page is open at a time
    private static opened: ContextMenu | null = null;

    /** The type, as jSuites names its components. */
    readonly type = 'contextmenu';

    /** The element of the menu. */
    element: HTMLDivElement & { contextmenu?: ContextMenu };

    /** The items of the menu, as open got them. */
    options: { items: gridwell.ContextMenuItem[] | null } = { items: null };

    // The element that had the focus when the menu opened
    private previousFocus: HTMLElement | null = null;

    constructor(element: HTMLDivElement & { contextmenu?: ContextMenu }) {
        this.element = element;
        element.classList.add('jcontextmenu');
        element.setAttribute('role', 'menu');
        element.setAttribute('tabindex', '-1');
        element.addEventListener('keydown', this.keydown);
        element.contextmenu = this;
    }

    /**
     * Open the menu, with new items when they are given.
     * @param e - The event that opens the menu, or a point of the window.
     * @param items - The items.
     */
    open(e: MenuPosition, items?: gridwell.ContextMenuItem[]) {
        if (items) {
            this.options.items = items;
            this.create(items);
        }

        if (ContextMenu.opened && ContextMenu.opened !== this) {
            ContextMenu.opened.close();
        }
        ContextMenu.opened = this;

        if (!this.isOpened()) {
            const active = document.activeElement as HTMLElement | null;
            this.previousFocus = active && !this.element.contains(active) ? active : null;
            this.element.classList.add('jcontextmenu-focus');
            document.addEventListener('mousedown', this.pressOutside, true);
            document.addEventListener('touchstart', this.pressOutside, true);
            window.addEventListener('wheel', this.wheel, true);
        }

        this.place(e);
        this.focusItem(this.items(this.element)[0]);
    }

    /** Tell whether the menu is open. */
    isOpened() {
        return this.element.classList.contains('jcontextmenu-focus');
    }

    /**
     * Close the menu. The focus goes back to the element that had it when the menu opened, when it is in the menu.
     */
    close() {
        this.hide(true);
    }

    /**
     * Replace the items of the menu.
     * @param items - The items.
     */
    create(items: gridwell.ContextMenuItem[]) {
        this.element.textContent = '';
        this.element.setAttribute('aria-label', translate('Menu'));
        this.element.appendChild(this.createHeader());
        for (const item of items) {
            this.element.appendChild(this.createItem(item));
        }
    }

    /**
     * Hide the menu and stop listening to the document.
     * @param restoreFocus - Move the focus back to the element that had it, when the focus is in the menu.
     */
    private hide(restoreFocus: boolean) {
        document.removeEventListener('mousedown', this.pressOutside, true);
        document.removeEventListener('touchstart', this.pressOutside, true);
        window.removeEventListener('wheel', this.wheel, true);
        if (ContextMenu.opened === this) {
            ContextMenu.opened = null;
        }
        if (!this.isOpened()) {
            return;
        }

        const hadFocus = this.element.contains(document.activeElement);
        this.element.classList.remove('jcontextmenu-focus');
        for (const item of this.element.querySelectorAll('.jcontextmenu-expanded')) {
            this.collapse(item as ItemElement);
        }

        if (restoreFocus && hadFocus && this.previousFocus?.isConnected) {
            this.previousFocus.focus({ preventScroll: true });
        }
        this.previousFocus = null;
    }

    /**
     * Place the menu at the pointer, inside the window. A menu that a key opens has no pointer: it opens under the element
     * of the event.
     * @param e - The event, or a point.
     */
    private place(e: MenuPosition) {
        let x: number;
        let y: number;
        if ('target' in e && e.target) {
            const touch = 'changedTouches' in e ? e.changedTouches[0] : null;
            x = touch ? touch.clientX : (e as MouseEvent).clientX;
            y = touch ? touch.clientY : (e as MouseEvent).clientY;
            if (!touch && x === 0 && y === 0 && typeof (e.target as Element).getBoundingClientRect === 'function') {
                const rect = (e.target as Element).getBoundingClientRect();
                x = rect.left;
                y = rect.bottom;
            }
        } else {
            ({ x, y } = e as { x: number; y: number });
        }

        const rect = this.element.getBoundingClientRect();
        this.element.style.top = `${window.innerHeight < y + rect.height ? Math.max(0, y - rect.height) : y}px`;
        if (window.innerWidth < x + rect.width) {
            this.element.style.left = `${x - rect.width > 0 ? x - rect.width : 10}px`;
        } else {
            this.element.style.left = `${x}px`;
        }
    }

    /** The header of the menu, which a small screen shows with a button that closes the menu. */
    private createHeader() {
        const header = document.createElement('div');
        header.className = 'header';
        // The keys close the menu, and the header is not an item of the menu
        header.setAttribute('aria-hidden', 'true');
        header.addEventListener('click', (e) => {
            e.preventDefault();
            e.stopPropagation();
        });

        const title = document.createElement('a');
        title.className = 'title';
        title.textContent = translate('Menu');
        header.appendChild(title);

        const close = document.createElement('a');
        close.className = 'close';
        close.textContent = translate('close');
        close.addEventListener('click', () => this.close());
        header.appendChild(close);

        return header;
    }

    /**
     * The element of an item: a line, or an item with its title, its shortcut and its submenu.
     * @param item - The item.
     */
    private createItem(item: gridwell.ContextMenuItem): HTMLElement {
        if (isLine(item)) {
            const line = document.createElement('hr');
            line.setAttribute('role', 'separator');
            return line;
        }

        const element: ItemElement = document.createElement('div');
        element.menuItem = item;
        element.setAttribute('role', 'menuitem');
        element.setAttribute('tabindex', '-1');

        const title = document.createElement('a');
        title.textContent = item.title;
        element.appendChild(title);

        if (item.tooltip) {
            element.setAttribute('title', item.tooltip);
        }
        if (item.icon) {
            element.setAttribute('data-icon', item.icon);
        }
        if (item.id) {
            element.id = item.id;
        }
        if (item.disabled) {
            element.className = 'jcontextmenu-disabled';
            element.setAttribute('aria-disabled', 'true');
        }

        // A press keeps the focus in the menu, and the release chooses the item. The right button, which opened the menu,
        // does not.
        element.addEventListener('mousedown', (e) => e.preventDefault());
        element.addEventListener('mouseup', (e) => {
            if (e.button !== 2 && (e.target as Element).closest('[role="menuitem"]') === element) {
                this.choose(element, e);
            }
        });
        // The item under the pointer has the focus, so that one item is highlighted
        element.addEventListener('mouseenter', () => {
            if (this.isOpened()) {
                element.focus({ preventScroll: true });
                if (item.submenu) {
                    this.expand(element);
                }
            }
        });

        if (item.submenu) {
            const arrow = document.createElement('span');
            arrow.textContent = '►';
            arrow.setAttribute('aria-hidden', 'true');
            element.appendChild(arrow);
            element.classList.add('jcontexthassubmenu');
            element.setAttribute('aria-haspopup', 'menu');
            element.setAttribute('aria-expanded', 'false');

            const submenu = document.createElement('div');
            submenu.className = 'jcontextmenu';
            submenu.setAttribute('role', 'menu');
            submenu.setAttribute('aria-label', item.title);
            for (const child of item.submenu) {
                submenu.appendChild(this.createItem(child));
            }
            element.appendChild(submenu);
            element.addEventListener('mouseleave', () => this.collapse(element));
        } else if (item.shortcut) {
            const shortcut = document.createElement('span');
            shortcut.textContent = item.shortcut;
            element.appendChild(shortcut);
        }

        return element;
    }

    /**
     * Choose an item: open its submenu, or close the menu and call its onclick. A disabled item does nothing.
     * @param element - The element of the item.
     * @param e - The event that chooses it.
     */
    private choose(element: ItemElement, e: MouseEvent | KeyboardEvent) {
        const item = element.menuItem;
        if (!item || item.disabled) {
            return;
        }
        if (item.submenu) {
            this.expand(element);
            this.focusItem(this.items(element.lastElementChild as HTMLElement)[0]);
            return;
        }
        // The menu closes first, so that an item that moves the focus, for example to an editor, keeps it there
        this.close();
        item.onclick?.(element, e);
    }

    /**
     * Show the submenu of an item, on the side of the item where it fits.
     * @param element - The element of the item.
     */
    private expand(element: ItemElement) {
        const submenu = element.lastElementChild as HTMLElement;
        element.classList.add('jcontextmenu-expanded');
        element.setAttribute('aria-expanded', 'true');
        const rect = element.getBoundingClientRect();
        const width = submenu.getBoundingClientRect().width;
        if (rect.right + width > window.innerWidth && rect.left - width >= 0) {
            submenu.style.left = 'auto';
            submenu.style.right = '99%';
        } else {
            submenu.style.left = '';
            submenu.style.right = '';
        }
    }

    /**
     * Hide the submenu of an item.
     * @param element - The element of the item.
     */
    private collapse(element: ItemElement) {
        element.classList.remove('jcontextmenu-expanded');
        element.setAttribute('aria-expanded', 'false');
    }

    /**
     * The elements of the items of a menu, without its lines and the items of its submenus.
     * @param menu - The menu.
     */
    private items(menu: HTMLElement) {
        return [...menu.children].filter((child): child is ItemElement => child.getAttribute('role') === 'menuitem');
    }

    /**
     * Move the focus to an item.
     * @param element - The element of the item, or undefined for a menu without items, which keeps the focus.
     */
    private focusItem(element: HTMLElement | undefined) {
        (element ?? this.element).focus({ preventScroll: true });
    }

    /** The keys of the menu. */
    private keydown = (e: KeyboardEvent) => {
        const target = e.target as HTMLElement;
        const item = target.getAttribute('role') === 'menuitem' ? (target as ItemElement) : null;
        // The menu of the item, or the menu itself when it has the focus
        const menu = item ? (item.parentElement as HTMLElement) : this.element;
        const items = this.items(menu);
        const index = item ? items.indexOf(item) : -1;
        // The item of the submenu that has the focus
        const parent = menu === this.element ? null : (menu.parentElement as ItemElement);

        if (e.key === 'ArrowDown') {
            this.focusItem(items[(index + 1) % items.length]);
        } else if (e.key === 'ArrowUp') {
            this.focusItem(items[index <= 0 ? items.length - 1 : index - 1]);
        } else if (e.key === 'Home') {
            this.focusItem(items[0]);
        } else if (e.key === 'End') {
            this.focusItem(items[items.length - 1]);
        } else if (e.key === 'Enter' || e.key === ' ') {
            if (item) {
                this.choose(item, e);
            }
        } else if (e.key === 'ArrowRight') {
            if (item?.menuItem?.submenu && !item.menuItem.disabled) {
                this.choose(item, e);
            }
        } else if (e.key === 'ArrowLeft' || e.key === 'Escape') {
            if (parent) {
                this.collapse(parent);
                this.focusItem(parent);
            } else if (e.key === 'Escape') {
                this.close();
            } else {
                return;
            }
        } else if (e.key === 'Tab') {
            this.close();
        } else {
            return;
        }

        e.preventDefault();
        e.stopPropagation();
    };

    /** A press outside the menu closes it. The press moves the focus, so the menu does not move it. */
    private pressOutside = (e: Event) => {
        if (!e.composedPath().includes(this.element)) {
            this.hide(false);
        }
    };

    /** The wheel outside the menu closes it, as the menu does not move with the page. */
    private wheel = (e: Event) => {
        if (!e.composedPath().includes(this.element)) {
            this.close();
        }
    };
}
