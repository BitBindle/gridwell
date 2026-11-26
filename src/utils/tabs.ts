import jSuites from 'jsuites';
import { translate } from './dictionary.js';
import type jspreadsheet from '../jspreadsheet';

/** A tab header. It keeps a reference to its panel. */
type TabHeader = HTMLButtonElement & { content?: HTMLElement };

/** A handler of a click on a tab, or of a change of the open tab. */
type TabHandler = (el: HTMLElement, tabs: Tabs, index: number, header: TabHeader | null, content: Element) => void;

/** The options of the tabs, which factory.ts gives. A missing option takes its default in the constructor. */
export interface TabsOptions {
    /** The tabs to create when the tabs start. content is HTML. */
    data: { title: string; content: string }[];
    allowCreate: boolean;
    hideHeaders: boolean;
    animation: boolean;
    onclick: TabHandler | null;
    onchange: TabHandler | null;
    oncreate: ((el: HTMLElement & { tabs: Tabs }, content: HTMLDivElement | undefined) => void) | null;
    ondelete: ((el: HTMLElement, index: number) => void) | null;
    /** Returns the title of the new tab, or false to cancel. */
    onbeforecreate: ((el: HTMLElement, title?: string | null) => string | false) | null;
    onchangeposition: ((headers: HTMLDivElement, from: number, to: number) => void) | null;
    /** The labels option of the spreadsheet, with the accessible names of the tabs. */
    labels: NonNullable<jspreadsheet.SpreadsheetOptions['labels']>;
}

/** English accessible names of the tabs, which the labels option replaces and jspreadsheet.setDictionary can translate. */
const defaultLabels = {
    worksheets: 'Worksheets',
    tabPanels: 'Tab Panels',
    addWorksheet: 'Add a new worksheet',
};

/**
 * Worksheet tabs. CE uses only part of jSuites.tabs, and needs the keyboard and ARIA behaviour of the
 * WAI-ARIA tabs pattern for the tab headers: https://www.w3.org/WAI/ARIA/apg/patterns/tabs/
 * The object has the methods of jSuites.tabs, and the same class names, so the jSuites styles still apply.
 *
 * A spreadsheet with worksheets has one, as element.tabs: new Tabs(element, options) builds the tabs in the element.
 */
export class Tabs {
    /** The element of the spreadsheet, which holds the headers and the panels. */
    el: HTMLElement & { tabs?: Tabs };

    options: TabsOptions;

    /** The tab list: a header for each tab, then the border. */
    headers!: HTMLDivElement;

    /** The panels, one for each tab. */
    content!: HTMLDivElement;

    /** The buttons that scroll the headers, which init creates. */
    private prev!: HTMLDivElement;
    private next!: HTMLDivElement;

    /** The line under the open tab, after the last header. init creates it. */
    private border!: HTMLDivElement;

    constructor(el: HTMLElement & { tabs?: Tabs }, options: Partial<TabsOptions>) {
        this.el = el;
        // Each option is set below
        this.options = {} as TabsOptions;

        // Default configuration
        const defaults: TabsOptions = {
            data: [],
            allowCreate: false,
            hideHeaders: false,
            animation: false,
            onclick: null,
            onchange: null,
            oncreate: null,
            ondelete: null,
            onbeforecreate: null,
            onchangeposition: null,
            labels: {},
        };

        for (const property in defaults) {
            if (options && Object.prototype.hasOwnProperty.call(options, property)) {
                (this.options as unknown as Record<string, unknown>)[property] = (options as Record<string, unknown>)[property];
            } else {
                (this.options as unknown as Record<string, unknown>)[property] = (defaults as unknown as Record<string, unknown>)[property];
            }
        }

        this.init();

        el.tabs = this;
    }

    private getHeader(index: number | false) {
        const header = this.headers.children[index as number];

        if (header && header !== this.border) {
            return header as TabHeader;
        }

        return null;
    }

    private getIndex(element: HTMLElement) {
        while (element && element.parentNode !== this.headers) {
            element = element.parentNode as HTMLElement;
        }

        if (!element || element === this.border) {
            return -1;
        }

        return Array.from(this.headers.children).indexOf(element);
    }

    private updateControls(x: number) {
        if (typeof this.headers.scrollTo === 'function') {
            this.headers.scrollTo({
                left: x,
                behavior: 'smooth',
            });
        } else {
            this.headers.scrollLeft = x;
        }

        if (x <= 1) {
            this.prev.classList.add('disabled');
        } else {
            this.prev.classList.remove('disabled');
        }

        if (x >= this.headers.scrollWidth - this.headers.offsetWidth) {
            this.next.classList.add('disabled');
        } else {
            this.next.classList.remove('disabled');
        }

        if (this.headers.scrollWidth <= this.headers.offsetWidth) {
            this.prev.style.display = 'none';
            this.next.style.display = 'none';
        } else {
            this.prev.style.display = '';
            this.next.style.display = '';
        }
    }

    /**
     * Show the headers when the user can choose a worksheet or create one, unless the option hideHeaders is true
     */
    private updateHeaders() {
        const container = this.headers.parentNode as HTMLElement;

        if (this.options.hideHeaders === true || (this.content.children.length < 2 && this.options.allowCreate !== true)) {
            container.style.display = 'none';
        } else {
            container.style.display = '';

            const index = this.getActive();
            const header = this.getHeader(index);

            if (header) {
                const x1 = header.offsetLeft;
                const x2 = x1 + header.offsetWidth;
                const r1 = this.headers.scrollLeft;
                const r2 = r1 + this.headers.offsetWidth;

                if (!(r1 <= x1 && r2 >= x2)) {
                    // Out of the viewport
                    this.updateControls(x1 - 1);
                }

                this.setBorder(index);
            }
        }
    }

    setBorder(index: number | false) {
        if (this.options.animation) {
            setTimeout(() => {
                const header = this.getHeader(index);

                if (header) {
                    this.border.style.width = `${header.getBoundingClientRect().width}px`;
                    this.border.style.left = `${header.offsetLeft}px`;
                    this.border.style.bottom = '0px';
                }
            }, 50);
        }
    }

    open(index: number) {
        if (!this.content.children[index] || !this.getHeader(index)) {
            return;
        }

        index = parseInt(index);

        const previous = this.getActive();

        for (let i = 0; i < this.content.children.length; i++) {
            const header = this.getHeader(i);
            const selected = i === index;

            if (header) {
                if (selected) {
                    header.classList.add('jtabs-selected');
                } else {
                    header.classList.remove('jtabs-selected');
                }
                header.setAttribute('aria-selected', selected ? 'true' : 'false');
                // Roving tabindex: only the selected tab is in the tab sequence
                header.tabIndex = selected ? 0 : -1;
            }

            if (selected) {
                this.content.children[i].classList.add('jtabs-selected');
            } else {
                this.content.children[i].classList.remove('jtabs-selected');
            }
        }

        if (previous !== index && typeof this.options.onchange === 'function') {
            this.options.onchange(this.el, this, index, this.getHeader(index), this.content.children[index]);
        }

        this.updateHeaders();
    }

    selectIndex(a: Element) {
        const index = Array.from(this.headers.children).indexOf(a);

        if (index >= 0) {
            this.open(index);
        }

        return index;
    }

    getActive(): number | false {
        for (let i = 0; i < this.headers.children.length; i++) {
            if (this.headers.children[i].classList.contains('jtabs-selected')) {
                return i;
            }
        }

        return false;
    }

    rename(i: number, title?: string | null) {
        const header = this.getHeader(i);

        if (!header) {
            return;
        }

        if (!title) {
            title = prompt('New title', header.textContent);
        }

        if (title) {
            header.textContent = title;
        }

        this.setBorder(this.getActive());
    }

    create(title?: string): HTMLDivElement | false | undefined {
        if (typeof this.options.onbeforecreate === 'function') {
            const ret = this.options.onbeforecreate(this.el, title);

            if (ret === false) {
                return false;
            } else {
                title = ret;
            }
        }

        const div = this.appendElement(title);

        if (typeof this.options.oncreate === 'function') {
            // The constructor sets el.tabs before a tab can be created
            this.options.oncreate(this.el as HTMLElement & { tabs: Tabs }, div);
        }

        this.setBorder(this.getActive());

        return div;
    }

    remove(index: number): void | false {
        return this.deleteElement(index);
    }

    nextNumber() {
        let num = 0;

        for (let i = 0; i < this.content.children.length; i++) {
            // The match is compared and parsed as text
            const tmp = this.getHeader(i)!.textContent.match(/[0-9].*/) as unknown as string | null;

            if (Number(tmp) > num) {
                num = parseInt(tmp);
            }
        }

        if (!num) {
            num = 1;
        } else {
            num++;
        }

        return num;
    }

    deleteElement(index: number): void | false {
        const current = this.getActive();
        const header = this.getHeader(index);

        if (!header) {
            return false;
        }

        const hadFocus = header === document.activeElement;

        this.headers.removeChild(header);
        this.content.removeChild(this.content.children[index]);

        if (current === index) {
            this.open(0);
        } else {
            this.updateHeaders();
        }

        // Keep the focus in the tab list
        if (hadFocus && this.getHeader(this.getActive())) {
            this.getHeader(this.getActive())!.focus();
        }

        if (typeof this.options.ondelete === 'function') {
            this.options.ondelete(this.el, index);
        }
    }

    appendElement(
        title?: string | null,
        cb?: ((div: HTMLDivElement, header: TabHeader) => void) | null,
        openTab?: boolean,
        position?: number
    ): HTMLDivElement | undefined {
        if (!title) {
            title = prompt('Title?', '');
        }

        if (title) {
            const headerId = jSuites.guid();
            const contentId = jSuites.guid();

            // Add content
            const div = document.createElement('div');
            div.setAttribute('id', contentId);
            div.setAttribute('role', 'tabpanel');
            div.setAttribute('aria-labelledby', headerId);

            // Add header
            const h = document.createElement('button');
            h.type = 'button';
            h.setAttribute('id', headerId);
            h.setAttribute('role', 'tab');
            h.setAttribute('aria-selected', 'false');
            h.setAttribute('aria-controls', contentId);
            h.tabIndex = -1;
            h.textContent = title;
            (h as TabHeader).content = div;

            if (typeof position === 'undefined' || !this.content.children[position]) {
                this.content.appendChild(div);
                this.headers.insertBefore(h, this.border);
            } else {
                this.content.insertBefore(div, this.content.children[position]);
                this.headers.insertBefore(h, this.getHeader(position));
            }

            if (openTab !== false) {
                this.selectIndex(h);
            } else {
                this.updateHeaders();
            }

            if (typeof cb === 'function') {
                cb(div, h);
            }

            return div;
        }
    }

    updateContent(position: number, newContent: string | HTMLElement) {
        const current = this.content.children[position];

        if (!current) {
            return;
        }

        let contentItem: HTMLElement;

        if (typeof newContent !== 'string') {
            contentItem = newContent;
        } else {
            contentItem = document.createElement('div');
            contentItem.innerHTML = newContent;
        }

        // Keep the relation with the tab, without a change to the attributes that the caller gave
        if (!contentItem.id) {
            contentItem.setAttribute('id', current.getAttribute('id')!);
        }
        if (!contentItem.hasAttribute('role')) {
            contentItem.setAttribute('role', 'tabpanel');
        }
        if (!contentItem.hasAttribute('aria-labelledby')) {
            contentItem.setAttribute('aria-labelledby', current.getAttribute('aria-labelledby')!);
        }

        if (current.classList.contains('jtabs-selected')) {
            contentItem.classList.add('jtabs-selected');
        }

        this.content.replaceChild(contentItem, current);

        if (this.getHeader(position)) {
            this.getHeader(position)!.content = contentItem;
            this.getHeader(position)!.setAttribute('aria-controls', contentItem.id);
        }

        this.setBorder(this.getActive());
    }

    updatePosition(f: number, t: number, ignoreEvents?: boolean, openTab?: boolean) {
        // Update the position of the content
        if (f > t) {
            this.content.insertBefore(this.content.children[f], this.content.children[t]);
        } else {
            this.content.insertBefore(this.content.children[f], this.content.children[t].nextSibling);
        }

        if (openTab !== false) {
            this.open(t);
        } else {
            this.setBorder(this.getActive());
        }

        if (!ignoreEvents && typeof this.options.onchangeposition === 'function') {
            this.options.onchangeposition(this.headers, f, t);
        }
    }

    move(f: number, t: number, ignoreEvents?: boolean, openTab?: boolean) {
        if (f > t) {
            this.headers.insertBefore(this.getHeader(f)!, this.getHeader(t));
        } else {
            this.headers.insertBefore(this.getHeader(f)!, this.getHeader(t)!.nextSibling);
        }

        this.updatePosition(f, t, ignoreEvents, openTab);
    }

    /**
     * The text of an accessible name of the tabs.
     * @param key - The name in the labels option.
     */
    private getLabelText(key: keyof typeof defaultLabels) {
        const text = this.options.labels[key];

        return typeof text === 'string' ? text : translate(defaultLabels[key]);
    }

    init() {
        this.el.classList.add('jtabs');
        this.el.textContent = '';

        this.headers = document.createElement('div');
        this.headers.className = 'jtabs-headers';
        this.headers.setAttribute('role', 'tablist');
        this.headers.setAttribute('aria-label', this.getLabelText('worksheets'));

        this.content = document.createElement('div');
        this.content.className = 'jtabs-content';
        this.content.setAttribute('role', 'region');
        this.content.setAttribute('aria-label', this.getLabelText('tabPanels'));

        const header = document.createElement('div');
        header.className = 'jtabs-headers-container';
        header.appendChild(this.headers);

        const controls = document.createElement('div');
        controls.className = 'jtabs-controls';
        controls.setAttribute('draggable', 'false');
        header.appendChild(controls);

        this.el.appendChild(header);
        this.el.appendChild(this.content);

        // Button to create a new worksheet
        if (this.options.allowCreate === true) {
            const add = document.createElement('button');
            add.type = 'button';
            add.className = 'jtabs-add';
            add.setAttribute('aria-label', this.getLabelText('addWorksheet'));
            add.onclick = () => {
                this.create();
            };
            controls.appendChild(add);
        }

        // Scroll buttons for the mouse. The keyboard moves between the tabs with the arrow keys.
        this.prev = document.createElement('div');
        this.prev.className = 'jtabs-prev';
        this.prev.setAttribute('aria-hidden', 'true');
        this.prev.onclick = () => {
            this.updateControls(this.headers.scrollLeft - this.headers.offsetWidth);
        };
        controls.appendChild(this.prev);

        this.next = document.createElement('div');
        this.next.className = 'jtabs-next';
        this.next.setAttribute('aria-hidden', 'true');
        this.next.onclick = () => {
            this.updateControls(this.headers.scrollLeft + this.headers.offsetWidth);
        };
        controls.appendChild(this.next);

        this.border = document.createElement('div');
        this.border.className = 'jtabs-border';
        this.border.setAttribute('aria-hidden', 'true');
        this.headers.appendChild(this.border);

        if (this.options.animation) {
            this.el.classList.add('jtabs-animation');
        }

        for (const tab of this.options.data) {
            const div = this.appendElement(tab.title, null, false);

            if (tab.content) {
                div!.innerHTML = tab.content;
            }
        }

        // Events
        this.headers.addEventListener('click', (e) => {
            const index = this.getIndex(e.target as HTMLElement);

            if (index >= 0) {
                this.open(index);

                if (typeof this.options.onclick === 'function') {
                    this.options.onclick(this.el, this, index, this.getHeader(index), this.content.children[index]);
                }
            }
        });

        this.headers.addEventListener('contextmenu', (e) => {
            const index = this.getIndex(e.target as HTMLElement);

            if (index >= 0) {
                this.open(index);
            }
        });

        // Keyboard: the arrow keys, Home and End move the focus and open the tab
        this.headers.addEventListener('keydown', (e) => {
            const index = this.getIndex(e.target as HTMLElement);
            const total = this.content.children.length;

            if (index < 0 || !total || e.altKey || e.ctrlKey || e.metaKey) {
                return;
            }

            let target = null;

            if (e.key === 'ArrowRight') {
                target = (index + 1) % total;
            } else if (e.key === 'ArrowLeft') {
                target = (index - 1 + total) % total;
            } else if (e.key === 'Home') {
                target = 0;
            } else if (e.key === 'End') {
                target = total - 1;
            }

            if (target !== null) {
                e.preventDefault();
                this.open(target);
                this.getHeader(target)!.focus();
            }
        });

        if (this.content.children.length) {
            this.open(0);
        } else {
            this.updateHeaders();
        }

        this.updateControls(0);
    }
}
