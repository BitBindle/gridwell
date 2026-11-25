import jSuites from 'jsuites';

/**
 * Worksheet tabs. CE uses only part of jSuites.tabs, and needs the keyboard and ARIA behaviour of the
 * WAI-ARIA tabs pattern for the tab headers: https://www.w3.org/WAI/ARIA/apg/patterns/tabs/
 * The object has the methods of jSuites.tabs, and the same class names, so the jSuites styles still apply.
 */
/** A tab header. It keeps a reference to its panel. */
type TabHeader = HTMLButtonElement & { content?: HTMLElement };

/** A handler of a click on a tab, or of a change of the open tab. */
type TabHandler = (el: HTMLElement, tabs: Tabs, index: number, header: TabHeader | null, content: Element) => void;

/** The options of the tabs, which factory.ts gives. A missing option takes its default in createTabs. */
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
}

export interface Tabs {
    options: TabsOptions;
    headers: HTMLDivElement;
    content: HTMLDivElement;
    setBorder(index: number | false): void;
    open(index: number): void;
    selectIndex(a: Element): number;
    /** The index of the open tab, or false when no tab is open. */
    getActive(): number | false;
    rename(i: number, title?: string | null): void;
    create(title?: string): HTMLDivElement | false | undefined;
    remove(index: number): void | false;
    nextNumber(): number;
    deleteElement(index: number): void | false;
    appendElement(
        title?: string | null,
        cb?: ((div: HTMLDivElement, header: TabHeader) => void) | null,
        openTab?: boolean,
        position?: number
    ): HTMLDivElement | undefined;
    updateContent(position: number, newContent: string | HTMLElement): void;
    updatePosition(f: number, t: number, ignoreEvents?: boolean, openTab?: boolean): void;
    move(f: number, t: number, ignoreEvents?: boolean, openTab?: boolean): void;
    init(): void;
}

export const createTabs = function (el: HTMLElement & { tabs?: Tabs }, options: Partial<TabsOptions>) {
    const obj = {} as Tabs;
    // Each option is set below
    obj.options = {} as TabsOptions;

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
    };

    for (const property in defaults) {
        if (options && Object.prototype.hasOwnProperty.call(options, property)) {
            (obj.options as unknown as Record<string, unknown>)[property] = (options as Record<string, unknown>)[property];
        } else {
            (obj.options as unknown as Record<string, unknown>)[property] = (defaults as unknown as Record<string, unknown>)[property];
        }
    }

    // init creates these elements
    let prev: HTMLDivElement = null!;
    let next: HTMLDivElement = null!;
    let border: HTMLDivElement = null!;

    const getHeader = function (index: number | false) {
        const header = obj.headers.children[index as number];

        if (header && header !== border) {
            return header as TabHeader;
        }

        return null;
    };

    const getIndex = function (element: HTMLElement) {
        while (element && element.parentNode !== obj.headers) {
            element = element.parentNode as HTMLElement;
        }

        if (!element || element === border) {
            return -1;
        }

        return Array.prototype.indexOf.call(obj.headers.children, element);
    };

    const updateControls = function (x: number) {
        if (typeof obj.headers.scrollTo == 'function') {
            obj.headers.scrollTo({
                left: x,
                behavior: 'smooth',
            });
        } else {
            obj.headers.scrollLeft = x;
        }

        if (x <= 1) {
            prev.classList.add('disabled');
        } else {
            prev.classList.remove('disabled');
        }

        if (x >= obj.headers.scrollWidth - obj.headers.offsetWidth) {
            next.classList.add('disabled');
        } else {
            next.classList.remove('disabled');
        }

        if (obj.headers.scrollWidth <= obj.headers.offsetWidth) {
            prev.style.display = 'none';
            next.style.display = 'none';
        } else {
            prev.style.display = '';
            next.style.display = '';
        }
    };

    /**
     * Show the headers when the user can choose a worksheet or create one, unless the option hideHeaders is true
     */
    const updateHeaders = function () {
        const container = obj.headers.parentNode as HTMLElement;

        if (obj.options.hideHeaders == true || (obj.content.children.length < 2 && obj.options.allowCreate != true)) {
            container.style.display = 'none';
        } else {
            container.style.display = '';

            const index = obj.getActive();
            const header = getHeader(index);

            if (header) {
                const x1 = header.offsetLeft;
                const x2 = x1 + header.offsetWidth;
                const r1 = obj.headers.scrollLeft;
                const r2 = r1 + obj.headers.offsetWidth;

                if (!(r1 <= x1 && r2 >= x2)) {
                    // Out of the viewport
                    updateControls(x1 - 1);
                }

                obj.setBorder(index);
            }
        }
    };

    obj.setBorder = function (index) {
        if (obj.options.animation) {
            setTimeout(function () {
                const header = getHeader(index);

                if (header) {
                    border.style.width = header.getBoundingClientRect().width + 'px';
                    border.style.left = header.offsetLeft + 'px';
                    border.style.bottom = '0px';
                }
            }, 50);
        }
    };

    obj.open = function (index) {
        if (!obj.content.children[index] || !getHeader(index)) {
            return;
        }

        index = parseInt(index);

        const previous = obj.getActive();

        for (let i = 0; i < obj.content.children.length; i++) {
            const header = getHeader(i);
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
                obj.content.children[i].classList.add('jtabs-selected');
            } else {
                obj.content.children[i].classList.remove('jtabs-selected');
            }
        }

        if (previous !== index && typeof obj.options.onchange == 'function') {
            obj.options.onchange(el, obj, index, getHeader(index), obj.content.children[index]);
        }

        updateHeaders();
    };

    obj.selectIndex = function (a) {
        const index = Array.prototype.indexOf.call(obj.headers.children, a);

        if (index >= 0) {
            obj.open(index);
        }

        return index;
    };

    obj.getActive = function () {
        for (let i = 0; i < obj.headers.children.length; i++) {
            if (obj.headers.children[i].classList.contains('jtabs-selected')) {
                return i;
            }
        }

        return false;
    };

    obj.rename = function (i, title) {
        const header = getHeader(i);

        if (!header) {
            return;
        }

        if (!title) {
            title = prompt('New title', header.textContent);
        }

        if (title) {
            header.textContent = title;
        }

        obj.setBorder(obj.getActive());
    };

    obj.create = function (title) {
        if (typeof obj.options.onbeforecreate == 'function') {
            const ret = obj.options.onbeforecreate(el, title);

            if (ret === false) {
                return false;
            } else {
                title = ret;
            }
        }

        const div = obj.appendElement(title);

        if (typeof obj.options.oncreate == 'function') {
            // createTabs sets el.tabs before a tab can be created
            obj.options.oncreate(el as HTMLElement & { tabs: Tabs }, div);
        }

        obj.setBorder(obj.getActive());

        return div;
    };

    obj.remove = function (index) {
        return obj.deleteElement(index);
    };

    obj.nextNumber = function () {
        let num = 0;

        for (let i = 0; i < obj.content.children.length; i++) {
            // The match is compared and parsed as text
            const tmp = getHeader(i)!.textContent.match(/[0-9].*/) as unknown as string | null;

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
    };

    obj.deleteElement = function (index) {
        const current = obj.getActive();
        const header = getHeader(index);

        if (!header) {
            return false;
        }

        const hadFocus = header === document.activeElement;

        obj.headers.removeChild(header);
        obj.content.removeChild(obj.content.children[index]);

        if (current === index) {
            obj.open(0);
        } else {
            updateHeaders();
        }

        // Keep the focus in the tab list
        if (hadFocus && getHeader(obj.getActive())) {
            getHeader(obj.getActive())!.focus();
        }

        if (typeof obj.options.ondelete == 'function') {
            obj.options.ondelete(el, index);
        }
    };

    obj.appendElement = function (title, cb, openTab, position) {
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

            if (typeof position === 'undefined' || !obj.content.children[position]) {
                obj.content.appendChild(div);
                obj.headers.insertBefore(h, border);
            } else {
                obj.content.insertBefore(div, obj.content.children[position]);
                obj.headers.insertBefore(h, getHeader(position));
            }

            if (openTab !== false) {
                obj.selectIndex(h);
            } else {
                updateHeaders();
            }

            if (typeof cb == 'function') {
                cb(div, h);
            }

            return div;
        }
    };

    obj.updateContent = function (position, newContent) {
        const current = obj.content.children[position];

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

        obj.content.replaceChild(contentItem, current);

        if (getHeader(position)) {
            getHeader(position)!.content = contentItem;
            getHeader(position)!.setAttribute('aria-controls', contentItem.id);
        }

        obj.setBorder(obj.getActive());
    };

    obj.updatePosition = function (f, t, ignoreEvents, openTab) {
        // Update the position of the content
        if (f > t) {
            obj.content.insertBefore(obj.content.children[f], obj.content.children[t]);
        } else {
            obj.content.insertBefore(obj.content.children[f], obj.content.children[t].nextSibling);
        }

        if (openTab !== false) {
            obj.open(t);
        } else {
            obj.setBorder(obj.getActive());
        }

        if (!ignoreEvents && typeof obj.options.onchangeposition == 'function') {
            obj.options.onchangeposition(obj.headers, f, t);
        }
    };

    obj.move = function (f, t, ignoreEvents, openTab) {
        if (f > t) {
            obj.headers.insertBefore(getHeader(f)!, getHeader(t));
        } else {
            obj.headers.insertBefore(getHeader(f)!, getHeader(t)!.nextSibling);
        }

        obj.updatePosition(f, t, ignoreEvents, openTab);
    };

    obj.init = function () {
        el.classList.add('jtabs');
        el.textContent = '';

        obj.headers = document.createElement('div');
        obj.headers.className = 'jtabs-headers';
        obj.headers.setAttribute('role', 'tablist');
        obj.headers.setAttribute('aria-label', jSuites.translate('Worksheets'));

        obj.content = document.createElement('div');
        obj.content.className = 'jtabs-content';
        obj.content.setAttribute('role', 'region');
        obj.content.setAttribute('aria-label', 'Tab Panels');

        const header = document.createElement('div');
        header.className = 'jtabs-headers-container';
        header.appendChild(obj.headers);

        const controls = document.createElement('div');
        controls.className = 'jtabs-controls';
        controls.setAttribute('draggable', 'false');
        header.appendChild(controls);

        el.appendChild(header);
        el.appendChild(obj.content);

        // Button to create a new worksheet
        if (obj.options.allowCreate == true) {
            const add = document.createElement('button');
            add.type = 'button';
            add.className = 'jtabs-add';
            add.setAttribute('aria-label', jSuites.translate('Add a new worksheet'));
            add.onclick = function () {
                obj.create();
            };
            controls.appendChild(add);
        }

        // Scroll buttons for the mouse. The keyboard moves between the tabs with the arrow keys.
        prev = document.createElement('div');
        prev.className = 'jtabs-prev';
        prev.setAttribute('aria-hidden', 'true');
        prev.onclick = function () {
            updateControls(obj.headers.scrollLeft - obj.headers.offsetWidth);
        };
        controls.appendChild(prev);

        next = document.createElement('div');
        next.className = 'jtabs-next';
        next.setAttribute('aria-hidden', 'true');
        next.onclick = function () {
            updateControls(obj.headers.scrollLeft + obj.headers.offsetWidth);
        };
        controls.appendChild(next);

        border = document.createElement('div');
        border.className = 'jtabs-border';
        border.setAttribute('aria-hidden', 'true');
        obj.headers.appendChild(border);

        if (obj.options.animation) {
            el.classList.add('jtabs-animation');
        }

        for (let i = 0; i < obj.options.data.length; i++) {
            const div = obj.appendElement(obj.options.data[i].title, null, false);

            if (obj.options.data[i].content) {
                div!.innerHTML = obj.options.data[i].content;
            }
        }

        // Events
        obj.headers.addEventListener('click', function (e) {
            const index = getIndex(e.target as HTMLElement);

            if (index >= 0) {
                obj.open(index);

                if (typeof obj.options.onclick == 'function') {
                    obj.options.onclick(el, obj, index, getHeader(index), obj.content.children[index]);
                }
            }
        });

        obj.headers.addEventListener('contextmenu', function (e) {
            const index = getIndex(e.target as HTMLElement);

            if (index >= 0) {
                obj.open(index);
            }
        });

        // Keyboard: the arrow keys, Home and End move the focus and open the tab
        obj.headers.addEventListener('keydown', function (e) {
            const index = getIndex(e.target as HTMLElement);
            const total = obj.content.children.length;

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
                obj.open(target);
                getHeader(target)!.focus();
            }
        });

        if (obj.content.children.length) {
            obj.open(0);
        } else {
            updateHeaders();
        }

        updateControls(0);
    };

    obj.init();

    el.tabs = obj;

    return obj;
};
