import libraryBase from './utils/libraryBase.js';
import { setEvents } from './events/index.js';
import { Toolbar } from './utils/toolbar.js';
import { Tabs } from './utils/tabs.js';
import { ContextMenu } from './utils/contextMenu.js';
import { Worksheet } from './Worksheet.js';
import type jspreadsheet from './jspreadsheet';
import type { EventHandlers, EventName, SpreadsheetState } from './types/internal';
import type { TabsOptions } from './utils/tabs.js';

/** The events that occur also while the events are ignored. dispatch tells why. */
const alwaysDispatched: readonly EventName[] = ['onbeforerequest', 'onbeforesave', 'onsave', 'oneditordetach', 'onrowsattached'];

/**
 * The fields of a spreadsheet and its public methods. The fields are in SpreadsheetState (src/types/internal.d.ts).
 */
// eslint-disable-next-line @typescript-eslint/no-empty-object-type, @typescript-eslint/no-unsafe-declaration-merging
export interface Spreadsheet extends SpreadsheetState {}

/**
 * A spreadsheet: the worksheets in their tabs, the toolbar, the context menu and the plugins.
 *
 * The constructor sets up the state that the spreadsheet has before it is built. build then creates the worksheets, the
 * toolbar and the context menu.
 */
// eslint-disable-next-line @typescript-eslint/no-unsafe-declaration-merging
export class Spreadsheet {
    constructor(element: HTMLElement, config: jspreadsheet.SpreadsheetOptions, worksheets: Worksheet[]) {
        // jspreadsheet() returns this array before the worksheets are built, and build fills it
        this.worksheets = worksheets;
        this.config = config;
        // The element gets the spreadsheet in its spreadsheet property when the spreadsheet is built
        this.element = element as Spreadsheet['element'];
        this.el = this.element;

        // The context menu, which keeps its ContextMenu as its contextmenu property
        this.contextMenu = document.createElement('div') as Spreadsheet['contextMenu'];
        this.contextMenu.className = 'jss_contextmenu';
        new ContextMenu(this.contextMenu);

        // The polite live region of announce
        this.liveRegion = document.createElement('div');
        this.liveRegion.className = 'jss_live';
        this.liveRegion.setAttribute('role', 'status');
        this.liveRegion.setAttribute('aria-live', 'polite');

        this.toolbarManager = new Toolbar(this);
    }

    /**
     * Build the spreadsheet: the plugins, the worksheets in their tabs, the context menu, the toolbar and the document
     * event handlers. The element gets the spreadsheet at the end.
     */
    async build() {
        // The options that jspreadsheet() got, and the element that the spreadsheet fills
        const options = this.config;
        const el = this.element;

        // setPlugins ignores a missing value
        this.setPlugins(options.plugins as Record<string, jspreadsheet.PluginFactory>);

        // Create as worksheets
        await this.createWorksheets(options, el);

        this.element.appendChild(this.contextMenu);
        this.element.appendChild(this.liveRegion);

        // Fullscreen
        if (this.config.fullscreen === true) {
            this.element.classList.add('fullscreen');
        }

        this.toolbarManager.show();

        // Build handlers
        if (options.root) {
            setEvents(options.root);
        } else {
            setEvents(document);
        }

        el.spreadsheet = this;
    }

    private async createWorksheets(options: jspreadsheet.SpreadsheetOptions, el: HTMLElement) {
        // Create worksheets
        const o = options.worksheets;
        if (o) {
            const tabsOptions: Partial<TabsOptions> = {
                animation: true,
                labels: options.labels ?? {},
                onbeforecreate: (element: HTMLElement, title?: string | null) => {
                    if (title) {
                        return title;
                    } else {
                        return this.getNextDefaultWorksheetName();
                    }
                },
                oncreate: (element: HTMLElement & { tabs: Tabs }, newTabContent: HTMLDivElement | undefined) => {
                    if (!this.creationThroughJss) {
                        const worksheetName = element.tabs.headers.children[element.tabs.headers.children.length - 2].innerHTML;

                        this.createWorksheetObj({
                            minDimensions: [10, 15],
                            worksheetName,
                        });
                    } else {
                        this.creationThroughJss = false;
                    }

                    const newWorksheet = this.worksheets[this.worksheets.length - 1];

                    // build adds jssWorksheet
                    // onbeforecreate always gives a title, so the tab has its content
                    newWorksheet.element = newTabContent as Worksheet['element'];

                    newWorksheet.build().then(() => {
                        this.toolbarManager.update(newWorksheet);

                        // createWorksheet can move the new worksheet before it is built
                        newWorksheet.dispatch('oncreateworksheet', newWorksheet, newWorksheet.options, this.worksheets.indexOf(newWorksheet));
                    });
                },
                onchange: (element: HTMLElement, instance: Tabs, tabIndex: number) => {
                    if (this.worksheets.length !== 0 && this.worksheets[tabIndex]) {
                        this.toolbarManager.update(this.worksheets[tabIndex]);
                    }
                },
            };

            if (options.tabs === true) {
                tabsOptions.allowCreate = true;
            }

            // Hide the tab headers with any number of worksheets, for example when the application shows its own tabs
            if (options.hideTabs === true) {
                tabsOptions.hideHeaders = true;
            }

            tabsOptions.data = [];

            let sheetNumber = 1;

            for (const worksheetOptions of o) {
                if (!worksheetOptions.worksheetName) {
                    worksheetOptions.worksheetName = `Sheet${sheetNumber++}`;
                }

                tabsOptions.data.push({
                    // Set above
                    title: worksheetOptions.worksheetName!,
                    content: '',
                });
            }

            el.classList.add('jss_spreadsheet');
            el.tabIndex = 0;

            const tabs = new Tabs(el, tabsOptions);

            const spreadsheetStyles = options.style;
            delete options.style;

            for (let i = 0; i < o.length; i++) {
                if (o[i].style) {
                    for (const [cellName, value] of Object.entries(o[i].style!)) {
                        if (typeof value === 'number') {
                            // A number is the index of a style in the spreadsheet options, which the application passes with it
                            o[i].style![cellName] = spreadsheetStyles![value];
                        }
                    }
                }

                this.worksheets.push(new Worksheet(this, o[i], tabs.content.children[i] as HTMLElement));

                await this.worksheets[i].build();
            }
        } else {
            throw new Error('JSS: worksheets are not defined');
        }
    }

    /**
     * The name of a new worksheet: Sheet and the number after the largest number of the worksheets called Sheet and a number.
     */
    private getNextDefaultWorksheetName() {
        const defaultWorksheetNameRegex = /^Sheet(\d+)$/;

        let largestWorksheetNumber = 0;

        this.worksheets.forEach((worksheet) => {
            const regexResult = defaultWorksheetNameRegex.exec(worksheet.options.worksheetName!);
            if (regexResult) {
                largestWorksheetNumber = Math.max(largestWorksheetNumber, parseInt(regexResult[1]));
            }
        });

        return `Sheet${largestWorksheetNumber + 1}`;
    }

    /**
     * Create a worksheet and add it and its options to the spreadsheet. The worksheet is built when its tab is created.
     */
    createWorksheetObj(options: jspreadsheet.WorksheetOptions) {
        if (!options.worksheetName) {
            options.worksheetName = this.getNextDefaultWorksheetName();
        }

        const newWorksheet = new Worksheet(this, options);

        this.config.worksheets.push(newWorksheet.options);
        this.worksheets.push(newWorksheet);

        return newWorksheet;
    }

    /**
     * Build the worksheet at a position again from new options, in its tab. The jspreadsheet.d.ts typing of loadWorksheet
     * tells what stays and what changes.
     * @returns The worksheet when it is built, or undefined for a position without a worksheet.
     */
    async loadWorksheet(position: number, options: jspreadsheet.WorksheetOptions) {
        const worksheet = this.worksheets[position];
        if (!worksheet) {
            return undefined;
        }

        if (worksheet.editor?.isOpen()) {
            worksheet.editor.close(worksheet.editor.cell, false);
        }
        // An open filter and the context menu belong to the old table, so they close without their events
        for (const dropdown of worksheet.element.querySelectorAll<HTMLElement & { dropdown?: { close: (ignoreEvents?: boolean) => void } }>('.jdropdown')) {
            dropdown.dropdown?.close(true);
        }
        this.contextMenu.contextmenu?.close();

        const oldName = worksheet.options.worksheetName!;
        const name = options.worksheetName;
        const renamed =
            typeof name === 'string' &&
            !!name.trim() &&
            name !== oldName &&
            !this.worksheets.some((other) => other !== worksheet && other.options.worksheetName === name);
        if (renamed) {
            this.element.tabs!.rename(position, name);
            worksheet.options.worksheetName = name;
            worksheet.dispatch('onrenameworksheet', worksheet, position, name, oldName);
        } else {
            options.worksheetName = oldName;
        }
        // The options of the first build give the styles of the cells with a getter, which the reset would empty
        if (options === worksheet.options && worksheet.records) {
            Object.defineProperty(options, 'style', { value: worksheet.getStyle(), writable: true, enumerable: true, configurable: true });
        }

        // The size observer of the frozen rows holds the old rows. The timer of a search, a touch or a rename of a header
        // belongs to the active worksheet, and acts on the old table. The timer of a lazy load checks its worksheet.
        worksheet.rowManager?.stopFreeze();
        worksheet.rowSource?.stop();
        if (libraryBase.jspreadsheet.current === worksheet && libraryBase.jspreadsheet.timeControl) {
            clearTimeout(libraryBase.jspreadsheet.timeControl);
            libraryBase.jspreadsheet.timeControl = null;
        }

        // A load that replaces a load that has not ended keeps the active worksheet and the focus from before the first one
        const element = worksheet.element;
        const pending = worksheet.pendingLoad;
        const focused = pending ? pending.focused : element.contains(document.activeElement);
        const current = pending ? pending.current : libraryBase.jspreadsheet.current;
        const ignore = worksheet.history?.ignore;
        const index = this.config.worksheets.indexOf(worksheet.options);
        if (index >= 0) {
            this.config.worksheets[index] = options;
        }

        // The state of the old table goes, and the methods stay on the class. The number of the load stays, so that a load
        // that has not ended can see that a later load replaced it.
        const generation = (worksheet.generation ?? 0) + 1;
        for (const key of Object.keys(worksheet) as (keyof Worksheet)[]) {
            if (key !== 'parent' && key !== 'element') {
                delete worksheet[key];
            }
        }
        element.innerHTML = '';
        element.classList.remove('with-toolbar');
        worksheet.generation = generation;
        worksheet.pendingLoad = { current, focused };
        worksheet.setup(options);
        worksheet.history.ignore = ignore === true;

        // build makes the worksheet the active worksheet. While its data or the sources of its dropdown columns load, the
        // worksheet has no table, so another worksheet stays active, or no worksheet is active. The user can make another
        // worksheet active during the load, and that worksheet stays active.
        const building = worksheet.build(true);
        if (libraryBase.jspreadsheet.current === worksheet) {
            libraryBase.jspreadsheet.current = current === worksheet ? null : current;
        }
        try {
            await building;
        } finally {
            // The last load gives back the active worksheet, also when the build fails, but not after destroy
            if (worksheet.generation === generation) {
                delete worksheet.pendingLoad;
                const active = libraryBase.jspreadsheet.current;
                // destroy empties the element of the spreadsheet, so the worksheet is not in it any more
                const destroyed = !this.element.contains(element);
                if (!destroyed && (active === worksheet || (active == null && current === worksheet))) {
                    libraryBase.jspreadsheet.current = current;
                }
            }
        }
        if (worksheet.generation !== generation) {
            return worksheet;
        }

        // The focus moves only when no other element took it during the load
        const lost = !document.activeElement || document.activeElement === document.body;
        if (focused && lost) {
            // The tab stop becomes the selection, so that the keys move it
            const td = element.querySelector<HTMLTableCellElement>('td[tabindex="0"]');
            if (td) {
                const x = Number(td.getAttribute('data-x'));
                const y = Number(td.getAttribute('data-y'));
                worksheet.updateSelectionFromCoords(x, y, x, y);
                td.focus({ preventScroll: true });
            }
        }
        if (this.worksheets[this.getWorksheetActive()] === worksheet) {
            this.toolbarManager.update(worksheet);
        }

        return worksheet;
    }

    /**
     * Move a worksheet, its options and its tab to another position. The open tab stays open, and an element of the
     * moved worksheet that has the focus keeps it.
     */
    placeWorksheet(from: number, to: number) {
        const worksheet = this.worksheets[from];
        const tabs = this.element.tabs!;
        const header = tabs.headers.children[from];
        const active = document.activeElement as HTMLElement | null;
        const focused = active && (header.contains(active) || worksheet.element.contains(active)) ? active : null;

        this.worksheets.splice(to, 0, this.worksheets.splice(from, 1)[0]);
        // The options are found by identity, so a list that does not match the worksheets keeps its other options
        const index = this.config.worksheets.indexOf(worksheet.options);
        if (index >= 0) {
            this.config.worksheets.splice(index, 1);
            this.config.worksheets.splice(Math.min(to, this.config.worksheets.length), 0, worksheet.options);
        }
        tabs.move(from, to, true, false);

        // The browser takes the focus from an element that the DOM moves
        if (focused && document.activeElement !== focused) {
            focused.focus({ preventScroll: true });
        }
    }

    /**
     * Trigger an event of the spreadsheet. The handlers get the spreadsheet as `this`.
     *
     * The first argument is the name of the event, for example "onload". The other arguments go to the event handlers.
     * It returns the result of the last handler, or null when no handler ran.
     */
    dispatch<E extends EventName>(event: E, ...args: Parameters<EventHandlers[E]>): ReturnType<EventHandlers[E]> | null {
        // The result of a handler of the event, as the type of dispatch says
        return this.callHandlers(this, event, args) as ReturnType<EventHandlers[E]> | null;
    }

    /**
     * Call the handlers of an event: the onevent option of the spreadsheet, then the handler of the event in the options,
     * then the onevent of each plugin. They get source, the worksheet or the spreadsheet that triggers the event, as
     * `this`. It returns false when a handler returns false, so that any handler can cancel an action. Otherwise it
     * returns the last result that is not undefined, or null.
     */
    callHandlers(source: Worksheet | Spreadsheet, event: EventName, args: unknown[]) {
        let ret: unknown = null;

        // A false stays, and a handler that returns nothing keeps the result of the handlers before it
        function keep(result: unknown) {
            if (ret !== false && result !== undefined) {
                ret = result;
            }
        }

        // Dispatch events. The persistence still saves while the events are ignored, for example the changes of an undo,
        // so its events, and the event of its request, occur too. The application can take the edit of an editor whose cell
        // leaves the table, also during an undo, so oneditordetach occurs too. The rows that lazy loading puts in the table,
        // for example during the undo of a sort, need the work of the application, so onrowsattached occurs too.
        if (!this.ignoreEvents || alwaysDispatched.includes(event)) {
            // Call global event
            if (typeof this.config.onevent === 'function') {
                // The arguments match the handler of the event, as the type of dispatch says
                keep((this.config.onevent as (...args: unknown[]) => unknown).call(source, event, ...args));
            }
            // Call specific events
            const config = this.config;
            if (typeof config[event] === 'function') {
                // The arguments match the handler of the event, as the type of dispatch says
                keep((config[event] as (...args: unknown[]) => unknown).apply(source, args));
            }

            if (typeof this.plugins === 'object') {
                for (const key of Object.keys(this.plugins)) {
                    const plugin = this.plugins[key];

                    if (typeof plugin.onevent === 'function') {
                        // The arguments match the handler of the event, as the type of dispatch says
                        keep((plugin.onevent as (...args: unknown[]) => unknown).call(source, event, ...args));
                    }
                }
            }
        }

        return ret;
    }

    /**
     * The index of the worksheet that the spreadsheet shows.
     */
    getWorksheetActive(): number {
        // The element of a spreadsheet with tabs has the tabs object
        // getActive returns false only when no tab is open, and a tab is open while the spreadsheet has worksheets
        return this.element.tabs ? (this.element.tabs.getActive() as number) : 0;
    }

    /**
     * The worksheet at an index, or the worksheet that the spreadsheet shows.
     */
    getWorksheetInstance(index?: number) {
        const worksheetIndex = typeof index !== 'undefined' ? index : this.getWorksheetActive();

        return this.worksheets[worksheetIndex];
    }

    /**
     * Show the spreadsheet in full screen, or show it in the page again.
     * @param activate - Whether to show it in full screen. Without it, the spreadsheet switches.
     */
    fullscreen(activate?: boolean) {
        // If activate not defined, get reverse options.fullscreen
        if (activate == null) {
            activate = !this.config.fullscreen;
        }

        // If change
        if (this.config.fullscreen !== activate) {
            this.config.fullscreen = activate;

            // Test LazyLoading conflict
            if (activate === true) {
                this.element.classList.add('fullscreen');
            } else {
                this.element.classList.remove('fullscreen');
            }
        }
    }

    /**
     * Read a text to the user of assistive technology through the polite live region of the spreadsheet.
     * @param text - Text to read.
     */
    announce(text: string) {
        // A new element, so that a screen reader also reads the same text again
        const item = document.createElement('div');
        item.textContent = text;
        this.liveRegion.replaceChildren(item);
    }

    showToolbar() {
        return this.toolbarManager.show();
    }

    hideToolbar() {
        return this.toolbarManager.hide();
    }

    /**
     * Get the options of the spreadsheet.
     */
    getConfig() {
        return this.config;
    }

    /**
     * Change options of the spreadsheet.
     * @param config - The options to change.
     */
    setConfig(config: Partial<jspreadsheet.SpreadsheetOptions> | Partial<jspreadsheet.WorksheetOptions>) {
        for (const [key, value] of Object.entries(config)) {
            // The options are set by name, and SpreadsheetOptions has no index signature
            (this.config as unknown as Record<string, unknown>)[key] = value;

            if (key === 'toolbar') {
                if (value === true) {
                    this.showToolbar();
                } else if (value === false) {
                    this.hideToolbar();
                }
            }
        }
    }

    /**
     * Add plugins to the spreadsheet. Each plugin factory gets the spreadsheet, empty options and the spreadsheet options.
     */
    setPlugins(newPlugins: Record<string, jspreadsheet.PluginFactory>) {
        if (!this.plugins) {
            this.plugins = {};
        }

        if (typeof newPlugins === 'object') {
            for (const [pluginName, plugin] of Object.entries(newPlugins)) {
                this.plugins[pluginName] = plugin.call(libraryBase.jspreadsheet, this, {}, this.config);
            }
        }
    }
}
