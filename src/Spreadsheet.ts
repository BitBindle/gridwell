import jSuites from 'jsuites';
import libraryBase from './utils/libraryBase.js';
import { setEvents } from './events/index.js';
import { Toolbar } from './utils/toolbar.js';
import { Tabs } from './utils/tabs.js';
import { Worksheet } from './Worksheet.js';
import type jspreadsheet from './jspreadsheet';
import type { EventHandlers, EventName, SpreadsheetState } from './types/internal';
import type { TabsOptions } from './utils/tabs.js';

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

        // The context menu container. jSuites.contextmenu adds its instance when the spreadsheet is built.
        this.contextMenu = document.createElement('div') as Spreadsheet['contextMenu'];
        this.contextMenu.className = 'jss_contextmenu';

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

        // Create element
        jSuites.contextmenu(this.contextMenu, {
            onclick: () => {
                this.contextMenu.contextmenu.close(false);
            },
        });

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

                        newWorksheet.dispatch('oncreateworksheet', newWorksheet, newWorksheet.options, this.worksheets.length - 1);
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
     * `this`. It returns the result of the last handler.
     */
    callHandlers(source: Worksheet | Spreadsheet, event: EventName, args: unknown[]) {
        let ret: unknown = null;

        // Dispatch events
        if (!this.ignoreEvents) {
            // Call global event
            if (typeof this.config.onevent === 'function') {
                // The arguments match the handler of the event, as the type of dispatch says
                ret = (this.config.onevent as (...args: unknown[]) => unknown).call(source, event, ...args);
            }
            // Call specific events
            const config = this.config;
            if (typeof config[event] === 'function') {
                // The arguments match the handler of the event, as the type of dispatch says
                ret = (config[event] as (...args: unknown[]) => unknown).apply(source, args);
            }

            if (typeof this.plugins === 'object') {
                for (const key of Object.keys(this.plugins)) {
                    const plugin = this.plugins[key];

                    if (typeof plugin.onevent === 'function') {
                        // The arguments match the handler of the event, as the type of dispatch says
                        ret = (plugin.onevent as (...args: unknown[]) => unknown).call(source, event, ...args);
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
    setConfig(config: jspreadsheet.SpreadsheetOptions | jspreadsheet.WorksheetOptions) {
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
