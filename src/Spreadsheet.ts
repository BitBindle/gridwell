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
        const spreadsheet = this;
        // The options that jspreadsheet() got, and the element that the spreadsheet fills
        const options = spreadsheet.config;
        const el = spreadsheet.element;

        // setPlugins ignores a missing value
        spreadsheet.setPlugins(options.plugins as Record<string, jspreadsheet.PluginFactory>);

        // Create as worksheets
        await this.createWorksheets(options, el);

        spreadsheet.element.appendChild(spreadsheet.contextMenu);

        // Create element
        jSuites.contextmenu(spreadsheet.contextMenu, {
            onclick() {
                spreadsheet.contextMenu.contextmenu.close(false);
            },
        });

        // Fullscreen
        if (spreadsheet.config.fullscreen == true) {
            spreadsheet.element.classList.add('fullscreen');
        }

        spreadsheet.toolbarManager.show();

        // Build handlers
        if (options.root) {
            setEvents(options.root);
        } else {
            setEvents(document);
        }

        el.spreadsheet = spreadsheet;
    }

    private async createWorksheets(options: jspreadsheet.SpreadsheetOptions, el: HTMLElement) {
        const spreadsheet = this;

        // Create worksheets
        const o = options.worksheets;
        if (o) {
            const tabsOptions: Partial<TabsOptions> = {
                animation: true,
                onbeforecreate(element: HTMLElement, title?: string | null) {
                    if (title) {
                        return title;
                    } else {
                        return spreadsheet.getNextDefaultWorksheetName();
                    }
                },
                oncreate(element: HTMLElement & { tabs: Tabs }, newTabContent: HTMLDivElement | undefined) {
                    if (!spreadsheet.creationThroughJss) {
                        const worksheetName = element.tabs.headers.children[element.tabs.headers.children.length - 2].innerHTML;

                        spreadsheet.createWorksheetObj({
                            minDimensions: [10, 15],
                            worksheetName,
                        });
                    } else {
                        spreadsheet.creationThroughJss = false;
                    }

                    const newWorksheet = spreadsheet.worksheets[spreadsheet.worksheets.length - 1];

                    // build adds jssWorksheet
                    // onbeforecreate always gives a title, so the tab has its content
                    newWorksheet.element = newTabContent as Worksheet['element'];

                    newWorksheet.build().then(() => {
                        spreadsheet.toolbarManager.update(newWorksheet);

                        newWorksheet.dispatch('oncreateworksheet', newWorksheet, newWorksheet.options, spreadsheet.worksheets.length - 1);
                    });
                },
                onchange(element: HTMLElement, instance: Tabs, tabIndex: number) {
                    if (spreadsheet.worksheets.length != 0 && spreadsheet.worksheets[tabIndex]) {
                        spreadsheet.toolbarManager.update(spreadsheet.worksheets[tabIndex]);
                    }
                },
            };

            if (options.tabs == true) {
                tabsOptions.allowCreate = true;
            }

            // Hide the tab headers with any number of worksheets, for example when the application shows its own tabs
            if (options.hideTabs == true) {
                tabsOptions.hideHeaders = true;
            }

            tabsOptions.data = [];

            let sheetNumber = 1;

            for (let i = 0; i < o.length; i++) {
                if (!o[i].worksheetName) {
                    o[i].worksheetName = 'Sheet' + sheetNumber++;
                }

                tabsOptions.data.push({
                    // Set above
                    title: o[i].worksheetName!,
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
                    Object.entries(o[i].style!).forEach(([cellName, value]) => {
                        if (typeof value === 'number') {
                            // A number is the index of a style in the spreadsheet options, which the application passes with it
                            o[i].style![cellName] = spreadsheetStyles![value];
                        }
                    });
                }

                spreadsheet.worksheets.push(new Worksheet(spreadsheet, o[i], tabs.content.children[i] as HTMLElement));

                await spreadsheet.worksheets[i].build();
            }
        } else {
            throw new Error('JSS: worksheets are not defined');
        }
    }

    /**
     * The name of a new worksheet: Sheet and the number after the largest number of the worksheets called Sheet and a number.
     */
    private getNextDefaultWorksheetName() {
        const spreadsheet = this;
        const defaultWorksheetNameRegex = /^Sheet(\d+)$/;

        let largestWorksheetNumber = 0;

        spreadsheet.worksheets.forEach((worksheet) => {
            const regexResult = defaultWorksheetNameRegex.exec(worksheet.options.worksheetName!);
            if (regexResult) {
                largestWorksheetNumber = Math.max(largestWorksheetNumber, parseInt(regexResult[1]));
            }
        });

        return 'Sheet' + (largestWorksheetNumber + 1);
    }

    /**
     * Create a worksheet and add it and its options to the spreadsheet. The worksheet is built when its tab is created.
     */
    createWorksheetObj(options: jspreadsheet.WorksheetOptions) {
        const spreadsheet = this;

        if (!options.worksheetName) {
            options.worksheetName = spreadsheet.getNextDefaultWorksheetName();
        }

        const newWorksheet = new Worksheet(spreadsheet, options);

        spreadsheet.config.worksheets.push(newWorksheet.options);
        spreadsheet.worksheets.push(newWorksheet);

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
        const spreadsheet = this;
        let ret: unknown = null;

        // Dispatch events
        if (!spreadsheet.ignoreEvents) {
            // Call global event
            if (typeof spreadsheet.config.onevent == 'function') {
                // The arguments match the handler of the event, as the type of dispatch says
                ret = (spreadsheet.config.onevent as (...args: unknown[]) => unknown).call(source, event, ...args);
            }
            // Call specific events
            const config = spreadsheet.config;
            if (typeof config[event] == 'function') {
                // The arguments match the handler of the event, as the type of dispatch says
                ret = (config[event] as (...args: unknown[]) => unknown).apply(source, args);
            }

            if (typeof spreadsheet.plugins === 'object') {
                const pluginKeys = Object.keys(spreadsheet.plugins);

                for (let pluginKeyIndex = 0; pluginKeyIndex < pluginKeys.length; pluginKeyIndex++) {
                    const key = pluginKeys[pluginKeyIndex];
                    const plugin = spreadsheet.plugins[key];

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
        const spreadsheet = this;

        // The element of a spreadsheet with tabs has the tabs object
        // getActive returns false only when no tab is open, and a tab is open while the spreadsheet has worksheets
        return spreadsheet.element.tabs ? (spreadsheet.element.tabs.getActive() as number) : 0;
    }

    /**
     * The worksheet at an index, or the worksheet that the spreadsheet shows.
     */
    getWorksheetInstance(index?: number) {
        const spreadsheet = this;

        const worksheetIndex = typeof index !== 'undefined' ? index : spreadsheet.getWorksheetActive();

        return spreadsheet.worksheets[worksheetIndex];
    }

    /**
     * Activate/Disable fullscreen
     * use programmatically : table.fullscreen(); or table.fullscreen(true); or table.fullscreen(false);
     * @Param {boolean} activate
     */
    fullscreen(activate?: boolean) {
        const spreadsheet = this;

        // If activate not defined, get reverse options.fullscreen
        if (activate == null) {
            activate = !spreadsheet.config.fullscreen;
        }

        // If change
        if (spreadsheet.config.fullscreen != activate) {
            spreadsheet.config.fullscreen = activate;

            // Test LazyLoading conflict
            if (activate == true) {
                spreadsheet.element.classList.add('fullscreen');
            } else {
                spreadsheet.element.classList.remove('fullscreen');
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
        const spreadsheet = this;

        return spreadsheet.config;
    }

    /**
     * Change options of the spreadsheet.
     * @param config - The options to change.
     */
    setConfig(config: jspreadsheet.SpreadsheetOptions | jspreadsheet.WorksheetOptions) {
        const spreadsheet = this;

        const keys = Object.keys(config);

        keys.forEach((key) => {
            // The options are set by name, and SpreadsheetOptions has no index signature
            (spreadsheet.config as unknown as Record<string, unknown>)[key] = (config as Record<string, unknown>)[key];

            if (key === 'toolbar') {
                if ((config as Record<string, unknown>)[key] === true) {
                    spreadsheet.showToolbar();
                } else if ((config as Record<string, unknown>)[key] === false) {
                    spreadsheet.hideToolbar();
                }
            }
        });
    }

    /**
     * Add plugins to the spreadsheet. Each plugin factory gets the spreadsheet, empty options and the spreadsheet options.
     */
    setPlugins(newPlugins: Record<string, jspreadsheet.PluginFactory>) {
        const spreadsheet = this;

        if (!spreadsheet.plugins) {
            spreadsheet.plugins = {};
        }

        if (typeof newPlugins == 'object') {
            Object.entries(newPlugins).forEach(([pluginName, plugin]) => {
                spreadsheet.plugins[pluginName] = plugin.call(libraryBase.jspreadsheet, spreadsheet, {}, spreadsheet.config);
            });
        }
    }
}
