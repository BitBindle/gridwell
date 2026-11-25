import jSuites from 'jsuites';

import libraryBase from './libraryBase.js';
import { setEvents } from '../events/index.js';
import { fullscreen } from './fullscreen.js';
import { getWorksheetActive } from './activeWorksheet.js';
import { hideToolbar, showToolbar, updateToolbar } from './toolbar.js';
import { buildWorksheet, createWorksheetObj, getNextDefaultWorksheetName } from './worksheets.js';
import dispatch from './dispatch.js';
import { createFromTable } from './helpers.js';
import { createTabs } from './tabs.js';
import { getSpreadsheetConfig, setConfig } from './config.js';
import type jspreadsheet from '../jspreadsheet';
import { Spreadsheet } from '../Spreadsheet.js';
import { Worksheet } from '../Worksheet.js';
import type { Tabs, TabsOptions } from './tabs.js';

const factory = function () {};

const createWorksheets = async function (spreadsheet: Spreadsheet, options: jspreadsheet.SpreadsheetOptions, el: HTMLElement) {
    // Create worksheets
    const o = options.worksheets;
    if (o) {
        const tabsOptions: Partial<TabsOptions> = {
            animation: true,
            onbeforecreate: function (element: HTMLElement, title?: string | null) {
                if (title) {
                    return title;
                } else {
                    return getNextDefaultWorksheetName(spreadsheet);
                }
            },
            oncreate: function (element: HTMLElement & { tabs: Tabs }, newTabContent: HTMLDivElement | undefined) {
                if (!spreadsheet.creationThroughJss) {
                    const worksheetName = element.tabs.headers.children[element.tabs.headers.children.length - 2].innerHTML;

                    createWorksheetObj.call(spreadsheet.worksheets[0], {
                        minDimensions: [10, 15],
                        worksheetName: worksheetName,
                    });
                } else {
                    spreadsheet.creationThroughJss = false;
                }

                const newWorksheet = spreadsheet.worksheets[spreadsheet.worksheets.length - 1];

                // buildWorksheet adds jssWorksheet
                // onbeforecreate always gives a title, so the tab has its content
                newWorksheet.element = newTabContent as Worksheet['element'];

                buildWorksheet.call(newWorksheet).then(function () {
                    updateToolbar(newWorksheet);

                    dispatch.call(newWorksheet, 'oncreateworksheet', newWorksheet, newWorksheet.options, spreadsheet.worksheets.length - 1);
                });
            },
            onchange: function (element: HTMLElement, instance: Tabs, tabIndex: number) {
                if (spreadsheet.worksheets.length != 0 && spreadsheet.worksheets[tabIndex]) {
                    updateToolbar(spreadsheet.worksheets[tabIndex]);
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

        const tabs = createTabs(el, tabsOptions);

        const spreadsheetStyles = options.style;
        delete options.style;

        for (let i = 0; i < o.length; i++) {
            if (o[i].style) {
                Object.entries(o[i].style!).forEach(function ([cellName, value]) {
                    if (typeof value === 'number') {
                        // A number is the index of a style in the spreadsheet options, which the application passes with it
                        o[i].style![cellName] = spreadsheetStyles![value];
                    }
                });
            }

            spreadsheet.worksheets.push(new Worksheet(spreadsheet, o[i], tabs.content.children[i] as HTMLElement));

            await buildWorksheet.call(spreadsheet.worksheets[i]);
        }
    } else {
        throw new Error('JSS: worksheets are not defined');
    }
};

// A table is replaced by a div. The div gets the spreadsheet in its spreadsheet property.
factory.spreadsheet = async function (el: HTMLElement & { spreadsheet?: Spreadsheet }, options: jspreadsheet.SpreadsheetOptions, worksheets: Worksheet[]) {
    if (el.tagName == 'TABLE') {
        if (!options) {
            // The worksheets are added below
            options = {} as jspreadsheet.SpreadsheetOptions;
        }

        if (!options.worksheets) {
            options.worksheets = [];
        }

        const tableOptions = createFromTable(el, options.worksheets[0]);

        // el is a table, so the options were created
        options.worksheets[0] = tableOptions!;

        const div = document.createElement('div');
        // A table outside the document has no parent, and this throws (issue #55)
        el.parentNode!.insertBefore(div, el);
        el.remove();
        el = div;
    }

    // el gets the spreadsheet at the end
    const spreadsheet = new Spreadsheet(el, options, worksheets);

    // setPlugins ignores a missing value
    spreadsheet.setPlugins(options.plugins as Record<string, jspreadsheet.PluginFactory>);

    // Create as worksheets
    await createWorksheets(spreadsheet, options, el);

    spreadsheet.element.appendChild(spreadsheet.contextMenu);

    // Create element
    jSuites.contextmenu(spreadsheet.contextMenu, {
        onclick: function () {
            spreadsheet.contextMenu.contextmenu.close(false);
        },
    });

    // Fullscreen
    if (spreadsheet.config.fullscreen == true) {
        spreadsheet.element.classList.add('fullscreen');
    }

    showToolbar.call(spreadsheet);

    // Build handlers
    if (options.root) {
        setEvents(options.root);
    } else {
        setEvents(document);
    }

    el.spreadsheet = spreadsheet;

    return spreadsheet;
};

/**
 * Add plugins to the spreadsheet. Each plugin factory gets the spreadsheet, empty options and the spreadsheet options.
 */
const setPlugins = function (this: Spreadsheet, newPlugins: Record<string, jspreadsheet.PluginFactory>) {
    const spreadsheet = this;

    if (!spreadsheet.plugins) {
        spreadsheet.plugins = {};
    }

    if (typeof newPlugins == 'object') {
        Object.entries(newPlugins).forEach(function ([pluginName, plugin]) {
            spreadsheet.plugins[pluginName] = plugin.call(libraryBase.jspreadsheet, spreadsheet, {}, spreadsheet.config);
        });
    }
};

// The public methods are on the prototype, as the methods of a class: not enumerable, and an instance can replace one
const spreadsheetPublicMethods: [keyof jspreadsheet.SpreadsheetInstance, (this: Spreadsheet, ...args: never[]) => unknown][] = [
    ['getWorksheetActive', getWorksheetActive],
    ['fullscreen', fullscreen],
    ['showToolbar', showToolbar],
    ['hideToolbar', hideToolbar],
    ['getConfig', getSpreadsheetConfig],
    ['setConfig', setConfig],
    ['setPlugins', setPlugins],
];

for (const [methodName, method] of spreadsheetPublicMethods) {
    Object.defineProperty(Spreadsheet.prototype, methodName, { value: method, writable: true, configurable: true });
}

export default factory;
