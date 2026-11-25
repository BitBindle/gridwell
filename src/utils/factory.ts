import jSuites from 'jsuites';

import libraryBase from './libraryBase.js';
import { setEvents } from './events.js';
import { fullscreen, getWorksheetActive } from './internal.js';
import { hideToolbar, showToolbar, updateToolbar } from './toolbar.js';
import { buildWorksheet, createWorksheetObj, getNextDefaultWorksheetName } from './worksheets.js';
import dispatch from './dispatch.js';
import { createFromTable } from './helpers.js';
import { createTabs } from './tabs.js';
import { getSpreadsheetConfig, setConfig } from './config.js';
import type { Spreadsheet, Worksheet } from '../types/internal';

const factory = function () {};

const createWorksheets = async function (spreadsheet: Spreadsheet, options: any, el: HTMLElement) {
    // Create worksheets
    let o = options.worksheets;
    if (o) {
        let tabsOptions: any = {
            animation: true,
            onbeforecreate: function (element: any, title: any) {
                if (title) {
                    return title;
                } else {
                    return getNextDefaultWorksheetName(spreadsheet);
                }
            },
            oncreate: function (element: any, newTabContent: any) {
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

                newWorksheet.element = newTabContent;

                buildWorksheet.call(newWorksheet).then(function () {
                    updateToolbar(newWorksheet);

                    dispatch.call(newWorksheet, 'oncreateworksheet', newWorksheet, options, spreadsheet.worksheets.length - 1);
                });
            },
            onchange: function (element: any, instance: any, tabIndex: number) {
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
                title: o[i].worksheetName,
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
                Object.entries(o[i].style).forEach(function ([cellName, value]) {
                    if (typeof value === 'number') {
                        o[i].style[cellName] = spreadsheetStyles[value];
                    }
                });
            }

            spreadsheet.worksheets.push({
                parent: spreadsheet,
                element: tabs.content.children[i],
                options: o[i],
                filters: [],
                formula: [],
                history: [],
                selection: [],
                historyIndex: -1,
            } as any);

            await buildWorksheet.call(spreadsheet.worksheets[i]);
        }
    } else {
        throw new Error('JSS: worksheets are not defined');
    }
};

factory.spreadsheet = async function (el: any, options: any, worksheets: Worksheet[]) {
    if (el.tagName == 'TABLE') {
        if (!options) {
            options = {};
        }

        if (!options.worksheets) {
            options.worksheets = [];
        }

        const tableOptions = createFromTable(el, options.worksheets[0]);

        options.worksheets[0] = tableOptions;

        const div = document.createElement('div');
        el.parentNode.insertBefore(div, el);
        el.remove();
        el = div;
    }

    let spreadsheet = {
        worksheets: worksheets,
        config: options,
        element: el,
        el,
    } as Spreadsheet;

    // Contextmenu container
    spreadsheet.contextMenu = document.createElement('div');
    spreadsheet.contextMenu.className = 'jss_contextmenu';

    spreadsheet.getWorksheetActive = getWorksheetActive.bind(spreadsheet);
    spreadsheet.fullscreen = fullscreen.bind(spreadsheet);
    spreadsheet.showToolbar = showToolbar.bind(spreadsheet);
    spreadsheet.hideToolbar = hideToolbar.bind(spreadsheet);
    spreadsheet.getConfig = getSpreadsheetConfig.bind(spreadsheet);
    spreadsheet.setConfig = setConfig.bind(spreadsheet);

    spreadsheet.setPlugins = function (newPlugins: any) {
        if (!spreadsheet.plugins) {
            spreadsheet.plugins = {};
        }

        if (typeof newPlugins == 'object') {
            Object.entries(newPlugins).forEach(function ([pluginName, plugin]: [string, any]) {
                spreadsheet.plugins[pluginName] = plugin.call(libraryBase.jspreadsheet, spreadsheet, {}, spreadsheet.config);
            });
        }
    };

    spreadsheet.setPlugins(options.plugins);

    // Create as worksheets
    await createWorksheets(spreadsheet, options, el);

    spreadsheet.element.appendChild(spreadsheet.contextMenu);

    // Create element
    jSuites.contextmenu(spreadsheet.contextMenu, {
        onclick: function () {
            (spreadsheet.contextMenu as any).contextmenu.close(false);
        },
    } as any);

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

factory.worksheet = function (spreadsheet: Spreadsheet, options: any, position?: number) {
    // Worksheet object
    let w: any = {
        // Parent of a worksheet is always the spreadsheet
        parent: spreadsheet,
        // Options for this worksheet
        options: {},
    };

    // Create the worksheets object
    if (typeof position === 'undefined') {
        spreadsheet.worksheets.push(w);
    } else {
        spreadsheet.worksheets.splice(position, 0, w);
    }
    // Keep configuration used
    Object.assign(w.options, options);

    return w;
};

export default factory;
