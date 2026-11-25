// @ts-check
import jSuites from 'jsuites';
/** @import { Worksheet, Spreadsheet } from '../types/internal' */

/**
 * Prepare JSON in the correct format
 * @this {Worksheet}
 * @param {{ x: number, y: number, value: any }[]} data - The changed cells.
 */
const prepareJson = function (data) {
    const obj = this;

    /** @type {{ row: number, data: Record<string, any> }[]} */
    const rows = [];
    for (let i = 0; i < data.length; i++) {
        const x = data[i].x;
        const y = data[i].y;
        const k = obj.options.columns[x].name ? obj.options.columns[x].name : x;

        // Create row
        if (!rows[y]) {
            rows[y] = {
                row: y,
                data: {},
            };
        }
        rows[y].data[k] = data[i].value;
    }

    // Filter rows
    return rows.filter(function (el) {
        return el != null;
    });
};

/**
 * Post json to a remote server
 * @this {Worksheet}
 * @param {string} url
 * @param {any} data
 */
const save = function (url, data) {
    const obj = this;

    // Parse anything in the data before sending to the server
    const ret = dispatch.call(obj.parent, 'onbeforesave', obj.parent, obj, data);
    if (ret) {
        data = ret;
    } else {
        if (ret === false) {
            return false;
        }
    }

    // Remove update
    jSuites.ajax({
        url: url,
        method: 'POST',
        dataType: 'json',
        data: { data: JSON.stringify(data) },
        /** @param {any} result */
        success: function (result) {
            // Event
            dispatch.call(obj, 'onsave', obj.parent, obj, data);
        },
    });
};

/**
 * Trigger events
 *
 * The first argument is the name of the event, for example "onchange". The other arguments go to the event handlers.
 * It returns the result of the last handler.
 *
 * @type {(this: Worksheet | Spreadsheet, event: string, ...args: any[]) => any}
 */
const dispatch = function (event) {
    const obj = this;
    let ret = null;

    let spreadsheet = /** @type {Spreadsheet} */ (obj.parent ? obj.parent : obj);

    // Dispatch events
    if (!spreadsheet.ignoreEvents) {
        // Call global event
        if (typeof spreadsheet.config.onevent == 'function') {
            ret = spreadsheet.config.onevent.apply(this, /** @type {any} */ (arguments));
        }
        // Call specific events
        /** @type {Record<string, any>} */
        const config = spreadsheet.config;
        if (typeof config[event] == 'function') {
            ret = config[event].apply(this, Array.prototype.slice.call(arguments, 1));
        }

        if (typeof spreadsheet.plugins === 'object') {
            const pluginKeys = Object.keys(spreadsheet.plugins);

            for (let pluginKeyIndex = 0; pluginKeyIndex < pluginKeys.length; pluginKeyIndex++) {
                const key = pluginKeys[pluginKeyIndex];
                const plugin = spreadsheet.plugins[key];

                if (typeof plugin.onevent === 'function') {
                    ret = plugin.onevent.apply(this, /** @type {any} */ (arguments));
                }
            }
        }
    }

    if (event == 'onafterchanges') {
        // Only a worksheet dispatches onafterchanges
        const worksheet = /** @type {Worksheet} */ (obj);
        const scope = arguments;

        if (typeof spreadsheet.plugins === 'object') {
            Object.entries(spreadsheet.plugins).forEach(function ([, plugin]) {
                if (typeof plugin.persistence === 'function') {
                    plugin.persistence(worksheet, 'setValue', { data: scope[2] });
                }
            });
        }

        if (worksheet.options.persistence) {
            const url = worksheet.options.persistence == true ? worksheet.options.url : worksheet.options.persistence;
            const data = prepareJson.call(worksheet, arguments[2]);
            save.call(worksheet, url, data);
        }
    }

    return ret;
};

export default dispatch;
