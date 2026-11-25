import jSuites from 'jsuites';
import type { Worksheet, Spreadsheet } from '../types/internal';

/**
 * Prepare JSON in the correct format
 * @param data - The changed cells.
 */
const prepareJson = function (this: Worksheet, data: { x: number; y: number; value: any }[]) {
    const obj = this;

    const rows: { row: number; data: Record<string, any> }[] = [];
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
 */
// The url is undefined when persistence is true and the worksheet has no url option
const save = function (this: Worksheet, url: string | undefined, data: any) {
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
        success: function (result: any) {
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
 */
const dispatch = function (this: Worksheet | Spreadsheet, event: string, ...args: any[]): any {
    const obj = this;
    let ret = null;

    const spreadsheet = (obj.parent ? obj.parent : obj) as Spreadsheet;

    // Dispatch events
    if (!spreadsheet.ignoreEvents) {
        // Call global event
        if (typeof spreadsheet.config.onevent == 'function') {
            ret = spreadsheet.config.onevent.call(this, event, ...args);
        }
        // Call specific events
        const config: Record<string, any> = spreadsheet.config;
        if (typeof config[event] == 'function') {
            ret = config[event].apply(this, args);
        }

        if (typeof spreadsheet.plugins === 'object') {
            const pluginKeys = Object.keys(spreadsheet.plugins);

            for (let pluginKeyIndex = 0; pluginKeyIndex < pluginKeys.length; pluginKeyIndex++) {
                const key = pluginKeys[pluginKeyIndex];
                const plugin = spreadsheet.plugins[key];

                if (typeof plugin.onevent === 'function') {
                    ret = plugin.onevent.call(this, event, ...args);
                }
            }
        }
    }

    if (event == 'onafterchanges') {
        // Only a worksheet dispatches onafterchanges
        const worksheet = obj as Worksheet;
        if (typeof spreadsheet.plugins === 'object') {
            Object.entries(spreadsheet.plugins).forEach(function ([, plugin]) {
                if (typeof plugin.persistence === 'function') {
                    plugin.persistence(worksheet, 'setValue', { data: args[1] });
                }
            });
        }

        if (worksheet.options.persistence) {
            const url = worksheet.options.persistence == true ? worksheet.options.url : worksheet.options.persistence;
            const data = prepareJson.call(worksheet, args[1]);
            save.call(worksheet, url, data);
        }
    }

    return ret;
};

export default dispatch;
