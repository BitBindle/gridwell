import jSuites from 'jsuites';
import type jspreadsheet from '../jspreadsheet';
import type { Worksheet, Spreadsheet, Dispatch, EventName } from '../types/internal';

/** The rows that persistence posts to the server, the data of onbeforesave */
type SaveData = Parameters<NonNullable<jspreadsheet.SpreadsheetOptions['onbeforesave']>>[2];

/** A record of onafterchanges. The coordinates can be text (issue #52), and a read-only cell has no value (issue #58). */
type ChangeRecord = { x: number; y: number; value?: jspreadsheet.CellValue };

/**
 * Prepare JSON in the correct format
 * @param data - The changed cells.
 */
const prepareJson = function (this: Worksheet, data: ChangeRecord[]) {
    const obj = this;

    const rows: SaveData = [];
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
        // A read-only cell has no value (issue #58)
        rows[y].data[k] = data[i].value!;
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
const save = function (this: Worksheet, url: string | undefined, data: SaveData) {
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
        success: function (result: unknown) {
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
const dispatch = function (this: Worksheet | Spreadsheet, event: EventName, ...args: unknown[]) {
    const obj = this;
    let ret: unknown = null;

    const spreadsheet = (obj.parent ? obj.parent : obj) as Spreadsheet;

    // Dispatch events
    if (!spreadsheet.ignoreEvents) {
        // Call global event
        if (typeof spreadsheet.config.onevent == 'function') {
            ret = spreadsheet.config.onevent.call(this, event, ...args);
        }
        // Call specific events
        const config = spreadsheet.config;
        if (typeof config[event] == 'function') {
            // The arguments match the handler of the event, as the type of dispatch says
            ret = (config[event] as (...args: unknown[]) => unknown).apply(this, args);
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
            const data = prepareJson.call(worksheet, args[1] as ChangeRecord[]);
            save.call(worksheet, url, data);
        }
    }

    return ret;
} as Dispatch;

export default dispatch;
