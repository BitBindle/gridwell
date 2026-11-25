import jSuites from 'jsuites';
import { updateResult } from './internal.js';
import { refreshSelection } from './selection.js';
import type { Worksheet, WorksheetMethod } from '../types/internal';

/**
 * Open the column filter
 */
export const openFilter: WorksheetMethod<'openFilter'> = function (columnId) {
    const obj = this;

    if (!obj.options.filters) {
        console.log('Jspreadsheet: filters not enabled.');
    } else {
        // Make sure is integer
        columnId = parseInt(columnId);
        // Reset selection
        obj.resetSelection();
        // Load options
        let optionsFiltered = [];
        if (obj.options.columns[columnId].type == 'checkbox') {
            optionsFiltered.push({ id: 'true', name: 'True' });
            optionsFiltered.push({ id: 'false', name: 'False' });
        } else {
            // Labels by value. An array used as a map, as before.
            const options = [] as unknown as Record<string, string>;
            let hasBlanks = false;
            for (let j = 0; j < obj.options.data.length; j++) {
                const k = obj.options.data[j][columnId];
                const v = obj.records[j][columnId].element.innerHTML;
                if (k && v) {
                    options[k as string] = v;
                } else {
                    hasBlanks = true;
                }
            }
            const keys = Object.keys(options);
            optionsFiltered = [];
            for (let j = 0; j < keys.length; j++) {
                optionsFiltered.push({ id: keys[j], name: options[keys[j]] });
            }
            // Has blank options
            if (hasBlanks) {
                optionsFiltered.push({ value: '', id: '', name: '(Blanks)' });
            }
        }

        // Create dropdown
        const div = document.createElement('div');
        obj.filter!.children[columnId + 1].innerHTML = '';
        obj.filter!.children[columnId + 1].appendChild(div);
        (obj.filter!.children[columnId + 1] as HTMLElement).style.paddingLeft = '0px';
        (obj.filter!.children[columnId + 1] as HTMLElement).style.paddingRight = '0px';
        (obj.filter!.children[columnId + 1] as HTMLElement).style.overflow = 'initial';

        const opt = {
            data: optionsFiltered,
            multiple: true,
            autocomplete: true,
            opened: true,
            value: obj.filters[columnId] !== undefined ? obj.filters[columnId] : null,
            width: '100%',
            position: obj.options.tableOverflow == true || obj.parent.config.fullscreen == true ? true : false,
            // The element of the dropdown, which holds the dropdown
            onclose: function (o: HTMLElement & { dropdown: ReturnType<typeof jSuites.dropdown> }) {
                resetFilters.call(obj);
                // A list, because the dropdown has multiple
                obj.filters[columnId] = o.dropdown.getValue(true) as string[];
                obj.filter!.children[columnId + 1].innerHTML = o.dropdown.getText() as string;
                (obj.filter!.children[columnId + 1] as HTMLElement).style.paddingLeft = '';
                (obj.filter!.children[columnId + 1] as HTMLElement).style.paddingRight = '';
                (obj.filter!.children[columnId + 1] as HTMLElement).style.overflow = '';
                closeFilter.call(obj, columnId);
                refreshSelection.call(obj);
            },
        };

        // Dynamic dropdown
        jSuites.dropdown(div, opt);
    }
};

export const closeFilter = function (this: Worksheet, columnId?: number) {
    const obj = this;

    if (!columnId) {
        for (let i = 0; i < obj.filter!.children.length; i++) {
            if (obj.filters[i]) {
                columnId = i;
            }
        }
    }

    // Search filter
    const search = function (query: string[], x: number, y: number) {
        for (let i = 0; i < query.length; i++) {
            const value = '' + obj.options.data[y][x];
            const label = '' + obj.records[y][x].element.innerHTML;
            if (query[i] == value || query[i] == label) {
                return true;
            }
        }
        return false;
    };

    // The column has a filter: openFilter passes its column, and orderBy calls this only while a filter gives results
    const query = obj.filters[columnId!]!;
    obj.results = [];
    for (let j = 0; j < obj.options.data.length; j++) {
        if (search(query, columnId!, j)) {
            obj.results.push(j);
        }
    }
    if (!obj.results.length) {
        obj.results = null;
    }

    updateResult.call(obj);
};

export const resetFilters: WorksheetMethod<'resetFilters'> = function () {
    const obj = this;

    if (obj.options.filters) {
        for (let i = 0; i < obj.filter!.children.length; i++) {
            obj.filter!.children[i].innerHTML = '&nbsp;';
            obj.filters[i] = null;
        }
    }

    obj.results = null;
    updateResult.call(obj);
};
