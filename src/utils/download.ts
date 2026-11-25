import { copy } from './copyPaste.js';
import type { Worksheet } from '../types/internal';

/**
 * Download CSV table
 * @param includeHeaders - Put the headers in the first row.
 * @param processed - Download the shown values instead of the data.
 */
export const download = function (this: Worksheet, includeHeaders?: boolean, processed?: boolean) {
    const obj = this;

    if (obj.parent.config.allowExport == false) {
        console.error('Export not allowed');
    } else {
        // Data
        let data = '';

        // Get data
        data += copy.call(obj, false, obj.options.csvDelimiter, true, includeHeaders, true, undefined, processed);

        // Download element
        const blob = new Blob(['\uFEFF' + data], { type: 'text/csv;charset=utf-8;' });

        // IE Compatibility
        if (window.navigator && window.navigator.msSaveOrOpenBlob) {
            window.navigator.msSaveOrOpenBlob(blob, (obj.options.csvFileName || obj.options.worksheetName) + '.csv');
        } else {
            // Download element
            const pom = document.createElement('a');
            pom.setAttribute('target', '_top');
            const url = URL.createObjectURL(blob);
            pom.href = url;
            pom.setAttribute('download', (obj.options.csvFileName || obj.options.worksheetName) + '.csv');
            document.body.appendChild(pom);
            pom.click();
            pom.parentNode!.removeChild(pom);
        }
    }
};
