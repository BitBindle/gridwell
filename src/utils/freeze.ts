import type { Worksheet } from '../types/internal';
// Get the width of the index column. The frozen columns stay at the left of the content when scrollLeft is more than this width.
export const getIndexWidth = function (this: Worksheet) {
    return this.table.classList.contains('jss_hidden_index') ? 0 : 50;
};

// Get width of all freezed cells together
export const getFreezeWidth = function (this: Worksheet) {
    const obj = this;

    let width = 0;
    if (obj.options.freezeColumns! > 0) {
        for (let i = 0; i < obj.options.freezeColumns!; i++) {
            let columnWidth;
            if (obj.options.columns && obj.options.columns[i] && obj.options.columns[i].width !== undefined) {
                columnWidth = parseInt(obj.options.columns[i].width);
            } else {
                columnWidth = obj.options.defaultColWidth !== undefined ? parseInt(obj.options.defaultColWidth) : 100;
            }

            width += columnWidth;
        }
    }
    return width;
};
