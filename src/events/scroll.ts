import libraryBase from '../utils/libraryBase.js';
import type { Worksheet } from '../types/internal';

export const wheelControls = function (this: Worksheet, e?: WheelEvent) {
    const obj = this;

    if (obj.options.lazyLoading == true) {
        if (libraryBase.jspreadsheet.timeControlLoading == null) {
            libraryBase.jspreadsheet.timeControlLoading = setTimeout(function () {
                if (obj.content.scrollTop + obj.content.clientHeight >= obj.content.scrollHeight - 10) {
                    if (obj.view.loadDown()) {
                        if (obj.content.scrollTop + obj.content.clientHeight > obj.content.scrollHeight - 10) {
                            obj.content.scrollTop = obj.content.scrollTop - obj.content.clientHeight;
                        }
                        obj.selection.updateCorner();
                    }
                } else if (obj.content.scrollTop <= obj.content.clientHeight) {
                    if (obj.view.loadUp()) {
                        if (obj.content.scrollTop < 10) {
                            obj.content.scrollTop = obj.content.scrollTop + obj.content.clientHeight;
                        }
                        obj.selection.updateCorner();
                    }
                }

                libraryBase.jspreadsheet.timeControlLoading = null;
            }, 100);
        }
    }
};

export const scrollControls = function (this: Worksheet, e: Event) {
    const obj = this;

    wheelControls.call(obj);

    if (obj.options.freezeColumns! > 0 && obj.content.scrollLeft != obj.lastScrollLeft) {
        obj.lastScrollLeft = obj.content.scrollLeft;
        obj.columns.updateFreezePosition();
    }

    // Close a floating editor. The text editor is inside the cell and scrolls with it.
    if (obj.options.lazyLoading == true || obj.options.tableOverflow == true) {
        if (obj.editor.isOpen() && hasFloatingEditor.call(obj, obj.editor.x) && (e.target as HTMLElement).className.substr(0, 9) != 'jdropdown') {
            obj.editor.close(obj.editor.cell, true);
        }
    }
};

/**
 * Check if the editor of a column floats over the page. It does not move with the cell.
 */
const hasFloatingEditor = function (this: Worksheet, x: number) {
    const obj = this;

    const column = obj.options.columns && obj.options.columns[x];
    if (!column || !column.type) {
        return false;
    }

    return typeof column.type === 'object' || ['dropdown', 'calendar', 'color', 'html', 'image'].indexOf(column.type) >= 0;
};
