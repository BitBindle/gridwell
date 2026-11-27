import libraryBase from '../utils/libraryBase.js';

/**
 * Check if the scrolled node contains the element. The check goes through the hosts of the shadow roots, because
 * Node.contains does not cross them.
 */
function scrolls(target: Node, element: Node) {
    let node: Node | null = element;
    while (node) {
        if (target.contains(node)) {
            return true;
        }
        node = (node.getRootNode() as ShadowRoot).host ?? null;
    }
    return false;
}

/**
 * A scroll of the page, or of an element around a worksheet, moves the cells under the sticky header rows and the
 * frozen rows. Each worksheet in the scrolled element checks again if they cover its fill handle. View.scroll checks it
 * for a scroll of the content of the worksheet.
 */
export function scrollControls(e: Event) {
    const target = e.target as Node | null;
    if (!target) {
        return;
    }
    for (const spreadsheet of libraryBase.jspreadsheet.spreadsheet) {
        for (const worksheet of spreadsheet.worksheets) {
            if (target !== worksheet.content && scrolls(target, worksheet.element)) {
                worksheet.selection.updateCornerCover();
            }
        }
    }
}
