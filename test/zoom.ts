import { expect } from 'chai';

import jspreadsheet, { type Jspreadsheet, type Worksheet } from './support/jspreadsheet';

// The spreadsheet adds its event handlers after a promise
function load(options: Jspreadsheet.SpreadsheetOptions) {
    return new Promise<Worksheet[]>((resolve) => {
        const instance = jspreadsheet(root, { ...options, onload: () => resolve(instance) });
    });
}

// jsdom has no layout: an element gets its place on the screen from this rectangle
function rect(element: HTMLElement, left: number, top: number, width: number, height: number) {
    element.getBoundingClientRect = () => ({ left, top, width, height, right: left + width, bottom: top + height, x: left, y: top }) as DOMRect;
}

function mouse(element: Element, type: string, init: MouseEventInit & { offsetX?: number; pageX?: number } = {}) {
    const { offsetX = 0, ...rest } = init;
    const event = new window.MouseEvent(type, { bubbles: true, cancelable: true, buttons: 1, ...rest });
    Object.defineProperty(event, 'offsetX', { value: offsetX });
    Object.defineProperty(event, 'pageX', { value: rest.pageX || 0 });
    element.dispatchEvent(event);
}

describe('Zoom', () => {
    it('puts the table, the corner and the clipboard helper in a wrapper with the zoom of the option', () => {
        const [worksheet] = jspreadsheet(root, { zoom: 1.5, worksheets: [{ minDimensions: [2, 2] }] });

        expect(Array.from(worksheet.content.children)).to.eql([worksheet.zoomed]);
        expect(worksheet.zoomed.className).to.equal('jss_zoom');
        expect(worksheet.zoomed.style.zoom).to.equal('1.5');
        expect(worksheet.zoomed.style.getPropertyValue('--jss-zoom')).to.equal('1.5');
    });

    it('changes the zoom of all worksheets with setZoom, and gives 1 for a value that is not a number above 0', async () => {
        const worksheets = await load({ worksheets: [{ minDimensions: [2, 2] }, { minDimensions: [2, 2] }] });
        expect(worksheets[0].zoomed.style.zoom).to.equal('');

        worksheets[0].setZoom(2);
        expect(worksheets.map((worksheet) => worksheet.zoomed.style.zoom)).to.eql(['2', '2']);
        expect(worksheets[0].parent.config.zoom).to.equal(2);

        for (const value of [0, -1, NaN, 'x']) {
            worksheets[1].setZoom(value as number);
            expect(worksheets.map((worksheet) => worksheet.zoomed.style.zoom)).to.eql(['', '']);
            expect(worksheets[0].parent.config.zoom).to.equal(1);
        }
    });

    it('gives the cell editor the size of the cell in the grid', () => {
        const [worksheet] = jspreadsheet(root, { zoom: 1.5, worksheets: [{ data: [['a']] }] });
        const cell = worksheet.getCell('A1');
        rect(cell, 0, 0, 150, 39);

        worksheet.openEditor(cell);

        const editor = cell.children[0] as HTMLElement;
        expect([editor.style.width, editor.style.height]).to.eql(['100px', '24px']);
    });

    it('places the selection corner at the lower-right corner of the selection in the grid', () => {
        const [worksheet] = jspreadsheet(root, { zoom: 2, worksheets: [{ data: [['a', 'b']], tableOverflow: true }] });
        rect(worksheet.content, 10, 10, 400, 300);
        rect(worksheet.getCell('B1'), 210, 60, 200, 50);
        Object.defineProperty(worksheet.content, 'scrollLeft', { value: 20, writable: true });

        worksheet.updateSelectionFromCoords(1, 0, 1, 0);

        // (210 - 10 + 20 + 200) / 2 - 4 and (60 - 10 + 50) / 2 - 4
        expect([worksheet.corner.style.left, worksheet.corner.style.top]).to.eql(['206px', '46px']);
    });

    it('gives a column the width of the move of the mouse in the grid', async () => {
        const [worksheet] = await load({ zoom: 2, worksheets: [{ minDimensions: [2, 2], columns: [{ width: 100 }, { width: 100 }] }] });
        rect(worksheet.headers[0], 0, 0, 200, 40);

        mouse(worksheet.headers[0], 'mousedown', { offsetX: 198, pageX: 198 });
        mouse(worksheet.headers[0], 'mousemove', { pageX: 258 });
        mouse(worksheet.headers[0], 'mouseup');

        expect(worksheet.getWidth(0)).to.equal(130);
    });

    it('freezes the columns when the scroll on the screen is more than the index column on the screen', () => {
        const [worksheet] = jspreadsheet(root, {
            zoom: 2,
            worksheets: [{ minDimensions: [4, 2], freezeColumns: 1, tableOverflow: true, columns: [{ width: 100 }] }],
        });
        Object.defineProperty(worksheet.content, 'scrollLeft', { value: 80, writable: true });

        // 80 screen pixels are 40 grid pixels, less than the index column of 50
        worksheet.columns.updateFreezePosition();
        expect(worksheet.headers[0].classList.contains('jss_freezed')).to.equal(false);

        worksheet.content.scrollLeft = 300;
        worksheet.columns.updateFreezePosition();
        expect(worksheet.headers[0].classList.contains('jss_freezed')).to.equal(true);
        // 300 / 2 - 50 - 1
        expect(worksheet.records[1][0].element.style.left).to.equal('99px');
    });
});
