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

function mouse(element: Element, type: string, init: MouseEventInit & { offsetX?: number; offsetY?: number; pageX?: number; pageY?: number } = {}) {
    const { offsetX = 0, offsetY = 0, ...rest } = init;
    const event = new window.MouseEvent(type, { bubbles: true, cancelable: true, buttons: 1, ...rest });
    Object.defineProperty(event, 'offsetX', { value: offsetX });
    Object.defineProperty(event, 'offsetY', { value: offsetY });
    Object.defineProperty(event, 'pageX', { value: rest.pageX || 0 });
    Object.defineProperty(event, 'pageY', { value: rest.pageY || 0 });
    element.dispatchEvent(event);
}

// The size of an element in the grid, which clientWidth and clientHeight give
function client(element: HTMLElement, width: number, height: number) {
    Object.defineProperty(element, 'clientWidth', { configurable: true, value: width });
    Object.defineProperty(element, 'clientHeight', { configurable: true, value: height });
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

    it('gives a row the height of the move of the mouse in the grid, and keeps its old height in the grid for the undo', async () => {
        const [worksheet] = await load({ zoom: 2, worksheets: [{ minDimensions: [2, 2] }] });
        const rowNumber = worksheet.rows[0].element.children[0] as HTMLElement;
        rect(rowNumber, 0, 0, 100, 50);

        mouse(rowNumber, 'mousedown', { offsetY: 48, pageY: 48 });
        mouse(rowNumber, 'mousemove', { pageY: 88 });
        mouse(rowNumber, 'mouseup');

        expect(worksheet.rows[0].element.getAttribute('height')).to.equal('45');

        // Without a height attribute, the old height comes from the screen
        worksheet.rows[1].element.removeAttribute('height');
        rect(worksheet.rows[1].element, 0, 0, 100, 50);
        worksheet.setHeight(1, 40);
        worksheet.undo();
        expect(worksheet.getHeight(1)).to.equal('25px');
    });

    it('puts a dragged column before the header when the mouse is over its left half in the grid', async () => {
        const [worksheet] = await load({ zoom: 2, worksheets: [{ minDimensions: [4, 1], columnResize: false }] });
        for (const header of worksheet.headers) {
            rect(header, 0, 0, 200, 40);
            client(header, 100, 20);
        }

        mouse(worksheet.headers[0], 'mousedown', { offsetY: 38 });
        // 80 screen pixels are 40 grid pixels, in the left half of a column of 100 grid pixels
        mouse(worksheet.headers[2], 'mousemove', { offsetX: 80 });

        expect(worksheet.pointer.dragging!.destination).to.equal(1);
        expect(worksheet.headers[2].classList.contains('dragging-left')).to.equal(true);
        mouse(worksheet.headers[2], 'mouseup');
    });

    it('scrolls a cell that the keys select into the view on the screen, to the right of the frozen column', () => {
        const [worksheet] = jspreadsheet(root, {
            zoom: 2,
            worksheets: [{ minDimensions: [10, 2], freezeColumns: 1, tableOverflow: true, columns: [{ width: 100 }] }],
        });
        rect(worksheet.content, 0, 0, 600, 300);
        Object.defineProperty(worksheet.content, 'scrollLeft', { value: 400, writable: true });
        // Column C starts 150 screen pixels right of the content: under the frozen column of 200 screen pixels
        rect(worksheet.records[0][2].element, 150, 40, 200, 50);
        worksheet.updateSelectionFromCoords(3, 0, 3, 0);

        worksheet.left();

        // 150 + 400 - 200 - 1: the cell starts right of the frozen column
        expect(worksheet.content.scrollLeft).to.equal(349);
    });

    it('keeps the same cells in view after setZoom, and saves and closes a pop-up editor', async () => {
        const [worksheet] = await load({
            zoom: 2,
            worksheets: [{ data: [['a', 'x']], tableOverflow: true, columns: [{ type: 'dropdown', source: ['a', 'b'] }, {}] }],
        });
        Object.defineProperty(worksheet.content, 'scrollLeft', { value: 300, writable: true });
        Object.defineProperty(worksheet.content, 'scrollTop', { value: 100, writable: true });
        worksheet.openEditor(worksheet.getCell('A1'));
        expect(worksheet.editor.isOpen()).to.equal(true);

        worksheet.setZoom(1);

        expect(worksheet.editor.isOpen()).to.equal(false);
        expect([worksheet.content.scrollLeft, worksheet.content.scrollTop]).to.eql([150, 50]);
    });
});
