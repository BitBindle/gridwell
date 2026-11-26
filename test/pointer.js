const { expect } = require('chai');

const jspreadsheet = require('../dist/index.js');

describe('Mouse gestures', () => {
    // The spreadsheet adds its event handlers after a promise
    const load = (options) =>
        new Promise((resolve) => {
            const instance = jspreadsheet(root, { ...options, onload: () => resolve(instance[0]) });
        });

    // jsdom has no layout: the event gets the offsets and the positions, and the elements get their size with rect
    const mouse = (element, type, init = {}) => {
        const { offsetX = 0, offsetY = 0, ...rest } = init;
        const event = new window.MouseEvent(type, { bubbles: true, cancelable: true, buttons: 1, ...rest });
        Object.defineProperty(event, 'offsetX', { value: offsetX });
        Object.defineProperty(event, 'offsetY', { value: offsetY });
        Object.defineProperty(event, 'pageX', { value: rest.pageX || 0 });
        Object.defineProperty(event, 'pageY', { value: rest.pageY || 0 });
        element.dispatchEvent(event);
    };

    const rect = (element, width, height) => {
        element.getBoundingClientRect = () => ({ left: 0, top: 0, right: width, bottom: height, width, height });
        Object.defineProperty(element, 'clientWidth', { configurable: true, value: width });
        Object.defineProperty(element, 'clientHeight', { configurable: true, value: height });
    };

    const rowNumber = (worksheet, y) => worksheet.rows[y].element.children[0];

    let error;
    let errors;

    beforeEach(() => {
        error = console.error;
        errors = [];
        console.error = (message) => errors.push(message);
    });

    afterEach(() => {
        console.error = error;
        // A gesture that a test leaves open does not go to the next test
        mouse(document.body, 'mouseup');
    });

    describe('resize', () => {
        it('sets the width of a column that the mouse resizes, with the history', async () => {
            const worksheet = await load({ worksheets: [{ minDimensions: [3, 2], columns: [{ width: 100 }, { width: 100 }, { width: 100 }] }] });
            rect(worksheet.headers[0], 100, 20);

            mouse(worksheet.headers[0], 'mousedown', { offsetX: 98, pageX: 98 });
            expect(worksheet.headers[0].classList.contains('resizing')).to.equal(true);

            mouse(worksheet.headers[0], 'mousemove', { pageX: 148 });
            mouse(worksheet.headers[0], 'mouseup');

            expect(worksheet.getWidth(0)).to.equal(150);
            expect(worksheet.headers[0].classList.contains('resizing')).to.equal(false);

            worksheet.undo();
            expect(worksheet.getWidth(0)).to.equal(100);
        });

        it('does not make a column narrower than nothing', async () => {
            const worksheet = await load({ worksheets: [{ minDimensions: [1, 1], columns: [{ width: 100 }] }] });
            rect(worksheet.headers[0], 100, 20);

            mouse(worksheet.headers[0], 'mousedown', { offsetX: 98, pageX: 98 });
            mouse(worksheet.headers[0], 'mousemove', { pageX: -50 });
            mouse(worksheet.headers[0], 'mouseup');

            expect(worksheet.getWidth(0)).to.equal(100);
        });

        it('sets the width of all selected columns', async () => {
            const worksheet = await load({ worksheets: [{ minDimensions: [3, 2], columns: [{ width: 100 }, { width: 100 }, { width: 100 }] }] });

            mouse(worksheet.headers[0], 'mousedown', { offsetX: -10, offsetY: -10 });
            mouse(worksheet.headers[0], 'mouseup');
            mouse(worksheet.headers[1], 'mousedown', { offsetX: -10, offsetY: -10, shiftKey: true });
            mouse(worksheet.headers[1], 'mouseup');
            expect(worksheet.getSelectedColumns()).to.eql([0, 1]);

            rect(worksheet.headers[1], 100, 20);
            mouse(worksheet.headers[1], 'mousedown', { offsetX: 98, pageX: 98 });
            mouse(worksheet.headers[1], 'mousemove', { pageX: 118 });
            mouse(worksheet.headers[1], 'mouseup');

            expect(worksheet.getWidth()).to.eql([120, 120, 100]);
        });

        it('sets the height of a row that the mouse resizes', async () => {
            const worksheet = await load({ worksheets: [{ minDimensions: [2, 2] }] });
            rect(rowNumber(worksheet, 0), 50, 25);

            mouse(rowNumber(worksheet, 0), 'mousedown', { offsetY: 23, pageY: 23 });
            expect(worksheet.rows[0].element.classList.contains('resizing')).to.equal(true);

            mouse(rowNumber(worksheet, 0), 'mousemove', { pageY: 38 });
            mouse(rowNumber(worksheet, 0), 'mousemove', { pageY: -100 });
            mouse(rowNumber(worksheet, 0), 'mouseup');

            expect(worksheet.getHeight(0)).to.equal('40px');
            expect(worksheet.rows[0].element.classList.contains('resizing')).to.equal(false);
        });
    });

    describe('move', () => {
        const data = () => [
            ['a', 1],
            ['b', 2],
            ['c', 3],
        ];

        it('moves a row that the mouse drags', async () => {
            const worksheet = await load({ worksheets: [{ data: data(), rowResize: false }] });
            rect(rowNumber(worksheet, 0), 50, 25);

            mouse(rowNumber(worksheet, 0), 'mousedown', { offsetX: 48 });
            expect(worksheet.rows[0].element.classList.contains('dragging')).to.equal(true);

            // As soon as the mouse is over the upper half of a row below, the row goes after it
            rect(worksheet.records[2][0].element, 100, 25);
            mouse(worksheet.records[2][0].element, 'mousemove', { offsetY: 5 });
            mouse(worksheet.records[2][0].element, 'mouseup');

            expect(worksheet.getColumnData(0)).to.eql(['b', 'c', 'a']);
        });

        it('moves a row up to the lower half of a row', async () => {
            const worksheet = await load({ worksheets: [{ data: data(), rowResize: false }] });
            rect(rowNumber(worksheet, 2), 50, 25);

            mouse(rowNumber(worksheet, 2), 'mousedown', { offsetX: 48 });

            rect(worksheet.records[0][1].element, 100, 25);
            mouse(worksheet.records[0][1].element, 'mousemove', { offsetY: 20 });
            mouse(worksheet.records[0][1].element, 'mouseup');

            // Over the lower half of a row above, the row goes before it
            expect(worksheet.getColumnData(0)).to.eql(['c', 'a', 'b']);
        });

        it('does not move a row of a merged cell', async () => {
            const worksheet = await load({ worksheets: [{ data: data(), rowResize: false, mergeCells: { A1: [1, 2] } }] });
            rect(rowNumber(worksheet, 0), 50, 25);

            mouse(rowNumber(worksheet, 0), 'mousedown', { offsetX: 48 });

            expect(worksheet.pointer.dragging).to.equal(null);
            expect(errors).to.eql(['Jspreadsheet: This row is part of a merged cell']);
        });

        it('does not move a row over a merged cell', async () => {
            const worksheet = await load({ worksheets: [{ data: data(), rowResize: false, mergeCells: { A2: [1, 2] } }] });
            rect(rowNumber(worksheet, 0), 50, 25);

            mouse(rowNumber(worksheet, 0), 'mousedown', { offsetX: 48 });
            mouse(worksheet.records[1][1].element, 'mousemove');
            mouse(worksheet.records[1][1].element, 'mouseup');

            expect(errors).to.eql(['Jspreadsheet: This row is part of a merged cell.']);
            // The merged cell keeps the value of its first cell
            expect(worksheet.getColumnData(0)).to.eql(['a', 'b', '']);
        });

        it('does not move a row while a search shows its results', async () => {
            const worksheet = await load({ worksheets: [{ data: data(), rowResize: false, search: true }] });
            worksheet.search('b');
            rect(rowNumber(worksheet, 1), 50, 25);

            mouse(rowNumber(worksheet, 1), 'mousedown', { offsetX: 48 });

            expect(worksheet.pointer.dragging).to.equal(null);
            expect(errors).to.eql(['Jspreadsheet: Please clear your search before perform this action']);
        });

        it('marks the side of the header where a dragged column goes', async () => {
            const worksheet = await load({ worksheets: [{ minDimensions: [4, 1], columnResize: false }] });
            for (const header of worksheet.headers) {
                rect(header, 100, 20);
            }

            mouse(worksheet.headers[1], 'mousedown', { offsetY: 18 });
            expect(worksheet.pointer.dragging.column).to.equal(1);

            mouse(worksheet.headers[1], 'mousemove', { offsetX: 10 });
            expect(worksheet.pointer.dragging.destination).to.equal(1);

            mouse(worksheet.headers[3], 'mousemove', { offsetX: 10 });
            expect(worksheet.pointer.dragging.destination).to.equal(2);
            expect(worksheet.headers[3].classList.contains('dragging-left')).to.equal(true);

            mouse(worksheet.headers[0], 'mousemove', { offsetX: 10 });
            expect(worksheet.pointer.dragging.destination).to.equal(0);
            expect(worksheet.headers[3].classList.contains('dragging-left')).to.equal(false);

            mouse(worksheet.headers[0], 'mousemove', { offsetX: 90 });
            expect(worksheet.pointer.dragging.destination).to.equal(1);
            expect(worksheet.headers[0].classList.contains('dragging-right')).to.equal(true);

            mouse(worksheet.headers[0], 'mouseup');
            expect(worksheet.headers[0].classList.contains('dragging-right')).to.equal(false);
            expect(worksheet.pointer.dragging).to.equal(null);
        });

        it('does not move a column of a merged cell', async () => {
            const worksheet = await load({ worksheets: [{ minDimensions: [3, 2], columnResize: false, mergeCells: { A1: [2, 1] } }] });
            rect(worksheet.headers[0], 100, 20);

            mouse(worksheet.headers[0], 'mousedown', { offsetY: 18 });

            expect(worksheet.pointer.dragging).to.equal(null);
            expect(errors).to.eql(['Jspreadsheet: This column is part of a merged cell.']);
        });

        it('does not move a column over a merged cell', async () => {
            const worksheet = await load({ worksheets: [{ minDimensions: [3, 2], columnResize: false, mergeCells: { B1: [2, 1] } }] });
            rect(worksheet.headers[0], 100, 20);

            mouse(worksheet.headers[0], 'mousedown', { offsetY: 18 });
            mouse(worksheet.headers[2], 'mousemove');

            expect(errors).to.eql(['Jspreadsheet: This column is part of a merged cell.']);
        });
    });

    describe('cursor', () => {
        const move = (element, clientX, clientY) => mouse(element, 'mousemove', { buttons: 0, clientX, clientY });

        it('shows where a press resizes or moves a header or a row', async () => {
            const worksheet = await load({ worksheets: [{ minDimensions: [2, 2] }] });
            rect(worksheet.headers[0], 100, 20);
            rect(rowNumber(worksheet, 0), 50, 25);

            move(worksheet.headers[0], 98, 5);
            expect(worksheet.headers[0].style.cursor).to.equal('col-resize');

            move(worksheet.headers[0], 50, 18);
            expect(worksheet.headers[0].style.cursor).to.equal('move');

            move(rowNumber(worksheet, 0), 10, 23);
            expect(worksheet.headers[0].style.cursor).to.equal('');
            expect(rowNumber(worksheet, 0).style.cursor).to.equal('row-resize');

            move(rowNumber(worksheet, 0), 48, 10);
            expect(rowNumber(worksheet, 0).style.cursor).to.equal('move');

            move(rowNumber(worksheet, 0), 10, 10);
            expect(rowNumber(worksheet, 0).style.cursor).to.equal('');
        });
    });

    describe('selection', () => {
        it('extends the selection of columns over the headers', async () => {
            const worksheet = await load({ worksheets: [{ minDimensions: [4, 2] }] });

            mouse(worksheet.headers[0], 'mousedown', { offsetX: -10, offsetY: -10 });
            mouse(worksheet.headers[2], 'mouseover');
            mouse(worksheet.headers[2], 'mouseup');

            expect(worksheet.getSelectedColumns()).to.eql([0, 1, 2]);
        });

        it('extends the selection of rows over the row numbers and with Ctrl', async () => {
            const worksheet = await load({ worksheets: [{ minDimensions: [2, 4] }] });

            mouse(rowNumber(worksheet, 0), 'mousedown', { offsetX: -10, offsetY: -10 });
            mouse(rowNumber(worksheet, 2), 'mouseover');
            mouse(rowNumber(worksheet, 2), 'mouseup');

            expect(worksheet.getSelectedRows(true)).to.eql([0, 1, 2]);

            mouse(rowNumber(worksheet, 3), 'mousedown', { offsetX: -10, offsetY: -10, ctrlKey: true });
            mouse(rowNumber(worksheet, 3), 'mouseup');

            expect(worksheet.getSelectedRows(true)).to.eql([0, 1, 2, 3]);
        });

        it('extends the selection of cells over the cells and with Shift', async () => {
            const worksheet = await load({ worksheets: [{ minDimensions: [4, 4] }] });

            mouse(worksheet.records[0][0].element, 'mousedown');
            mouse(worksheet.records[1][1].element, 'mouseover');
            mouse(worksheet.records[1][1].element, 'mouseup');

            expect(worksheet.selection.range).to.eql([0, 0, 1, 1]);

            mouse(worksheet.records[2][3].element, 'mousedown', { shiftKey: true });
            mouse(worksheet.records[2][3].element, 'mouseup');

            expect(worksheet.selection.range).to.eql([0, 0, 3, 2]);
        });

        it('selects the cell of an element in the cell', async () => {
            const worksheet = await load({ worksheets: [{ data: [[true, false]], columns: [{ type: 'checkbox' }, { type: 'checkbox' }] }] });

            mouse(worksheet.records[0][1].element.children[0], 'mousedown');
            mouse(worksheet.records[0][1].element.children[0], 'mouseup');

            expect(worksheet.selection.range).to.eql([1, 0, 1, 0]);
        });

        it('selects the columns of a nested header', async () => {
            const worksheet = await load({
                worksheets: [{ minDimensions: [3, 2], nestedHeaders: [[{ title: 'First', colspan: 2 }, { title: 'Second' }]] }],
            });

            mouse(root.querySelector('.jss_nested td[data-column="0,1"]'), 'mousedown');
            mouse(document.body, 'mouseup');

            expect(worksheet.selection.range).to.eql([0, 0, 1, 1]);
        });

        it('closes the editor of another cell', async () => {
            const worksheet = await load({ worksheets: [{ data: [['a', 'b']] }] });

            worksheet.openEditor(worksheet.records[0][0].element);
            worksheet.records[0][0].element.children[0].value = 'x';

            mouse(worksheet.records[0][1].element, 'mousedown');
            mouse(worksheet.records[0][1].element, 'mouseup');

            expect(worksheet.editor.isOpen()).to.equal(false);
            expect(worksheet.getData()).to.eql([['x', 'b']]);
            expect(worksheet.selection.range).to.eql([1, 0, 1, 0]);
        });

        it('opens the pages of the pagination', async () => {
            const data = [];
            for (let i = 0; i < 30; i++) {
                data.push([i]);
            }
            const worksheet = await load({ worksheets: [{ data, pagination: 2 }] });

            const page = (text) => [...root.querySelectorAll('.jss_page')].find((element) => element.textContent === text);

            mouse(page('3'), 'mousedown');
            expect(worksheet.view.pageNumber).to.equal(2);

            mouse(page('>'), 'mousedown');
            expect(worksheet.view.pageNumber).to.equal(14);

            mouse(page('<'), 'mousedown');
            expect(worksheet.view.pageNumber).to.equal(0);
        });
    });

    describe('fill', () => {
        const corner = (worksheet) => worksheet.corner;

        it('fills the cells below with the numbers after the selection', async () => {
            const worksheet = await load({ worksheets: [{ data: [[1, 'a'], [''], [''], ['']] }] });

            worksheet.updateSelectionFromCoords(0, 0, 0, 0);
            mouse(corner(worksheet), 'mousedown');
            mouse(worksheet.records[2][1].element, 'mouseover');

            expect(worksheet.records[1][0].element.classList.contains('selection')).to.equal(true);
            expect(worksheet.records[2][0].element.classList.contains('selection-bottom')).to.equal(true);

            mouse(worksheet.records[2][1].element, 'mouseup');

            expect(worksheet.getColumnData(0)).to.eql([1, 2, 3, '']);
        });

        it('fills the cells above and to the left', async () => {
            const worksheet = await load({
                worksheets: [
                    {
                        data: [
                            ['', '', ''],
                            ['', '', ''],
                            ['', '', 5],
                        ],
                    },
                ],
            });

            worksheet.updateSelectionFromCoords(2, 2, 2, 2);
            mouse(corner(worksheet), 'mousedown');
            mouse(worksheet.records[0][2].element, 'mouseover');
            mouse(worksheet.records[0][2].element, 'mouseup');

            expect(worksheet.getColumnData(2)).to.eql([3, 4, 5]);

            worksheet.updateSelectionFromCoords(2, 2, 2, 2);
            mouse(corner(worksheet), 'mousedown');
            mouse(worksheet.records[2][0].element, 'mouseover');
            mouse(worksheet.records[2][0].element, 'mouseup');

            expect(worksheet.getRowData(2)).to.eql([5, 5, 5]);
        });

        it('moves the references of a formula and the dates of a calendar', async () => {
            const worksheet = await load({
                worksheets: [
                    {
                        data: [
                            ['=B1', '2024-01-31 00:00:00'],
                            ['', ''],
                        ],
                        columns: [{ type: 'text' }, { type: 'calendar' }],
                    },
                ],
            });

            worksheet.updateSelectionFromCoords(0, 0, 0, 0);
            mouse(corner(worksheet), 'mousedown');
            mouse(worksheet.records[1][0].element, 'mouseover');
            mouse(worksheet.records[1][0].element, 'mouseup');

            worksheet.updateSelectionFromCoords(1, 0, 1, 0);
            mouse(corner(worksheet), 'mousedown');
            mouse(worksheet.records[1][1].element, 'mouseover');
            mouse(worksheet.records[1][1].element, 'mouseup');

            expect(worksheet.getValueFromCoords(0, 1)).to.equal('=B2');
            expect(worksheet.getValueFromCoords(1, 1)).to.equal('2024-02-01 00:00:00');
        });

        it('fills no cell with a corner that did not move', async () => {
            const worksheet = await load({ worksheets: [{ data: [[1], ['']] }] });

            worksheet.updateSelectionFromCoords(0, 0, 0, 0);
            mouse(corner(worksheet), 'mousedown');
            mouse(corner(worksheet), 'mouseup');

            expect(worksheet.pointer.corner).to.equal(false);
            expect(worksheet.getColumnData(0)).to.eql([1, '']);
        });
    });
});
