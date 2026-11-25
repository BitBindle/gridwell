const { expect } = require('chai');

const jspreadsheet = require('../dist/index.js');

describe('Use the rows method', () => {
    it('deleteRow and a row is removed', () => {
        const instance = jspreadsheet(root, {
            tabs: true,
            worksheets: [
                {
                    minDimensions: [7, 7],
                    data: [
                        [1, 2, 3],
                        [4, 5, 6],
                    ],
                    worksheetName: 'Countries',
                },
            ],
        });

        let table = root.querySelector('tbody');
        let rows = table.children;
        let firstRow = rows[0];

        // Check that first row has the value of [1, 2, 3]
        expect(firstRow.children[1].innerHTML).to.include(1);
        expect(firstRow.children[2].innerHTML).to.include(2);
        expect(firstRow.children[3].innerHTML).to.include(3);

        instance[0].deleteRow(0);

        table = root.querySelector('tbody');
        rows = table.children;
        firstRow = rows[0];

        // Check that the value of the first row now is [4, 5, 6] since the first one got removed
        expect(firstRow.children[1].innerHTML).to.include(4);
        expect(firstRow.children[2].innerHTML).to.include(5);
        expect(firstRow.children[3].innerHTML).to.include(6);
    });

    it('insertRow and a row is added', () => {
        const instance = jspreadsheet(root, {
            tabs: true,
            worksheets: [
                {
                    minDimensions: [7, 7],
                    data: [
                        [1, 2, 3],
                        [4, 5, 6],
                    ],
                    worksheetName: 'Countries',
                },
            ],
        });

        let table = root.querySelector('tbody');
        let rows = table.children;
        let firstRow = rows[0];

        // Check that first row has the value of [1, 2, 3]
        expect(firstRow.children[1].innerHTML).to.include(1);
        expect(firstRow.children[2].innerHTML).to.include(2);
        expect(firstRow.children[3].innerHTML).to.include(3);

        instance[0].insertRow([9, 9, 9], 0, 1);

        table = root.querySelector('tbody');
        rows = table.children;
        firstRow = rows[0];

        // Check that the value of the first row now is [9, 9, 9]
        expect(firstRow.children[1].innerHTML).to.include(9);
        expect(firstRow.children[2].innerHTML).to.include(9);
        expect(firstRow.children[3].innerHTML).to.include(9);
    });

    it('moveRow and the row is moved', () => {
        const instance = jspreadsheet(root, {
            tabs: true,
            worksheets: [
                {
                    minDimensions: [7, 7],
                    data: [
                        [1, 2, 3],
                        [4, 5, 6],
                    ],
                },
            ],
        });

        let table = root.querySelector('tbody');
        let rows = table.children;
        let firstRow = rows[0];
        let secondRow = rows[1];
        let A1 = firstRow.children[1];
        let A2 = secondRow.children[1];

        expect(A1.innerHTML).to.include(1);
        expect(A2.innerHTML).to.include(4);

        instance[0].moveRow(0, 1);

        table = root.querySelector('tbody');
        rows = table.children;
        firstRow = rows[0];
        secondRow = rows[1];
        A1 = firstRow.children[1];
        A2 = secondRow.children[1];

        expect(A1.innerHTML).to.include(4);
        expect(A2.innerHTML).to.include(1);
    });

    it('deleteRow history', () => {
        const instance = jspreadsheet(root, {
            tabs: true,
            worksheets: [
                {
                    minDimensions: [7, 7],
                    data: [
                        [1, 2, 3],
                        [4, 5, 6],
                    ],
                    worksheetName: 'Countries',
                },
            ],
        });

        let table = root.querySelector('tbody');
        let rows = table.children;

        expect(rows[0].children[1].innerHTML).to.include(1);
        expect(rows[0].children[2].innerHTML).to.include(2);
        expect(rows[0].children[3].innerHTML).to.include(3);

        instance[0].deleteRow(0);

        expect(rows[0].children[1].innerHTML).to.include(4);
        expect(rows[0].children[2].innerHTML).to.include(5);
        expect(rows[0].children[3].innerHTML).to.include(6);

        instance[0].undo();

        expect(rows[0].children[1].innerHTML).to.include(1);
        expect(rows[0].children[2].innerHTML).to.include(2);
        expect(rows[0].children[3].innerHTML).to.include(3);

        instance[0].redo();

        expect(rows[0].children[1].innerHTML).to.include(4);
        expect(rows[0].children[2].innerHTML).to.include(5);
        expect(rows[0].children[3].innerHTML).to.include(6);
    });

    it('insertRow history', () => {
        const instance = jspreadsheet(root, {
            tabs: true,
            worksheets: [
                {
                    minDimensions: [7, 7],
                    data: [
                        [1, 2, 3],
                        [4, 5, 6],
                    ],
                    worksheetName: 'Countries',
                },
            ],
        });

        let table = root.querySelector('tbody');
        let rows = table.children;

        expect(rows[0].children[1].innerHTML).to.include(1);
        expect(rows[0].children[2].innerHTML).to.include(2);
        expect(rows[0].children[3].innerHTML).to.include(3);

        instance[0].insertRow([9, 9, 9], 0, 1);

        expect(rows[0].children[1].innerHTML).to.include(9);
        expect(rows[0].children[2].innerHTML).to.include(9);
        expect(rows[0].children[3].innerHTML).to.include(9);

        instance[0].undo();

        expect(rows[0].children[1].innerHTML).to.include(1);
        expect(rows[0].children[2].innerHTML).to.include(2);
        expect(rows[0].children[3].innerHTML).to.include(3);

        instance[0].redo();

        expect(rows[0].children[1].innerHTML).to.include(9);
        expect(rows[0].children[2].innerHTML).to.include(9);
        expect(rows[0].children[3].innerHTML).to.include(9);
    });

    it('moveRow history', () => {
        const instance = jspreadsheet(root, {
            tabs: true,
            worksheets: [
                {
                    minDimensions: [7, 7],
                    data: [
                        [1, 2, 3],
                        [4, 5, 6],
                    ],
                },
            ],
        });

        let table = root.querySelector('tbody');

        expect(table.children[0].children[1].innerHTML).to.include(1);
        expect(table.children[1].children[1].innerHTML).to.include(4);

        instance[0].moveRow(0, 1);

        expect(table.children[0].children[1].innerHTML).to.include(4);
        expect(table.children[1].children[1].innerHTML).to.include(1);

        instance[0].undo();

        expect(table.children[0].children[1].innerHTML).to.include(1);
        expect(table.children[1].children[1].innerHTML).to.include(4);

        instance[0].redo();

        expect(table.children[0].children[1].innerHTML).to.include(4);
        expect(table.children[1].children[1].innerHTML).to.include(1);
    });
});

describe('Drag a row', () => {
    // The spreadsheet adds its event handlers after a promise
    const load = (options) =>
        new Promise((resolve) => {
            const instance = jspreadsheet(root, { ...options, onload: () => resolve(instance[0]) });
        });

    // jsdom has no layout, so offsetX and offsetY are set: a press at the right edge of a row number starts a drag
    const mouse = (element, type) => {
        const event = new window.MouseEvent(type, { bubbles: true, buttons: 1 });
        Object.defineProperty(event, 'offsetX', { value: 0 });
        Object.defineProperty(event, 'offsetY', { value: 0 });
        element.dispatchEvent(event);
    };

    it('ends the drag of the only row without an error', async () => {
        // Without row resize, the press at the edge is for a drag
        const worksheet = await load({ worksheets: [{ data: [['a']], rowResize: false }] });

        const errors = [];
        const onerror = (event) => {
            errors.push(event.message);
            // Keep jsdom from reporting the error to the console
            event.preventDefault();
        };
        window.addEventListener('error', onerror);

        const rowNumber = worksheet.rows[0].element.children[0];
        mouse(rowNumber, 'mousedown');
        expect(worksheet.dragging).to.not.equal(null);
        mouse(rowNumber, 'mouseup');

        window.removeEventListener('error', onerror);

        expect(errors).to.eql([]);
        expect(worksheet.dragging).to.equal(null);
        expect(worksheet.getData()).to.eql([['a']]);
    });
});
