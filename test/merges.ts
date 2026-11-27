import { expect } from 'chai';

import jspreadsheet, { type Jspreadsheet } from './support/jspreadsheet';

describe('Merge tests', () => {
    describe('Get merge', () => {
        it('Worksheet started with a merge', () => {
            const instance = jspreadsheet(root, {
                toolbar: true,
                worksheets: [
                    {
                        data: [
                            ['Mazda', 2001, 2000, '2006-01-01 12:00:00'],
                            ['Peugeot', 2010, 5000, '2005-01-01 13:00:00'],
                            ['Honda Fit', 2009, 3000, '2004-01-01 14:01:00'],
                            ['Honda CRV', 2010, 6000, '2003-01-01 23:30:00'],
                        ],
                        minDimensions: [10, 15],
                        columns: [
                            {
                                width: '300px',
                            },
                            {
                                width: '80px',
                            },
                            {
                                width: '100px',
                            },
                            {
                                width: '150px',
                            },
                        ],
                        mergeCells: {
                            C1: [1, 2],
                        },
                    },
                ],
            });

            expect(instance[0].getMerge('C1')).to.eql([1, 2]);
            expect(instance[0].getMerge('C2')).to.equal(null);

            expect(instance[0].getMerge()).to.eql({ C1: [1, 2] });
        });

        it('Worksheet started without merges', () => {
            const instance = jspreadsheet(root, {
                toolbar: true,
                worksheets: [
                    {
                        data: [
                            ['Mazda', 2001, 2000, '2006-01-01 12:00:00'],
                            ['Peugeot', 2010, 5000, '2005-01-01 13:00:00'],
                            ['Honda Fit', 2009, 3000, '2004-01-01 14:01:00'],
                            ['Honda CRV', 2010, 6000, '2003-01-01 23:30:00'],
                        ],
                        minDimensions: [10, 15],
                        columns: [
                            {
                                width: '300px',
                            },
                            {
                                width: '80px',
                            },
                            {
                                width: '100px',
                            },
                            {
                                width: '150px',
                            },
                        ],
                    },
                ],
            });

            expect(instance[0].getMerge('C1')).to.equal(null);

            expect(instance[0].getMerge()).to.eql({});
        });
    });

    it('Set merge', () => {
        const instance = jspreadsheet(root, {
            toolbar: true,
            worksheets: [
                {
                    data: [
                        ['Mazda', 2001, 2000, '2006-01-01 12:00:00'],
                        ['Peugeot', 2010, 5000, '2005-01-01 13:00:00'],
                        ['Honda Fit', 2009, 3000, '2004-01-01 14:01:00'],
                        ['Honda CRV', 2010, 6000, '2003-01-01 23:30:00'],
                    ],
                    minDimensions: [10, 15],
                    columns: [
                        {
                            width: '300px',
                        },
                        {
                            width: '80px',
                        },
                        {
                            width: '100px',
                        },
                        {
                            width: '150px',
                        },
                    ],
                },
            ],
        });

        instance[0].setMerge('A3', 2, 3);

        const table = root.querySelector('tbody')!;
        const rows = table.children;

        expect(rows[2].children[1].getAttribute('colspan')).to.equal('2');
        expect(rows[2].children[1].getAttribute('rowspan')).to.equal('3');
    });

    it('Remove merge', () => {
        const instance = jspreadsheet(root, {
            toolbar: true,
            worksheets: [
                {
                    data: [
                        ['Mazda', 2001, 2000, '2006-01-01 12:00:00'],
                        ['Peugeot', 2010, 5000, '2005-01-01 13:00:00'],
                        ['Honda Fit', 2009, 3000, '2004-01-01 14:01:00'],
                        ['Honda CRV', 2010, 6000, '2003-01-01 23:30:00'],
                    ],
                    minDimensions: [10, 15],
                    columns: [
                        {
                            width: '300px',
                        },
                        {
                            width: '80px',
                        },
                        {
                            width: '100px',
                        },
                        {
                            width: '150px',
                        },
                    ],
                    mergeCells: {
                        A1: [2, 2],
                        E5: [3, 2],
                    },
                },
            ],
        });

        const table = root.querySelector('tbody')!;
        const rows = table.children;

        instance[0].removeMerge('A1');

        expect(rows[0].children[1].getAttribute('colspan')).to.equal(null);
        expect(rows[0].children[1].getAttribute('rowspan')).to.equal(null);

        expect(rows[4].children[5].getAttribute('colspan')).to.equal('3');
        expect(rows[4].children[5].getAttribute('rowspan')).to.equal('2');
    });

    it('Remove all merge', () => {
        const instance = jspreadsheet(root, {
            toolbar: true,
            worksheets: [
                {
                    data: [
                        ['Mazda', 2001, 2000, '2006-01-01 12:00:00'],
                        ['Peugeot', 2010, 5000, '2005-01-01 13:00:00'],
                        ['Honda Fit', 2009, 3000, '2004-01-01 14:01:00'],
                        ['Honda CRV', 2010, 6000, '2003-01-01 23:30:00'],
                    ],
                    minDimensions: [10, 15],
                    columns: [
                        {
                            width: '300px',
                        },
                        {
                            width: '80px',
                        },
                        {
                            width: '100px',
                        },
                        {
                            width: '150px',
                        },
                    ],
                    mergeCells: {
                        A1: [2, 2],
                        E5: [3, 2],
                    },
                },
            ],
        });

        const table = root.querySelector('tbody')!;
        const rows = table.children;

        instance[0].destroyMerge();

        expect(rows[0].children[1].getAttribute('colspan')).to.equal(null);
        expect(rows[0].children[1].getAttribute('rowspan')).to.equal(null);

        expect(rows[4].children[5].getAttribute('colspan')).to.equal(null);
        expect(rows[4].children[5].getAttribute('rowspan')).to.equal(null);
    });

    it('setMerge history', () => {
        const instance = jspreadsheet(root, {
            toolbar: true,
            worksheets: [
                {
                    data: [
                        ['Mazda', 2001, 2000, '2006-01-01 12:00:00'],
                        ['Peugeot', 2010, 5000, '2005-01-01 13:00:00'],
                        ['Honda Fit', 2009, 3000, '2004-01-01 14:01:00'],
                        ['Honda CRV', 2010, 6000, '2003-01-01 23:30:00'],
                    ],
                    minDimensions: [10, 15],
                    columns: [
                        {
                            width: '300px',
                        },
                        {
                            width: '80px',
                        },
                        {
                            width: '100px',
                        },
                        {
                            width: '150px',
                        },
                    ],
                },
            ],
        });

        instance[0].setMerge('A3', 2, 3);

        const table = root.querySelector('tbody')!;
        const rows = table.children;

        expect(rows[2].children[1].getAttribute('colspan')).to.equal('2');
        expect(rows[2].children[1].getAttribute('rowspan')).to.equal('3');

        instance[0].undo();

        expect(rows[0].children[1].getAttribute('colspan')).to.equal(null);
        expect(rows[0].children[1].getAttribute('rowspan')).to.equal(null);

        instance[0].redo();

        expect(rows[2].children[1].getAttribute('colspan')).to.equal('2');
        expect(rows[2].children[1].getAttribute('rowspan')).to.equal('3');
    });

    it('removeMerge event', () => {
        let eventCalled = false;
        let eventCellName = null as string | null;
        let eventInstance = null as Jspreadsheet.WorksheetInstance | null;
        let eventBeforeMerges = null as Record<string, [number, number]> | null;

        const instance = jspreadsheet(root, {
            toolbar: true,
            worksheets: [
                {
                    data: [
                        ['Mazda', 2001, 2000, '2006-01-01 12:00:00'],
                        ['Peugeot', 2010, 5000, '2005-01-01 13:00:00'],
                        ['Honda Fit', 2009, 3000, '2004-01-01 14:01:00'],
                        ['Honda CRV', 2010, 6000, '2003-01-01 23:30:00'],
                    ],
                    minDimensions: [10, 15],
                    columns: [
                        {
                            width: '300px',
                        },
                        {
                            width: '80px',
                        },
                        {
                            width: '100px',
                        },
                        {
                            width: '150px',
                        },
                    ],
                    mergeCells: {
                        A1: [2, 2],
                        E5: [3, 2],
                    },
                },
            ],
            onunmerge: (worksheetInstance, cellName, beforeMerges) => {
                eventCalled = true;
                eventCellName = cellName;
                eventInstance = worksheetInstance;
                eventBeforeMerges = beforeMerges;
            },
        });

        instance[0].removeMerge('A1');
        instance[0].getMerge();

        expect(eventCalled).to.equal(true);
        expect(eventCellName).to.equal('A1');
        expect(eventInstance).to.equal(instance[0]);
        expect(eventBeforeMerges!['A1'][0]).to.eql(2);
        expect(eventBeforeMerges!['A1'][1]).to.eql(2);
        expect(Object.keys(eventBeforeMerges!).length).to.eql(1);

        expect(instance[0].getMerge('A1')).to.equal(null);
    });

    it('selects no cell and dispatches no onselection for the merged ranges of the options', async () => {
        let selections = 0;
        const worksheet = jspreadsheet(root, {
            worksheets: [
                {
                    data: [
                        ['a', 'b', 'c'],
                        ['d', 'e', 'f'],
                        ['g', 'h', 'i'],
                    ],
                    mergeCells: { A1: [2, 1], B2: [2, 2] },
                },
            ],
            onselection: () => {
                selections++;
            },
        })[0];
        await new Promise((resolve) => setTimeout(resolve));

        expect(worksheet.getMerge()).to.eql({ A1: [2, 1], B2: [2, 2] });
        expect(worksheet.getSelection()).to.equal(null);
        expect(selections).to.equal(0);

        // A merge through the API still selects the first cell of the range
        worksheet.removeMerge('A1');
        worksheet.setMerge('A1', 2, 1);
        expect(worksheet.getSelection()).to.eql([0, 0, 0, 0]);
    });

    it('setMerge without arguments merges the selected cells', () => {
        const instance = jspreadsheet(root, {
            worksheets: [
                {
                    minDimensions: [4, 4],
                },
            ],
        });

        instance[0].updateSelectionFromCoords(1, 1, 2, 2);
        instance[0].setMerge();

        expect(instance[0].getMerge('B2')).to.eql([2, 2]);
    });

    it('setMerge with one span finds a conflict with another merge', () => {
        const instance = jspreadsheet(root, {
            worksheets: [
                {
                    minDimensions: [3, 3],
                    mergeCells: { A2: [2, 2] },
                },
            ],
        });

        const alerts: string[] = [];
        // The library calls the global alert
        const alert = global.alert;
        global.alert = (message: string) => alerts.push(message);

        try {
            instance[0].setMerge('A1', undefined, 2);
        } finally {
            global.alert = alert;
        }

        expect(alerts).to.eql(['There is a conflict with another merged cell']);
        expect(instance[0].getMerge()).to.eql({ A2: [2, 2] });

        // One span without a conflict still merges
        instance[0].setMerge('B1', 2);
        expect(instance[0].getMerge('B1')).to.eql([2, 1]);
    });
    describe('Insert a row or a column inside a merged range', () => {
        const confirm = globalThis.confirm;

        afterEach(() => {
            globalThis.confirm = confirm;
        });

        const create = () =>
            jspreadsheet(root, {
                worksheets: [{ minDimensions: [4, 4], mergeCells: { B2: [2, 2] } }],
            })[0];

        it('Keeps the merge and the rows when the user cancels', () => {
            const worksheet = create();
            globalThis.confirm = () => false;

            worksheet.insertRow(1, 2, true);

            expect(worksheet.getMerge()).to.eql({ B2: [2, 2] });
            expect(worksheet.rows.length).to.equal(4);
        });

        it('Removes the merges and inserts the row when the user agrees', () => {
            const worksheet = create();
            globalThis.confirm = () => true;

            worksheet.insertRow(1, 2, true);

            expect(worksheet.getMerge()).to.eql({});
            expect(worksheet.rows.length).to.equal(5);
        });

        it('Removes the merges and inserts the column when the user agrees', () => {
            const worksheet = create();
            globalThis.confirm = () => true;

            worksheet.insertColumn(1, 2, true);

            expect(worksheet.getMerge()).to.eql({});
            expect(worksheet.headers.length).to.equal(5);
        });

        it('Does not ask when the row is outside the merged range', () => {
            const worksheet = create();
            globalThis.confirm = () => {
                throw new Error('confirm');
            };

            worksheet.insertRow(1, 3);

            expect(worksheet.getMerge()).to.eql({ B2: [2, 2] });
            expect(worksheet.rows.length).to.equal(5);
        });
    });

    describe('Insert and delete remove only the merged ranges that they cross', () => {
        const confirm = globalThis.confirm;
        let questions: (string | undefined)[];

        beforeEach(() => {
            questions = [];
            globalThis.confirm = (message) => {
                questions.push(message);
                return true;
            };
        });

        afterEach(() => {
            globalThis.confirm = confirm;
        });

        function create(mergeCells: Record<string, [number, number]>) {
            const data = Array.from({ length: 6 }, (row, y) => Array.from({ length: 8 }, (cell, x) => `${x}:${y}`));
            return jspreadsheet(root, { worksheets: [{ data, mergeCells }] })[0];
        }

        it('keeps the other ranges when insertRow splits one range', () => {
            const worksheet = create({ A1: [2, 1], B2: [2, 2], A5: [3, 1] });

            worksheet.insertRow(1, 2, true);

            expect(questions).to.have.length(1);
            expect(worksheet.getMerge()).to.eql({ A1: [2, 1], A6: [3, 1] });
            expect(worksheet.records[0][0].element.getAttribute('colspan')).to.equal('2');
            expect(worksheet.records[5][0].element.getAttribute('colspan')).to.equal('3');
            expect(worksheet.records[1][1].element.getAttribute('rowspan')).to.equal(null);
        });

        it('keeps the other ranges when insertColumn splits one range', () => {
            const worksheet = create({ A1: [1, 2], B3: [2, 2], E1: [1, 3] });

            worksheet.insertColumn(1, 2, true);

            expect(questions).to.have.length(1);
            expect(worksheet.getMerge()).to.eql({ A1: [1, 2], F1: [1, 3] });
            expect(worksheet.records[0][5].element.getAttribute('rowspan')).to.equal('3');
            expect(worksheet.records[2][1].element.getAttribute('colspan')).to.equal(null);
        });

        it('keeps the other ranges when deleteRow deletes a row of one range', () => {
            const worksheet = create({ A1: [2, 1], B3: [2, 2], A6: [2, 1] });

            worksheet.deleteRow(2, 1);

            expect(questions).to.have.length(1);
            expect(worksheet.getMerge()).to.eql({ A1: [2, 1], A5: [2, 1] });
            expect(worksheet.records[4][0].element.getAttribute('colspan')).to.equal('2');
            expect(worksheet.getValue('A3')).to.equal('0:3');
        });

        it('keeps a range in the row that stays when deleteRow names all rows', () => {
            const worksheet = create({ A6: [2, 1] });

            worksheet.deleteRow(0, 6);

            expect(questions).to.have.length(0);
            expect(worksheet.rows.length).to.equal(1);
            expect(worksheet.getMerge()).to.eql({ A1: [2, 1] });
        });

        it('keeps the other ranges when deleteColumn deletes a column of one range', () => {
            const worksheet = create({ A1: [1, 2], C2: [2, 2], H1: [1, 2] });

            worksheet.deleteColumn(3, 1);

            expect(questions).to.have.length(1);
            expect(worksheet.getMerge()).to.eql({ A1: [1, 2], G1: [1, 2] });
            expect(worksheet.records[0][6].element.getAttribute('rowspan')).to.equal('2');
        });

        it('keeps all ranges and rows when the user cancels a delete', () => {
            const worksheet = create({ A1: [2, 1], B3: [2, 2] });
            globalThis.confirm = () => false;

            worksheet.deleteRow(3, 1);

            expect(worksheet.getMerge()).to.eql({ A1: [2, 1], B3: [2, 2] });
            expect(worksheet.rows.length).to.equal(6);
        });

        it('logs nothing when the user cancels a delete of all rows', () => {
            const worksheet = create({ A1: [2, 1] });
            globalThis.confirm = () => false;
            const error = console.error;
            const errors: unknown[] = [];
            console.error = (message: unknown) => errors.push(message);
            try {
                expect(worksheet.deleteRow(0, 6)).to.equal(false);
            } finally {
                console.error = error;
            }

            expect(errors).to.eql([]);
            expect(worksheet.rows.length).to.equal(6);
            expect(worksheet.getMerge()).to.eql({ A1: [2, 1] });
        });

        it('removes a range of one row when deleteRow deletes its row (#214)', () => {
            const worksheet = create({ A1: [2, 1] });

            worksheet.deleteRow(0, 1);

            expect(questions).to.have.length(1);
            expect(worksheet.getMerge()).to.eql({});
            expect(worksheet.records[0][0].element.getAttribute('colspan')).to.equal(null);
        });

        it('removes a range when deleteRow deletes its last row', () => {
            const worksheet = create({ A1: [1, 2] });

            worksheet.deleteRow(1, 1);

            expect(questions).to.have.length(1);
            expect(worksheet.getMerge()).to.eql({});
            expect(worksheet.records[0][0].element.getAttribute('rowspan')).to.equal(null);
            expect(worksheet.getValue('A2')).to.equal('0:2');
        });

        it('keeps a range that moves onto the old name of another range (#195)', () => {
            const columns = create({ A1: [1, 2], B1: [1, 2] });
            columns.insertColumn(1, 0, true);
            expect(columns.getMerge()).to.eql({ B1: [1, 2], C1: [1, 2] });

            const rows = create({ A1: [2, 1], A2: [2, 1] });
            rows.insertRow(1, 0, true);
            expect(rows.getMerge()).to.eql({ A2: [2, 1], A3: [2, 1] });

            const sizes = create({ A1: [1, 2], C1: [2, 2] });
            sizes.insertColumn(2, 0, true);
            expect(sizes.getMerge()).to.eql({ C1: [1, 2], E1: [2, 2] });
            expect(questions).to.have.length(0);
        });
    });
});

describe('replaceMerges', () => {
    function build(events: string[]) {
        const data: string[][] = [];
        for (let j = 0; j < 4; j++) {
            data.push(['a', 'b', 'c', 'd'].map((letter) => `${letter}${j + 1}`));
        }
        return jspreadsheet(root, {
            worksheets: [{ data, mergeCells: { A1: [2, 1], C1: [2, 2] } }],
            onmerge: (instance, merges) => {
                events.push(`merge ${JSON.stringify(merges)}`);
            },
            onunmerge: (instance, cellName) => {
                events.push(`unmerge ${cellName}`);
            },
        })[0];
    }

    it('keeps the ranges that do not change, removes the others and adds the new ones', () => {
        const events: string[] = [];
        const worksheet = build(events);
        const first = worksheet.records[0][0].element;

        const refused = worksheet.replaceMerges({ A1: [2, 1], A3: [1, 2] });

        expect(refused).to.eql([]);
        expect(worksheet.getMerge()).to.eql({ A1: [2, 1], A3: [1, 2] });
        // The range that stays keeps its cell, and its covered cell stays hidden
        expect(worksheet.records[0][0].element).to.equal(first);
        expect(worksheet.records[0][1].element).to.equal(first);
        // The cells of the removed range C1:D2 show again, empty
        expect(worksheet.records[1][3].element.style.display).to.equal('');
        expect(worksheet.records[1][3].element.textContent).to.equal('');
        expect(worksheet.records[0][2].element.hasAttribute('colspan')).to.equal(false);
        // The new range A3:A4 hides A4 and clears its value
        expect(worksheet.records[2][0].element.getAttribute('rowspan')).to.equal('2');
        expect(worksheet.records[3][0].element).to.equal(worksheet.records[2][0].element);
        expect(worksheet.getValueFromCoords(0, 3)).to.equal('');
        expect(events).to.eql(['unmerge C1', 'merge {"A3":[1,2]}']);
    });

    it('gives the ranges that it cannot merge, and shows no dialog', () => {
        // The library calls the global alert
        const alerts: string[] = [];
        const alert = global.alert;
        global.alert = (text?: string) => {
            alerts.push(`${text}`);
        };
        try {
            const worksheet = build([]);

            const refused = worksheet.replaceMerges({ A1: [2, 1], C1: [2, 2], B1: [1, 2], D4: [2, 1], A3: [1, 1], a3: [1, 2], A4: [0, 1] });

            // The entries that are not ranges come first, then the ranges that do not fit
            expect(refused).to.eql(['A3', 'a3', 'A4', 'B1', 'D4']);
            expect(worksheet.getMerge()).to.eql({ A1: [2, 1], C1: [2, 2] });
            expect(alerts).to.eql([]);
        } finally {
            global.alert = alert;
        }
    });

    it('checks each entry before it changes a range', () => {
        const events: string[] = [];
        const worksheet = build(events);

        const entries = JSON.parse('{"A3":[2,1],"A0":[2,1],"A01":[2,1],"constructor":[2,1],"__proto__":[2,1],"B4":null,"C4":[2,"1"]}');
        const refused = worksheet.replaceMerges(entries);

        expect(refused).to.have.members(['A0', 'A01', 'constructor', '__proto__', 'B4', 'C4']);
        expect(worksheet.getMerge()).to.eql({ A3: [2, 1] });
        expect(events).to.eql(['unmerge A1', 'unmerge C1', 'merge {"A3":[2,1]}']);

        // @ts-expect-error -- a caller in JavaScript can give null
        expect(() => worksheet.replaceMerges(null)).to.throw(TypeError);
        expect(worksheet.getMerge()).to.eql({ A3: [2, 1] });
    });

    it('writes no history entry and keeps the selection', () => {
        const worksheet = build([]);
        worksheet.updateSelectionFromCoords(1, 3, 1, 3);
        const index = worksheet.history.index;

        worksheet.replaceMerges({ A3: [2, 1] });

        expect(worksheet.history.index).to.equal(index);
        expect(worksheet.getSelection()).to.eql([1, 3, 1, 3]);
    });

    it('selects the first cell of a new range that hides the selected cell, as setMerge does', () => {
        const worksheet = build([]);
        worksheet.updateSelectionFromCoords(1, 3, 1, 3);

        worksheet.replaceMerges({ A1: [2, 1], C1: [2, 2], A3: [2, 2] });

        const first = worksheet.records[2][0].element;
        expect(worksheet.getSelection()).to.eql([0, 2, 0, 2]);
        expect(first.tabIndex).to.equal(0);
        expect(first.getAttribute('aria-selected')).to.equal('true');
        expect(Array.from(worksheet.tbody.querySelectorAll('td[tabindex="0"]')).every((td) => (td as HTMLElement).style.display !== 'none')).to.equal(true);
    });

    it('selects the first cell of a new range also for a selection from the bottom right', () => {
        const worksheet = build([]);
        worksheet.updateSelectionFromCoords(3, 3, 0, 2);

        worksheet.replaceMerges({ A1: [2, 1], C1: [2, 2], C3: [2, 2] });

        expect(worksheet.selection.range!.map(Number)).to.eql([2, 2, 2, 2]);
        expect(worksheet.records[2][2].element.classList.contains('highlight-selected')).to.equal(true);
    });

    it('moves the tab stop to a visible cell when onbeforeselection cancels the selection', () => {
        const data: string[][] = [];
        for (let j = 0; j < 4; j++) {
            data.push(['a', 'b', 'c', 'd']);
        }
        let cancel = false;
        const worksheet = jspreadsheet(root, {
            worksheets: [{ data }],
            onbeforeselection: () => (cancel ? false : undefined),
        })[0];
        worksheet.updateSelectionFromCoords(3, 3, 3, 3);

        cancel = true;
        worksheet.replaceMerges({ C3: [2, 2] });

        const stops = Array.from(worksheet.tbody.querySelectorAll('td[tabindex="0"]')) as HTMLElement[];
        expect(stops.length).to.equal(1);
        expect(stops[0].style.display).to.not.equal('none');
    });

    it('selects again a cell that a removed range shows, as removeMerge does', () => {
        const worksheet = build([]);
        worksheet.updateSelectionFromCoords(3, 1, 3, 1);

        worksheet.replaceMerges({ A1: [2, 1] });

        const cell = worksheet.records[1][3].element;
        expect(cell.style.display).to.equal('');
        expect(cell.tabIndex).to.equal(0);
        expect(cell.getAttribute('aria-selected')).to.equal('true');
    });

    it('closes without a save an editor in a cell that a new range hides', () => {
        const worksheet = build([]);
        worksheet.openEditor(worksheet.records[2][1].element);

        worksheet.replaceMerges({ A1: [2, 1], C1: [2, 2], A3: [2, 1] });

        expect(worksheet.editor.isOpen()).to.equal(false);
        expect(worksheet.getMerge()).to.eql({ A1: [2, 1], C1: [2, 2], A3: [2, 1] });
    });

    it('gives the visible columns of a new range in a hidden column to its colspan', () => {
        const worksheet = build([]);
        worksheet.hideColumn(1);

        worksheet.replaceMerges({ A3: [3, 1] });
        expect(worksheet.records[2][0].element.getAttribute('colspan')).to.equal('2');

        worksheet.showColumn(1);
        expect(worksheet.records[2][0].element.getAttribute('colspan')).to.equal('3');
    });

    it('dispatches no event while ignoreEvents is set', () => {
        const events: string[] = [];
        const worksheet = build(events);

        worksheet.parent.ignoreEvents = true;
        try {
            worksheet.replaceMerges({ A3: [2, 1] });
        } finally {
            worksheet.parent.ignoreEvents = false;
        }

        expect(events).to.eql([]);
        expect(worksheet.getMerge()).to.eql({ A3: [2, 1] });
    });

    it('treats a range of the options that the build did not merge as not merged', () => {
        const alert = global.alert;
        global.alert = () => undefined;
        try {
            const worksheet = jspreadsheet(root, {
                worksheets: [{ minDimensions: [4, 4], mergeCells: { A1: [2, 2], B2: [2, 1] } }],
            })[0];

            expect(worksheet.replaceMerges({ A1: [2, 2] })).to.eql([]);
            expect(worksheet.getMerge()).to.eql({ A1: [2, 2] });
        } finally {
            global.alert = alert;
        }
    });
});

describe('removeMerge and the selection events', () => {
    it('lets onselection see the range while removeMerge selects it', () => {
        const seen: unknown[] = [];
        const worksheet = jspreadsheet(root, {
            worksheets: [{ minDimensions: [3, 3], mergeCells: { A1: [2, 2] } }],
            onselection: (instance) => {
                seen.push(instance.getMerge('A1'));
            },
        })[0];

        worksheet.removeMerge('A1');

        expect(seen).to.eql([[2, 2]]);
        expect(worksheet.getMerge()).to.eql({});
    });
});
