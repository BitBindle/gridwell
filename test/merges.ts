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
});
