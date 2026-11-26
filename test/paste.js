const { expect } = require('chai');

const jspreadsheet = require('../dist/index.js');

global.document.execCommand = function execCommandMock() {};

const fixtureData = () => [
    ['Mazda', 2001, 2000, 1],
    ['Peugeot', 2010, 5000, '=B2+C2'],
    ['Honda Fit', 2009, 3000, '=B3+C3'],
    ['Honda CRV', 2010, 6000, '=B4+C4'],
];

describe('Paste', () => {
    it('no expand', () => {
        let sheet = jspreadsheet(root, {
            worksheets: [
                {
                    data: fixtureData(),
                },
            ],
        })[0];

        const pasteText = '0-0\t0-1\t0-2\t0-3\n1-0\t1-1\t1-2\t1-3\n2-0\t2-1\t2-2\t2-3\n3-0\t3-1\t3-2\t3-3';
        sheet.updateSelectionFromCoords(0, 0, 0, 0);
        sheet.paste(sheet.selection.range[0], sheet.selection.range[1], pasteText);

        expect(sheet.getData()).to.eql([
            ['0-0', '0-1', '0-2', '0-3'],
            ['1-0', '1-1', '1-2', '1-3'],
            ['2-0', '2-1', '2-2', '2-3'],
            ['3-0', '3-1', '3-2', '3-3'],
        ]);
    });

    it('expand', () => {
        let sheet = jspreadsheet(root, {
            worksheets: [
                {
                    data: fixtureData(),
                },
            ],
        })[0];

        const pasteText = '0-0\t0-1\t0-2\t0-3\n1-0\t1-1\t1-2\t1-3\n2-0\t2-1\t2-2\t2-3\n3-0\t3-1\t3-2\t3-3';
        sheet.updateSelectionFromCoords(3, 3, 3, 3);
        sheet.paste(sheet.selection.range[0], sheet.selection.range[1], pasteText);

        expect(sheet.getData()).to.eql([
            ['Mazda', 2001, 2000, 1, '', '', ''],
            ['Peugeot', 2010, 5000, '=B2+C2', '', '', ''],
            ['Honda Fit', 2009, 3000, '=B3+C3', '', '', ''],
            ['Honda CRV', 2010, 6000, '0-0', '0-1', '0-2', '0-3'],
            ['', '', '', '1-0', '1-1', '1-2', '1-3'],
            ['', '', '', '2-0', '2-1', '2-2', '2-3'],
            ['', '', '', '3-0', '3-1', '3-2', '3-3'],
        ]);
    });

    it('repeat horizontal', () => {
        let sheet = jspreadsheet(root, {
            worksheets: [
                {
                    data: fixtureData(),
                },
            ],
        })[0];

        const pasteText = '0-0\t0-1';
        sheet.updateSelectionFromCoords(0, 0, 4, 0);
        sheet.paste(sheet.selection.range[0], sheet.selection.range[1], pasteText);

        expect(sheet.getData()).to.eql([
            ['0-0', '0-1', '0-0', '0-1'],
            ['Peugeot', 2010, 5000, '=B2+C2'],
            ['Honda Fit', 2009, 3000, '=B3+C3'],
            ['Honda CRV', 2010, 6000, '=B4+C4'],
        ]);
    });

    it('repeat vertical', () => {
        let sheet = jspreadsheet(root, {
            worksheets: [
                {
                    data: fixtureData(),
                },
            ],
        })[0];

        const pasteText = '0-0\n1-0';
        sheet.updateSelectionFromCoords(0, 0, 0, 4);
        sheet.paste(sheet.selection.range[0], sheet.selection.range[1], pasteText);

        expect(sheet.getData()).to.eql([
            ['0-0', 2001, 2000, 1],
            ['1-0', 2010, 5000, '=B2+C2'],
            ['0-0', 2009, 3000, '=B3+C3'],
            ['1-0', 2010, 6000, '=B4+C4'],
        ]);
    });

    it('repeat rectangle', () => {
        let sheet = jspreadsheet(root, {
            worksheets: [
                {
                    data: fixtureData(),
                },
            ],
        })[0];

        const pasteText = '0-0\t0-1\n1-0\t1-1';
        sheet.updateSelectionFromCoords(1, 0, 1, 3);
        sheet.paste(sheet.selection.range[0], sheet.selection.range[1], pasteText);

        expect(sheet.getData()).to.eql([
            ['Mazda', '0-0', '0-1', 1],
            ['Peugeot', '1-0', '1-1', '=B2+C2'],
            ['Honda Fit', '0-0', '0-1', '=B3+C3'],
            ['Honda CRV', '1-0', '1-1', '=B4+C4'],
        ]);
    });

    it('skip hidden column', () => {
        let sheet = jspreadsheet(root, {
            worksheets: [
                {
                    columns: [
                        { type: 'text' },
                        { type: 'text' },
                        { type: 'hidden' }, // paste is skipped.
                        { type: 'text' },
                    ],
                    data: fixtureData(),
                },
            ],
        })[0];

        const pasteText = '0-0\t0-1\n1-0\t1-1';
        sheet.updateSelectionFromCoords(1, 0, 1, 0);
        sheet.paste(sheet.selection.range[0], sheet.selection.range[1], pasteText);

        expect(sheet.getData()).to.eql([
            ['Mazda', '0-0', 2000, '0-1'],
            ['Peugeot', '1-0', 5000, '1-1'],
            ['Honda Fit', 2009, 3000, '=B3+C3'],
            ['Honda CRV', 2010, 6000, '=B4+C4'],
        ]);
    });

    it('skip hidden row', () => {
        let sheet = jspreadsheet(root, {
            worksheets: [
                {
                    data: fixtureData(),
                },
            ],
        })[0];

        const pasteText = '0-0\t0-1\n1-0\t1-1';
        sheet.hideRow(1);
        sheet.updateSelectionFromCoords(1, 0, 1, 0);
        sheet.paste(sheet.selection.range[0], sheet.selection.range[1], pasteText);

        expect(sheet.getData()).to.eql([
            ['Mazda', '0-0', '0-1', 1],
            ['Peugeot', 2010, 5000, '=B2+C2'],
            ['Honda Fit', '1-0', '1-1', '=B3+C3'],
            ['Honda CRV', 2010, 6000, '=B4+C4'],
        ]);
    });

    it('see https://github.com/jspreadsheet/ce/pull/1717#issuecomment-2576060698', () => {
        let sheet = jspreadsheet(root, {
            worksheets: [
                {
                    minDimensions: [4, 4],
                    data: [
                        [1, 2],
                        [3, 4],
                    ],
                },
            ],
        })[0];

        sheet.updateSelectionFromCoords(0, 0, 1, 1);
        sheet.copy();
        sheet.hideRow(0);
        sheet.hideColumn(0);
        sheet.updateSelectionFromCoords(2, 2, 2, 2);
        sheet.paste(sheet.selection.range[0], sheet.selection.range[1], sheet.clipboard.data);

        expect(sheet.getData()).to.eql([
            [1, 2, '', ''],
            [3, 4, '', ''],
            ['', '', '1', '2'],
            ['', '', '3', '4'],
        ]);
    });

    it('see https://github.com/jspreadsheet/ce/pull/1717#issuecomment-2580087207', () => {
        let sheet = jspreadsheet(root, {
            worksheets: [
                {
                    minDimensions: [10, 4],
                    data: [
                        [1, 2, 3],
                        [4, 5, 6],
                    ],
                },
            ],
        })[0];

        sheet.updateSelectionFromCoords(0, 0, 2, 1);
        sheet.copy();
        sheet.hideColumn(8);
        sheet.hideColumn(9);
        sheet.hideRow(3);
        sheet.updateSelectionFromCoords(7, 2, 7, 2);
        expect(sheet.getData()).to.eql([
            [1, 2, 3, '', '', '', '', '', '', ''],
            [4, 5, 6, '', '', '', '', '', '', ''],
            ['', '', '', '', '', '', '', '', '', ''],
            ['', '', '', '', '', '', '', '', '', ''],
        ]);

        sheet.paste(7, 2, sheet.clipboard.data);

        expect(sheet.getData()).to.eql([
            [1, 2, 3, '', '', '', '', '', '', '', '', ''],
            [4, 5, 6, '', '', '', '', '', '', '', '', ''],
            ['', '', '', '', '', '', '', '1', '', '', '2', '3'],
            ['', '', '', '', '', '', '', '', '', '', '', ''],
            ['', '', '', '', '', '', '', '4', '', '', '5', '6'],
        ]);
    });

    it('large data paste', () => {
        let count = {};
        let sheet = jspreadsheet(root, {
            worksheets: [
                {
                    data: fixtureData(),
                },
            ],
            onevent: (event) => {
                count[event] = (count[event] ?? 0) + 1;
            },
        })[0];

        const pasteText = new Array(1000)
            .fill(0)
            .map((v, i) =>
                new Array(20)
                    .fill(0)
                    .map((v2, i2) => `${i}-${i2}`)
                    .join('\t')
            )
            .join('\n');
        sheet.updateSelectionFromCoords(3, 3, 3, 3);
        sheet.paste(sheet.selection.range[0], sheet.selection.range[1], pasteText);
        expect(count.onbeforechange).to.eql(20000);
        expect(count.onbeforeinsertcolumn).to.eql(1);
        expect(count.onbeforeinsertrow).to.eql(1);
        expect(count.onchange).to.eql(20000);
        expect(count.oninsertcolumn).to.eql(1);
        expect(count.oninsertrow).to.eql(1);
        expect(count.onselection).to.eql(3);
    }).timeout(10 * 1000);

    it('expand with hidden cells', () => {
        let count = {};
        let sheet = jspreadsheet(root, {
            worksheets: [
                {
                    columns: [
                        { type: 'text' },
                        { type: 'text' },
                        { type: 'hidden' }, // paste is skipped.
                        { type: 'text' },
                    ],
                    data: fixtureData(),
                },
            ],
            onevent: (event) => {
                count[event] = (count[event] ?? 0) + 1;
            },
        })[0];

        const pasteText = '0-0\t0-1\t0-2\t0-3\n1-0\t1-1\t1-2\t1-3\n2-0\t2-1\t2-2\t2-3\n3-0\t3-1\t3-2\t3-3';
        sheet.hideRow(2);
        sheet.updateSelectionFromCoords(1, 1, 1, 1);
        sheet.paste(sheet.selection.range[0], sheet.selection.range[1], pasteText);
        expect(sheet.getData()).to.eql([
            ['Mazda', 2001, 2000, 1, '', ''],
            ['Peugeot', '0-0', 5000, '0-1', '0-2', '0-3'],
            ['Honda Fit', 2009, 3000, '=B3+C3', '', ''],
            ['Honda CRV', '1-0', 6000, '1-1', '1-2', '1-3'],
            ['', '2-0', '', '2-1', '2-2', '2-3'],
            ['', '3-0', '', '3-1', '3-2', '3-3'],
        ]);
        expect(count.onbeforeinsertcolumn).to.eql(1);
        expect(count.onbeforeinsertrow).to.eql(1);
    }).timeout(10 * 1000);

    it('copy and paste with style', () => {
        let sheet = jspreadsheet(root, {
            worksheets: [
                {
                    data: fixtureData(),
                },
            ],
        })[0];

        sheet.setStyle('A1', 'color', 'red');
        sheet.updateSelectionFromCoords(0, 0, 1, 1);
        sheet.copy();
        sheet.paste(2, 2, sheet.clipboard.data);
        expect(sheet.getData()).to.eql([
            ['Mazda', 2001, 2000, 1],
            ['Peugeot', 2010, 5000, '=B2+C2'],
            ['Honda Fit', 2009, 'Mazda', '2001'],
            ['Honda CRV', 2010, 'Peugeot', '2010'],
        ]);
        expect(sheet.getStyle('A1', 'color')).to.eql('red');
        expect(sheet.getStyle('C3', 'color')).to.eql('red');
    });

    it('copy and repeat paste with style', () => {
        global.document.execCommand = function execCommandMock() {};

        let sheet = jspreadsheet(root, {
            worksheets: [
                {
                    data: fixtureData(),
                },
            ],
        })[0];

        sheet.setStyle('A1', 'color', 'red');
        sheet.updateSelectionFromCoords(0, 0, 1, 0);
        sheet.copy();
        sheet.updateSelectionFromCoords(0, 2, 4, 4);
        sheet.paste(0, 2, sheet.clipboard.data);
        expect(sheet.getData()).to.eql([
            ['Mazda', 2001, 2000, 1],
            ['Peugeot', 2010, 5000, '=B2+C2'],
            ['Mazda', '2001', 'Mazda', '2001'],
            ['Mazda', '2001', 'Mazda', '2001'],
        ]);
        //
        expect(sheet.getStyle('A1', 'color')).to.eql('red');
        //
        expect(sheet.getStyle('A3', 'color')).to.eql('red');
        expect(sheet.getStyle('B3', 'color')).to.eql('');
        expect(sheet.getStyle('C3', 'color')).to.eql('red');
        expect(sheet.getStyle('D3', 'color')).to.eql('');
        //
        expect(sheet.getStyle('A4', 'color')).to.eql('red');
        expect(sheet.getStyle('B4', 'color')).to.eql('');
        expect(sheet.getStyle('C4', 'color')).to.eql('red');
        expect(sheet.getStyle('D4', 'color')).to.eql('');
    });

    it('copy and paste to another sheet', async () => {
        global.document.execCommand = function execCommandMock() {};

        let isLoaded = false;
        let sheets = jspreadsheet(root, {
            tabs: true,
            worksheets: [
                {
                    data: fixtureData(),
                    worksheetName: 'Sheet1',
                },
                {
                    data: fixtureData(),
                    worksheetName: 'Sheet2',
                },
            ],
            onload: () => {
                isLoaded = true;
            },
        });

        const awaitLoop = (resolve) => {
            setTimeout(() => {
                if (isLoaded) {
                    resolve();
                } else {
                    resolve(awaitLoop);
                }
            }, 100);
        };
        // NOTE: jpreadsheet constructor is acutally async. So it waits for load events in await.
        await new Promise(awaitLoop);

        const from = sheets[0];
        from.setStyle('A1', 'color', 'red');
        from.updateSelectionFromCoords(0, 0, 1, 0);
        from.copy();
        const to = sheets[1];
        to.updateSelectionFromCoords(0, 2, 4, 4);
        to.paste(0, 2, from.clipboard.data);
        expect(to.getData()).to.eql([
            ['Mazda', 2001, 2000, 1],
            ['Peugeot', 2010, 5000, '=B2+C2'],
            ['Mazda', '2001', 'Mazda', '2001'],
            ['Mazda', '2001', 'Mazda', '2001'],
        ]);
    });

    it('fix - u0000 is pasted, when the last cell ends in a tab', () => {
        let sheet = jspreadsheet(root, {
            worksheets: [
                {
                    data: fixtureData(),
                },
            ],
        })[0];

        const pasteText = '0-0\t0-1\t0-2\t0-3\n1-0\t1-1\t1-2\t1-3\n2-0\t2-1\t2-2\t2-3\n3-0\t3-1\t3-2\t';
        sheet.updateSelectionFromCoords(0, 0, 0, 0);
        sheet.paste(sheet.selection.range[0], sheet.selection.range[1], pasteText);

        expect(sheet.getData()).to.eql([
            ['0-0', '0-1', '0-2', '0-3'],
            ['1-0', '1-1', '1-2', '1-3'],
            ['2-0', '2-1', '2-2', '2-3'],
            ['3-0', '3-1', '3-2', ''],
        ]);
    });

    it('paste lacked columns data', () => {
        let sheet = jspreadsheet(root, {
            worksheets: [
                {
                    data: fixtureData(),
                },
            ],
        })[0];

        const pasteText = '0-0\t\n' + '1-0\t1-1\t1-2\t1-3\n' + '2-0\n' + '3-0\t3-1\t3-2\t';
        sheet.updateSelectionFromCoords(0, 0, 0, 0);
        sheet.paste(sheet.selection.range[0], sheet.selection.range[1], pasteText);

        expect(sheet.getData()).to.eql([
            ['0-0', '', '', ''],
            ['1-0', '1-1', '1-2', '1-3'],
            ['2-0', '', '', ''],
            ['3-0', '3-1', '3-2', ''],
        ]);
    });

    describe('onbeforepaste', () => {
        const pasteWith = function (onbeforepaste) {
            const sheet = jspreadsheet(root, {
                worksheets: [{ minDimensions: [3, 3] }],
                onbeforepaste,
            })[0];
            sheet.updateSelectionFromCoords(0, 0, 0, 0);
            sheet.paste(0, 0, 'a\tb');
            return sheet;
        };

        it('pastes the text that it returns', () => {
            const sheet = pasteWith(() => 'x\ty\nz');

            expect(sheet.getData()).to.eql([
                ['x', 'y', ''],
                ['z', '', ''],
                ['', '', ''],
            ]);
        });

        it('pastes the data when it returns true', () => {
            const sheet = pasteWith(() => true);

            expect(sheet.getData()).to.eql([
                ['a', 'b', ''],
                ['', '', ''],
                ['', '', ''],
            ]);
        });

        it('pastes the rows of values that it returns', () => {
            const sheet = pasteWith(() => [['r']]);

            expect(sheet.getValue('A1')).to.equal('r');
            expect(sheet.getValue('B1')).to.equal('');
        });

        it('pastes the values of the { value } objects that it returns (issue #54)', () => {
            const sheet = pasteWith((instance, cells) => cells.map((row) => row.map((cell) => ({ value: cell.value.toUpperCase() }))));

            expect(sheet.getData()).to.eql([
                ['A', 'B', ''],
                ['', '', ''],
                ['', '', ''],
            ]);
        });

        it('cancels the paste when it returns false', () => {
            const sheet = pasteWith(() => false);

            expect(sheet.getData()).to.eql([
                ['', '', ''],
                ['', '', ''],
                ['', '', ''],
            ]);
        });
    });
});

describe('Copy and paste without a selection', () => {
    it('copies nothing and keeps the clipboard', () => {
        const events = [];
        const commands = [];
        const execCommand = document.execCommand;
        document.execCommand = (command) => commands.push(command);

        try {
            const worksheet = jspreadsheet(root, {
                worksheets: [
                    {
                        data: [
                            ['a', 'b'],
                            ['c', 'd'],
                        ],
                    },
                ],
                oncopy: () => events.push('oncopy'),
            })[0];

            worksheet.copy();
            worksheet.copy(true);

            expect(events).to.deep.equal([]);
            expect(commands).to.deep.equal([]);
            expect(worksheet.getData()).to.deep.equal([
                ['a', 'b'],
                ['c', 'd'],
            ]);
        } finally {
            document.execCommand = execCommand;
        }
    });

    it('pastes at the target cell', () => {
        const worksheet = jspreadsheet(root, { worksheets: [{ minDimensions: [3, 3] }] })[0];

        worksheet.paste(1, 1, 'x\ty\nz');

        expect(worksheet.getData()).to.deep.equal([
            ['', '', ''],
            ['', 'x', 'y'],
            ['', 'z', ''],
        ]);
        expect(worksheet.getSelection()).to.deep.equal([1, 1, 2, 2]);
    });
});

describe('Cut', () => {
    it('keeps the cells when oncopy cancels the cut', () => {
        const execCommand = document.execCommand;
        document.execCommand = () => {};

        try {
            const worksheet = jspreadsheet(root, {
                worksheets: [{ data: [['a', 'b']] }],
                oncopy: () => false,
            })[0];
            worksheet.updateSelectionFromCoords(0, 0, 1, 0);

            worksheet.copy(true);

            expect(worksheet.getData()).to.deep.equal([['a', 'b']]);
        } finally {
            document.execCommand = execCommand;
        }
    });

    it('clears the cells after a cut', () => {
        const execCommand = document.execCommand;
        document.execCommand = () => {};

        try {
            const worksheet = jspreadsheet(root, { worksheets: [{ data: [['a', 'b']] }] })[0];
            worksheet.updateSelectionFromCoords(0, 0, 0, 0);

            worksheet.copy(true);

            expect(worksheet.getData()).to.deep.equal([['', 'b']]);
            expect(worksheet.clipboard.data).to.equal('a');
        } finally {
            document.execCommand = execCommand;
        }
    });
});

describe('The styles of the copied cells', () => {
    const create = (options) =>
        jspreadsheet(root, {
            worksheets: [
                {
                    data: [
                        ['a', 'b'],
                        ['', ''],
                    ],
                    style: { A1: 'color:red', B1: 'color:blue' },
                },
            ],
            ...options,
        })[0];

    it('are kept when oncopy cancels a copy', () => {
        let cancel = false;
        const worksheet = create({ oncopy: () => (cancel ? false : undefined) });

        worksheet.updateSelectionFromCoords(0, 0, 0, 0);
        worksheet.copy();
        const copied = worksheet.clipboard.data;

        cancel = true;
        worksheet.updateSelectionFromCoords(1, 0, 1, 0);
        worksheet.copy();

        worksheet.updateSelectionFromCoords(0, 1, 0, 1);
        worksheet.paste(0, 1, copied);

        expect(worksheet.getValueFromCoords(0, 1)).to.equal('a');
        expect(worksheet.getStyle('A2', 'color')).to.equal('red');
    });

    it('are kept when the data is downloaded', () => {
        const createObjectURL = URL.createObjectURL;
        URL.createObjectURL = () => 'blob:test';

        try {
            const worksheet = create();

            worksheet.updateSelectionFromCoords(0, 0, 0, 0);
            worksheet.copy();
            const copied = worksheet.clipboard.data;

            worksheet.download();

            worksheet.updateSelectionFromCoords(0, 1, 0, 1);
            worksheet.paste(0, 1, copied);

            expect(worksheet.getValueFromCoords(0, 1)).to.equal('a');
            expect(worksheet.getStyle('A2', 'color')).to.equal('red');
        } finally {
            URL.createObjectURL = createObjectURL;
        }
    });
});

describe('The clipboard', () => {
    const create = () =>
        jspreadsheet(root, {
            worksheets: [
                {
                    data: [
                        ['a', 'b'],
                        ['c', 'd'],
                    ],
                    style: { A1: 'color:red' },
                },
            ],
        })[0];

    it('is empty before the first copy', () => {
        const worksheet = create();

        expect(worksheet.clipboard.data).to.equal('');
        expect(worksheet.clipboard.hash).to.equal(null);
        expect(worksheet.clipboard.style).to.deep.equal([]);
        expect(worksheet.clipboard.matches('')).to.equal(false);
    });

    it('keeps the text and the styles of a copy', () => {
        const commands = [];
        const execCommand = document.execCommand;
        document.execCommand = (command) => commands.push(command);

        try {
            const worksheet = create();

            worksheet.updateSelectionFromCoords(0, 0, 1, 0);
            worksheet.copy();

            expect(commands).to.deep.equal(['copy']);
            expect(worksheet.clipboard.textarea.value).to.equal('a\tb');
            expect(worksheet.clipboard.data).to.equal('a\tb');
            expect(worksheet.clipboard.style.length).to.equal(2);
            expect(worksheet.clipboard.style[0]).to.include('color: red');
            expect(worksheet.clipboard.matches('a\tb')).to.equal(true);
            expect(worksheet.clipboard.matches('a\tc')).to.equal(false);
        } finally {
            document.execCommand = execCommand;
        }
    });

    it('draws the border around the copied cells until the next paste', () => {
        const worksheet = create();
        const cell = (x, y) => worksheet.records[y][x].element;

        worksheet.updateSelectionFromCoords(0, 0, 1, 0);
        worksheet.copy();

        expect(cell(0, 0).classList.contains('copying')).to.equal(true);
        expect(cell(0, 0).classList.contains('copying-left')).to.equal(true);
        expect(cell(1, 0).classList.contains('copying-right')).to.equal(true);
        expect(cell(0, 1).classList.contains('copying')).to.equal(false);

        worksheet.updateSelectionFromCoords(0, 1, 0, 1);
        worksheet.paste(0, 1, worksheet.clipboard.data);

        expect(root.querySelectorAll('.copying').length).to.equal(0);
        expect(worksheet.getStyle('A2', 'color')).to.equal('red');
    });

    it('is the textarea that the worksheet labels', () => {
        const worksheet = create();

        expect(worksheet.content.contains(worksheet.clipboard.textarea)).to.equal(true);
        // jsdom does not reflect ariaHidden to the attribute
        expect(worksheet.clipboard.textarea.ariaHidden).to.equal('true');
        expect(worksheet.clipboard.textarea.tabIndex).to.equal(-1);
    });

    it('is not shared between worksheets', async () => {
        const worksheets = jspreadsheet(root, {
            worksheets: [{ data: [['A']] }, { data: [['B']] }],
        });

        // The worksheets after the first are built asynchronously
        await new Promise((resolve) => setTimeout(resolve));

        worksheets[0].updateSelectionFromCoords(0, 0, 0, 0);
        worksheets[0].copy();

        expect(worksheets[0].clipboard.data).to.equal('A');
        expect(worksheets[1].clipboard.data).to.equal('');
        expect(worksheets[0].clipboard).to.not.equal(worksheets[1].clipboard);
    });
});

describe('Download', () => {
    const getDownload = (worksheet, ...args) => {
        const createObjectURL = URL.createObjectURL;
        let blob = null;
        URL.createObjectURL = (b) => {
            blob = b;
            return 'blob:test';
        };
        try {
            worksheet.download(...args);
        } finally {
            URL.createObjectURL = createObjectURL;
        }
        // The Blob of jsdom has no text method. The download starts with a byte order mark.
        return new Promise((resolve) => {
            const reader = new window.FileReader();
            reader.onload = () => resolve(reader.result.replace(/^\uFEFF/, ''));
            reader.readAsText(blob);
        });
    };

    it('downloads the data with the headers and the nested headers', async () => {
        const worksheet = jspreadsheet(root, {
            worksheets: [
                {
                    data: [
                        ['a', 'b,c'],
                        ['say "hi"', 2],
                    ],
                    columns: [{ title: 'One' }, { title: 'Two' }],
                    nestedHeaders: [[{ title: 'Both', colspan: 2 }]],
                    csvDelimiter: ',',
                },
            ],
        })[0];

        expect(await getDownload(worksheet, true)).to.equal('Both,\r\nOne,Two\r\na,"b,c"\r\n"say ""hi""",2');
        expect(await getDownload(worksheet)).to.equal('a,"b,c"\r\n"say ""hi""",2');
    });

    it('downloads the shown values when processed', async () => {
        const worksheet = jspreadsheet(root, {
            worksheets: [{ data: [[1, '=A1+1']] }],
        })[0];

        expect(await getDownload(worksheet, false, false)).to.equal('1\t=A1+1');
        expect(await getDownload(worksheet, false, true)).to.equal('1\t2');
    });
});

describe('Paste in a calendar column', () => {
    it('reads a date with the format of the column that gets the value, not of the column at its position in the text', () => {
        const worksheet = jspreadsheet(root, {
            worksheets: [
                {
                    data: [['', '', '']],
                    columns: [
                        { type: 'calendar', options: { format: 'DD/MM/YYYY' } },
                        { type: 'text' },
                        { type: 'calendar', options: { format: 'DD/MM/YYYY' } },
                    ],
                },
            ],
        })[0];

        worksheet.paste(1, 0, '25/12/2024\t25/12/2024');

        expect(worksheet.getValueFromCoords(1, 0)).to.equal('25/12/2024');
        expect(worksheet.getValueFromCoords(2, 0)).to.match(/^2024-12-25/);
    });
});
