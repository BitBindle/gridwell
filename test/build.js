const { expect } = require('chai');
const jSuites = require('jsuites');

const jspreadsheet = require('../dist/index.js');

// Create a spreadsheet and wait until it is built
const create = function (options) {
    return new Promise((resolve) => {
        const worksheets = jspreadsheet(root, {
            ...options,
            onload: function () {
                resolve(worksheets);
            },
        });
    });
};

describe('Build of a worksheet', () => {
    let ajax;
    let requests;
    // The response of the stubbed ajax, by url
    let responses;

    beforeEach(() => {
        ajax = jSuites.ajax;
        requests = [];
        responses = {};
        jSuites.ajax = function (options, complete) {
            // An array of requests, then a callback when all of them are done
            const list = Array.isArray(options) ? options : [options];
            for (const request of list) {
                requests.push(request);
                request.success.call(request, responses[request.url]);
            }
            if (complete) {
                complete();
            }
        };
    });

    afterEach(() => {
        jSuites.ajax = ajax;
    });

    it('loads the data and the headers of a CSV file', async () => {
        responses['/data.csv'] = 'Name;Age\nAnn;30\nBob;40';

        const [worksheet] = await create({
            worksheets: [{ csv: '/data.csv', csvHeaders: true, csvDelimiter: ';', columns: [{}, { title: 'Years' }] }],
        });

        expect(requests[0].dataType).to.equal('text');
        expect(worksheet.getData()).to.deep.equal([
            ['Ann', '30'],
            ['Bob', '40'],
        ]);
        // A title in the options takes precedence over the header of the file
        expect(worksheet.getHeader(0)).to.equal('Name');
        expect(worksheet.getHeader(1)).to.equal('Years');
    });

    it('loads the data from a url, as an array or in the data property of an object', async () => {
        responses['/rows'] = [['a', 'b']];
        responses['/object'] = { data: [['c', 'd']] };

        const [first, second] = await create({
            worksheets: [{ url: '/rows' }, { url: '/object' }],
        });

        expect(first.getData()).to.deep.equal([['a', 'b']]);
        expect(second.getData()).to.deep.equal([['c', 'd']]);
    });

    it('loads the source of a dropdown column from its url before it creates the table', async () => {
        responses['/options'] = [
            { id: 1, name: 'One' },
            { id: 2, name: 'Two' },
        ];

        const [worksheet] = await create({
            worksheets: [{ data: [[2]], columns: [{ type: 'dropdown', url: '/options' }] }],
        });

        expect(worksheet.options.columns[0].source).to.deep.equal(responses['/options']);
        expect(worksheet.getCell('A1').textContent).to.equal('Two');
    });

    it('calls beforeinit of the plugins before the worksheet is built, and init after', async () => {
        const calls = [];

        const [worksheet] = await create({
            worksheets: [{ data: [['a']] }],
            plugins: {
                test: function () {
                    return {
                        beforeinit: function (instance) {
                            calls.push(['beforeinit', instance, !!instance.table]);
                        },
                        init: function (instance) {
                            calls.push(['init', instance, instance.getValue('A1')]);
                        },
                    };
                },
            },
        });

        expect(calls).to.deep.equal([
            ['beforeinit', worksheet, false],
            ['init', worksheet, 'a'],
        ]);
    });

    it('applies the classes option to the cells', async () => {
        const [worksheet] = await create({
            worksheets: [{ data: [['a', 'b']], classes: { B1: 'highlight' } }],
        });

        expect(worksheet.getCell('A1').classList.contains('highlight')).to.equal(false);
        expect(worksheet.getCell('B1').classList.contains('highlight')).to.equal(true);
    });
});

describe('Worksheets of a spreadsheet', () => {
    it('names a new worksheet after the largest number of the worksheets called Sheet', async () => {
        const worksheets = await create({
            worksheets: [{ worksheetName: 'Sheet1' }, { worksheetName: 'Sheet5' }, { worksheetName: 'Sheet9b' }],
        });

        worksheets[0].createWorksheet({ minDimensions: [1, 1] });

        expect(worksheets.length).to.equal(4);
        expect(worksheets[3].options.worksheetName).to.equal('Sheet6');
        expect(worksheets[0].parent.config.worksheets[3]).to.equal(worksheets[3].options);
        expect(worksheets[3].element.jssWorksheet).to.equal(worksheets[3]);
    });

    it('removes a worksheet and calls ondeleteworksheet with it and its position', async () => {
        const calls = [];
        const worksheets = await new Promise((resolve) => {
            const worksheets = jspreadsheet(root, {
                worksheets: [{ worksheetName: 'A' }, { worksheetName: 'B' }, { worksheetName: 'C' }],
                onload: function () {
                    resolve(worksheets);
                },
                ondeleteworksheet: function (worksheet, position) {
                    calls.push([worksheet, position, this]);
                },
            });
        });
        const second = worksheets[1];
        const spreadsheet = second.parent;

        worksheets[0].deleteWorksheet(1);

        expect(worksheets.map((w) => w.options.worksheetName)).to.deep.equal(['A', 'C']);
        expect(calls).to.deep.equal([[second, 1, spreadsheet]]);
    });
});
