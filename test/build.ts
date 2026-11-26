import { expect } from 'chai';

import { stubFetch } from './support/fetch';
import jspreadsheet, { type Jspreadsheet, type Worksheet } from './support/jspreadsheet';

// Create a spreadsheet and wait until it is built
function create(options: Jspreadsheet.SpreadsheetOptions) {
    return new Promise<Worksheet[]>((resolve) => {
        const worksheets = jspreadsheet(root, {
            ...options,
            onload() {
                resolve(worksheets);
            },
        });
    });
}

describe('Build of a worksheet', () => {
    let server: ReturnType<typeof stubFetch>;

    beforeEach(() => {
        server = stubFetch();
    });

    afterEach(() => {
        server.restore();
    });

    it('loads the data and the headers of a CSV file', async () => {
        server.responses['/data.csv'] = { body: 'Name;Age\nAnn;30\nBob;40' };

        const [worksheet] = await create({
            worksheets: [{ csv: '/data.csv', csvHeaders: true, csvDelimiter: ';', columns: [{}, { title: 'Years' }] }],
        });

        expect(server.requests[0].init.method).to.equal('GET');
        expect(worksheet.getData()).to.deep.equal([
            ['Ann', '30'],
            ['Bob', '40'],
        ]);
        // A title in the options takes precedence over the header of the file
        expect(worksheet.getHeader(0)).to.equal('Name');
        expect(worksheet.getHeader(1)).to.equal('Years');
    });

    it('loads the data from a url, as an array or in the data property of an object', async () => {
        server.responses['/rows'] = { body: [['a', 'b']] };
        server.responses['/object'] = { body: { data: [['c', 'd']] } };

        const [first, second] = await create({
            worksheets: [{ url: '/rows' }, { url: '/object' }],
        });

        expect(first.getData()).to.deep.equal([['a', 'b']]);
        expect(second.getData()).to.deep.equal([['c', 'd']]);
    });

    it('builds the worksheet with its other options when the request of the csv or url option fails', async () => {
        const errors: unknown[] = [];
        const error = console.error;
        console.error = (message: unknown) => errors.push(message);
        server.responses['/broken'] = { status: 500 };

        try {
            // Before the fix, the worksheets were never built and onload never came
            const loaded = await Promise.race([
                create({
                    worksheets: [
                        { csv: '/missing.csv', minDimensions: [2, 2] },
                        { url: '/broken', data: [['a']] },
                    ],
                }),
                new Promise<null>((resolve) => setTimeout(() => resolve(null), 200)),
            ]);

            expect(loaded).to.not.equal(null);
            const [first, second] = loaded!;
            expect(first.getData()).to.eql([
                ['', ''],
                ['', ''],
            ]);
            expect(second.getData()).to.eql([['a']]);
            expect(errors.map(String)).to.eql([
                'Error: Jspreadsheet: the GET request to /missing.csv failed with the status 404',
                'Error: Jspreadsheet: the GET request to /broken failed with the status 500',
            ]);
        } finally {
            console.error = error;
        }
    });

    it('lets onbeforerequest change the options of a request, replace them or cancel it', async () => {
        server.responses['/rows'] = { body: [['a']] };
        server.responses['/options'] = { body: ['x'] };
        const calls: [string, string][] = [];

        const [first, second, third] = await create({
            worksheets: [{ url: '/rows' }, { csv: '/data.csv', minDimensions: [1, 1] }, { data: [['x']], columns: [{ type: 'dropdown', url: '/options' }] }],
            onbeforerequest(instance, url, options) {
                calls.push([instance.options.worksheetName!, url]);
                if (url === '/rows') {
                    (options.headers as Record<string, string>).Authorization = 'Bearer token';
                } else if (url === '/options') {
                    return { method: 'GET', headers: { Accept: 'application/json' } };
                } else {
                    return false;
                }
            },
        });

        expect(calls.map(([, url]) => url)).to.have.members(['/rows', '/data.csv', '/options']);
        expect(calls.find(([, url]) => url === '/rows')![0]).to.equal(first.options.worksheetName);
        // The cancelled request is not sent, and the worksheet is built without its data
        expect(server.requests.map((request) => request.url)).to.have.members(['/rows', '/options']);
        expect(server.requests.find((request) => request.url === '/rows')!.init.headers).to.eql({
            'X-Requested-With': 'XMLHttpRequest',
            Authorization: 'Bearer token',
        });
        expect(server.requests.find((request) => request.url === '/options')!.init).to.eql({ method: 'GET', headers: { Accept: 'application/json' } });
        expect(first.getData()).to.eql([['a']]);
        expect(second.getData()).to.eql([['']]);
        await server.settle();
        expect(third.options.columns[0].source).to.eql(['x']);
    });

    it('loads the source of a dropdown column from its url before it creates the table', async () => {
        const options = [
            { id: 1, name: 'One' },
            { id: 2, name: 'Two' },
        ];
        server.responses['/options'] = { body: options };

        const [worksheet] = await create({
            worksheets: [{ data: [[2]], columns: [{ type: 'dropdown', url: '/options' }] }],
        });

        expect(worksheet.options.columns[0].source).to.deep.equal(options);
        expect(worksheet.getCell('A1').textContent).to.equal('Two');
    });

    it('calls beforeinit of the plugins before the worksheet is built, and init after', async () => {
        const calls: unknown[][] = [];

        const [worksheet] = await create({
            worksheets: [{ data: [['a']] }],
            plugins: {
                test() {
                    return {
                        beforeinit(instance) {
                            calls.push(['beforeinit', instance, !!instance.table]);
                        },
                        init(instance) {
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
        const calls: unknown[][] = [];
        const worksheets = await new Promise<Worksheet[]>((resolve) => {
            const worksheets = jspreadsheet(root, {
                worksheets: [{ worksheetName: 'A' }, { worksheetName: 'B' }, { worksheetName: 'C' }],
                onload() {
                    resolve(worksheets);
                },
                ondeleteworksheet(worksheet, position) {
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
