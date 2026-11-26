import { expect } from 'chai';
import jSuites from 'jsuites';

import jspreadsheet, { type Jspreadsheet, type Worksheet } from './support/jspreadsheet';

// A request of the persistence, as the stub of jSuites.ajax reads it
type Request = { url: string; method: string; data: { data: string }; success: (result: object) => void };

describe('Events of the spreadsheet', () => {
    it('gives the handlers the worksheet or the spreadsheet that triggers the event as this', () => {
        const calls: unknown[][] = [];
        let spreadsheet: Jspreadsheet.SpreadsheetInstance | undefined;
        const worksheet = jspreadsheet(root, {
            worksheets: [{ data: [['a']] }],
            // onevent takes its arguments as a rest of a union of tuples, which a handler with named parameters does not fit (issue #133)
            onevent(...[event]) {
                if (event === 'onchange' || event === 'onload') {
                    calls.push(['onevent', event, this]);
                }
            },
            onchange() {
                calls.push(['onchange', this]);
            },
            onload(instance) {
                spreadsheet = instance;
                calls.push(['onload', this]);
            },
            plugins: {
                test() {
                    return {
                        // As onevent above (issue #133)
                        onevent(...[event]) {
                            if (event === 'onchange') {
                                calls.push(['plugin', event, this]);
                            }
                        },
                    };
                },
            },
        })[0];

        worksheet.setValue('A1', 'b');

        return new Promise((resolve) => setTimeout(resolve)).then(() => {
            expect(spreadsheet).to.equal(worksheet.parent);
            expect(calls).to.eql([
                ['onevent', 'onchange', worksheet],
                ['onchange', worksheet],
                ['plugin', 'onchange', worksheet],
                ['onevent', 'onload', spreadsheet],
                ['onload', spreadsheet],
            ]);
        });
    });

    it('returns the result of the last handler from dispatch, and no handler runs while ignoreEvents is set', () => {
        const worksheet = jspreadsheet(root, {
            worksheets: [{ data: [['a']] }],
            onbeforechange() {
                return 'replaced';
            },
        })[0];

        expect(worksheet.dispatch('onbeforechange', worksheet, worksheet.getCell('A1'), 0, 0, 'b')).to.equal('replaced');

        worksheet.parent.ignoreEvents = true;
        expect(worksheet.dispatch('onbeforechange', worksheet, worksheet.getCell('A1'), 0, 0, 'b')).to.equal(null);
        worksheet.parent.ignoreEvents = false;
    });

    describe('Persistence', () => {
        let ajax: typeof jSuites.ajax;
        let requests: Request[];

        beforeEach(() => {
            ajax = jSuites.ajax;
            requests = [];
            jSuites.ajax = function (options: Request) {
                requests.push(options);
                options.success({});
            };
        });

        afterEach(() => {
            jSuites.ajax = ajax;
        });

        it('posts the changed rows to the url and triggers onsave', () => {
            const saved: unknown[] = [];
            const worksheet = jspreadsheet(root, {
                worksheets: [{ data: [{ first: 'a', second: 'b' }], columns: [{ name: 'first' }, { name: 'second' }], persistence: '/save' }],
                onsave(spreadsheet, instance, data) {
                    saved.push(data);
                },
            })[0];

            worksheet.setValue('A1', 'x');
            worksheet.setValue('B1', 'y');

            expect(requests.length).to.equal(2);
            expect(requests[0].url).to.equal('/save');
            expect(requests[0].method).to.equal('POST');
            expect(JSON.parse(requests[0].data.data)).to.eql([{ row: 0, data: { first: 'x' } }]);
            expect(JSON.parse(requests[1].data.data)).to.eql([{ row: 0, data: { second: 'y' } }]);
            expect(saved).to.eql([[{ row: 0, data: { first: 'x' } }], [{ row: 0, data: { second: 'y' } }]]);
        });

        it('posts the data that onbeforesave returns, or nothing when it returns false', () => {
            let result: { replaced: boolean } | false = { replaced: true };
            const worksheet = jspreadsheet(root, {
                worksheets: [{ data: [['a']], persistence: '/other' }],
                onbeforesave() {
                    return result;
                },
            })[0];

            worksheet.setValue('A1', 'x');

            expect(requests.length).to.equal(1);
            expect(requests[0].url).to.equal('/other');
            expect(JSON.parse(requests[0].data.data)).to.eql({ replaced: true });

            result = false;
            worksheet.setValue('A1', 'y');

            expect(requests.length).to.equal(1);
        });

        it('triggers onbeforesave and onsave for an undo and a redo, so that onbeforesave can cancel the post (issue #57)', () => {
            const calls: unknown[][] = [];
            let result: false | undefined = undefined;
            const worksheet = jspreadsheet(root, {
                worksheets: [{ data: [['a', 'b']], persistence: '/save' }],
                onbeforesave: (spreadsheet, instance, data) => {
                    calls.push(['onbeforesave', data]);
                    return result;
                },
                onsave: (spreadsheet, instance, data) => calls.push(['onsave', data]),
            })[0];

            worksheet.setValue('B1', 'x');
            worksheet.undo();

            expect(requests.length).to.equal(2);
            expect(calls).to.eql([
                ['onbeforesave', [{ row: 0, data: { 1: 'x' } }]],
                ['onsave', [{ row: 0, data: { 1: 'x' } }]],
                ['onbeforesave', [{ row: 0, data: { 1: 'b' } }]],
                ['onsave', [{ row: 0, data: { 1: 'b' } }]],
            ]);

            result = false;
            worksheet.redo();

            expect(requests.length).to.equal(2);
            expect(worksheet.parent.ignoreEvents).to.equal(false);
        });

        it('gives the changes to the persistence of a plugin', () => {
            const changes: unknown[][] = [];
            const worksheet = jspreadsheet(root, {
                worksheets: [{ data: [['a']] }],
                plugins: {
                    test() {
                        return {
                            persistence(instance, method, args) {
                                changes.push([instance, method, args.data.map((change) => change.value)]);
                            },
                        };
                    },
                },
            })[0];

            worksheet.setValue('A1', 'x');

            expect(changes).to.eql([[worksheet, 'setValue', ['x']]]);
            expect(requests.length).to.equal(0);
        });
    });
});

describe('Spreadsheet methods', () => {
    it('gets and sets the options of the spreadsheet and of a worksheet', () => {
        const worksheet = jspreadsheet(root, {
            toolbar: false,
            worksheets: [{ data: [['a']] }],
        })[0];
        const spreadsheet = worksheet.parent;

        expect(spreadsheet.getConfig()).to.equal(spreadsheet.config);
        expect(worksheet.getConfig()).to.equal(worksheet.options);

        worksheet.setConfig({ columnSorting: false });
        expect(worksheet.options.columnSorting).to.equal(false);
        // @ts-expect-error -- the test checks that the spreadsheet has no option of a worksheet
        expect(spreadsheet.config.columnSorting).to.equal(undefined);

        worksheet.setConfig({ tabs: true }, true);
        expect(spreadsheet.config.tabs).to.equal(true);
        // @ts-expect-error -- the test checks that the worksheet has no option of the spreadsheet
        expect(worksheet.options.tabs).to.equal(undefined);

        spreadsheet.setConfig({ toolbar: true });
        expect(spreadsheet.config.toolbar).to.equal(true);
        expect(spreadsheet.toolbar.style.display).to.equal('');
    });

    it('adds plugins with setPlugins', () => {
        const worksheet = jspreadsheet(root, { worksheets: [{ data: [['a']] }] })[0];
        const spreadsheet = worksheet.parent;
        const received: unknown[] = [];

        spreadsheet.setPlugins({
            test(instance, options, config) {
                received.push(this, instance, options, config);
                // The plugin is an object that the test finds again, without any method of a plugin
                return { name: 'test' } as Jspreadsheet.Plugin;
            },
        });

        expect(spreadsheet.plugins.test).to.eql({ name: 'test' });
        expect(received).to.eql([jspreadsheet, spreadsheet, {}, spreadsheet.config]);
    });

    it('turns the fullscreen on and off', () => {
        const worksheet = jspreadsheet(root, { worksheets: [{ data: [['a']] }] })[0];
        const spreadsheet = worksheet.parent;

        spreadsheet.fullscreen();
        expect(spreadsheet.element.classList.contains('fullscreen')).to.equal(true);
        expect(spreadsheet.config.fullscreen).to.equal(true);

        spreadsheet.fullscreen(true);
        expect(spreadsheet.element.classList.contains('fullscreen')).to.equal(true);

        spreadsheet.fullscreen(false);
        expect(spreadsheet.element.classList.contains('fullscreen')).to.equal(false);
    });

    it('gives the index of the worksheet that the spreadsheet shows', async () => {
        // The spreadsheet builds the second worksheet after the first
        const worksheets = await new Promise<Worksheet[]>((resolve) => {
            const instance = jspreadsheet(root, { worksheets: [{ data: [['a']] }, { data: [['b']] }], onload: () => resolve(instance) });
        });

        expect(worksheets[0].parent.getWorksheetActive()).to.equal(0);
        expect(worksheets[1].getWorksheetActive()).to.equal(0);

        worksheets[0].openWorksheet(1);

        expect(worksheets[0].parent.getWorksheetActive()).to.equal(1);
        expect(worksheets[0].getWorksheetActive()).to.equal(1);
    });
});
