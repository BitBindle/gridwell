import { expect } from 'chai';

import jspreadsheet, { type Jspreadsheet } from './support/jspreadsheet';

describe('The confirm and alert options', () => {
    // The library calls the global confirm and alert
    let calls: string[];
    let confirm: typeof global.confirm;
    let alert: typeof global.alert;

    beforeEach(() => {
        calls = [];
        confirm = global.confirm;
        alert = global.alert;
        global.confirm = (text?: string) => {
            calls.push(`confirm ${text}`);
            return true;
        };
        global.alert = (text?: string) => {
            calls.push(`alert ${text}`);
        };
    });

    afterEach(() => {
        global.confirm = confirm;
        global.alert = alert;
    });

    function build(options: Partial<Jspreadsheet.SpreadsheetOptions>) {
        return jspreadsheet(root, { worksheets: [{ minDimensions: [3, 5], mergeCells: { A1: [1, 3] } }], ...options })[0];
    }

    it('asks the browser by default before an insert destroys merged cells', () => {
        const worksheet = build({});

        worksheet.insertRow(1, 1);

        expect(calls).to.eql(['confirm This action will destroy any existing merged cells. Are you sure?']);
        expect(worksheet.getMerge()).to.eql({});
    });

    it('does the action without a dialog when confirm is false', () => {
        const worksheet = build({ confirm: false });

        worksheet.insertRow(1, 1);

        expect(calls).to.eql([]);
        expect(worksheet.getMerge()).to.eql({});
        expect(worksheet.rows.length).to.equal(6);
    });

    it('lets the confirm function decide with the translated question', () => {
        const questions: string[] = [];
        let answer = false;
        const worksheet = build({
            confirm: (text) => {
                questions.push(text);
                return answer;
            },
        });

        worksheet.insertRow(1, 1);
        expect(worksheet.rows.length).to.equal(5);
        expect(worksheet.getMerge()).to.eql({ A1: [1, 3] });

        answer = true;
        worksheet.insertRow(1, 1);
        expect(worksheet.rows.length).to.equal(6);

        expect(questions).to.eql([
            'This action will destroy any existing merged cells. Are you sure?',
            'This action will destroy any existing merged cells. Are you sure?',
        ]);
        expect(calls).to.eql([]);
    });

    it('shows a message with the alert function, or nothing when alert is false', () => {
        const messages: string[] = [];
        const shown = build({ alert: (text) => messages.push(text) });
        shown.setMerge('A2', 2, 1);
        expect(messages).to.eql(['There is a conflict with another merged cell']);

        const silent = build({ alert: false });
        silent.setMerge('A2', 2, 1);
        expect(calls).to.eql([]);
        expect(silent.getMerge()).to.eql({ A1: [1, 3] });
    });

    it('does each insert and delete of rows and columns over a merged range without a dialog when confirm is false', () => {
        const actions: [string, (w: ReturnType<typeof build>) => unknown][] = [
            ['insertRow', (w) => w.insertRow(1, 1)],
            ['deleteRow', (w) => w.deleteRow(1, 1)],
            ['insertColumn', (w) => w.insertColumn(1, 0)],
            ['deleteColumn', (w) => w.deleteColumn(0, 1)],
        ];
        for (const [name, action] of actions) {
            const worksheet = jspreadsheet(root, { worksheets: [{ minDimensions: [3, 5], mergeCells: { A1: [2, 3] } }], confirm: false })[0];

            expect(action(worksheet), name).to.not.equal(false);
            expect(worksheet.getMerge(), name).to.eql({});
        }
        expect(calls).to.eql([]);
    });

    it('cancels the action for each answer of the confirm function other than true', () => {
        for (const answer of [1, 'yes', undefined, null, Promise.resolve(true)]) {
            // @ts-expect-error -- a function in JavaScript can return any value
            const worksheet = build({ confirm: () => answer });

            expect(worksheet.insertRow(1, 1)).to.equal(false);
            expect(worksheet.rows.length).to.equal(5);
        }
    });

    it('gives the translated texts to the confirm and alert functions', () => {
        jspreadsheet.setDictionary({
            'This action will destroy any existing merged cells. Are you sure?': 'Verbundene Zellen gehen verloren. Fortfahren?',
            'There is a conflict with another merged cell': 'Konflikt mit einer verbundenen Zelle',
        });
        try {
            const texts: string[] = [];
            const worksheet = build({
                confirm: (text) => {
                    texts.push(text);
                    return false;
                },
                alert: (text) => texts.push(text),
            });

            worksheet.insertRow(1, 1);
            worksheet.setMerge('A2', 2, 1);

            expect(texts).to.eql(['Verbundene Zellen gehen verloren. Fortfahren?', 'Konflikt mit einer verbundenen Zelle']);
        } finally {
            // An empty translation shows the English text again
            jspreadsheet.setDictionary({
                'This action will destroy any existing merged cells. Are you sure?': '',
                'There is a conflict with another merged cell': '',
            });
        }
    });
});
