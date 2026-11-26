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
});
