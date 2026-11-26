import { expect } from 'chai';

import jspreadsheet from './support/jspreadsheet';

describe('Comment tests', () => {
    it('Set comment', () => {
        const instance = jspreadsheet(root, {
            worksheets: [
                {
                    data: [
                        ['US', 'Cheese', '2019-02-12'],
                        ['CA', 'Apples', '2019-03-01'],
                        ['CA', 'Carrots', '2018-11-10'],
                        ['BR', 'Oranges', '2019-01-12'],
                    ],
                    columns: [{ width: '300px' }, { width: '200px' }, { width: '200px' }],
                    allowComments: true,
                },
            ],
        });

        const table = root.querySelector('tbody')!;
        const rows = table.children;

        instance[0].setComments('C2', 'Test');

        expect(rows[1].children[3].getAttribute('title')).to.equal('Test');

        instance[0].setComments('C2', '');

        expect(rows[1].children[3].getAttribute('title')).to.equal('');
    });

    it('Get comment', () => {
        const instance = jspreadsheet(root, {
            worksheets: [
                {
                    data: [
                        ['US', 'Cheese', '2019-02-12'],
                        ['CA', 'Apples', '2019-03-01'],
                        ['CA', 'Carrots', '2018-11-10'],
                        ['BR', 'Oranges', '2019-01-12'],
                    ],
                    columns: [{ width: '300px' }, { width: '200px' }, { width: '200px' }],
                    allowComments: true,
                },
            ],
        });

        instance[0].setComments('B3', 'something');

        expect(instance[0].getComments('B3')).to.equal('something');
    });

    it('setComments history', () => {
        const instance = jspreadsheet(root, {
            worksheets: [
                {
                    data: [
                        ['US', 'Cheese', '2019-02-12'],
                        ['CA', 'Apples', '2019-03-01'],
                        ['CA', 'Carrots', '2018-11-10'],
                        ['BR', 'Oranges', '2019-01-12'],
                    ],
                    columns: [{ width: '300px' }, { width: '200px' }, { width: '200px' }],
                    allowComments: true,
                },
            ],
        });

        const table = root.querySelector('tbody')!;
        const rows = table.children;

        instance[0].setComments('C2', 'Test');

        expect(rows[1].children[3].getAttribute('title')).to.equal('Test');

        instance[0].undo();

        expect(rows[1].children[3].getAttribute('title')).to.equal('');

        instance[0].redo();

        expect(rows[1].children[3].getAttribute('title')).to.equal('Test');
    });

    it('loads the comments option without a history entry or oncomments (issue #87)', () => {
        const calls: string[] = [];
        const worksheet = jspreadsheet(root, {
            worksheets: [{ minDimensions: [2, 2], comments: { B1: 'A note' } }],
            oncomments: () => calls.push('oncomments'),
        })[0];

        expect(worksheet.history.records.length).to.equal(0);
        expect(calls).to.eql([]);

        worksheet.undo();

        expect(worksheet.getComments('B1')).to.equal('A note');
        expect(worksheet.getCell('B1').classList.contains('jss_comments')).to.equal(true);
    });
});
