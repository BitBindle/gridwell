const { expect } = require('chai');

const jspreadsheet = require('../dist/index.js');

describe('Style', () => {
    it('Initial worksheet styles are not in the history', () => {
        const instance = jspreadsheet(root, {
            worksheets: [{ minDimensions: [2, 2], style: { A1: 'font-weight:bold' } }],
        });

        expect(instance[0].history.records.length).to.equal(0);

        instance[0].undo();
        expect(instance[0].getStyle('A1', 'font-weight')).to.equal('bold');
    });

    it('setStyle adds no history when ignoreHistoryAndEvents is true', () => {
        const instance = jspreadsheet(root, {
            worksheets: [{ minDimensions: [2, 2] }],
        });

        instance[0].setStyle('A1', 'color', 'red', false, true);
        instance[0].setStyle({ B1: 'color:blue' }, null, null, false, true);

        expect(instance[0].getStyle('A1', 'color')).to.equal('red');
        expect(instance[0].getStyle('B1', 'color')).to.equal('blue');
        expect(instance[0].history.records.length).to.equal(0);
    });

    it('setStyle adds the change to the history by default', () => {
        const instance = jspreadsheet(root, {
            worksheets: [{ minDimensions: [2, 2] }],
        });

        instance[0].setStyle('A1', 'color', 'red');
        expect(instance[0].history.records.length).to.equal(1);

        instance[0].undo();
        expect(instance[0].getStyle('A1', 'color')).to.equal('');
    });

    it('resetStyle adds no history when ignoreHistoryAndEvents is true', () => {
        const instance = jspreadsheet(root, {
            worksheets: [{ minDimensions: [2, 2] }],
        });

        instance[0].setStyle('A1', 'color', 'red', false, true);
        instance[0].resetStyle({ A1: 'font-weight:bold' }, true);

        expect(instance[0].getStyle('A1', 'color')).to.equal('');
        expect(instance[0].getStyle('A1', 'font-weight')).to.equal('bold');
        expect(instance[0].history.records.length).to.equal(0);
    });

    it('setStyle does not change the styles that it gets', () => {
        const instance = jspreadsheet(root, {
            worksheets: [{ minDimensions: [2, 2] }],
        });

        const styles = { A1: ['color:red', 'font-weight:bold'] };
        instance[0].setStyle(styles);

        expect(styles).to.eql({ A1: ['color:red', 'font-weight:bold'] });
        expect(instance[0].getStyle('A1', 'color')).to.equal('red');
        expect(instance[0].getStyle('A1', 'font-weight')).to.equal('bold');
    });

    it('setStyle keeps a value that has a colon', () => {
        const instance = jspreadsheet(root, {
            worksheets: [{ minDimensions: [2, 2] }],
        });

        // No space after the colons: jsdom drops a value that starts with a space, which browsers trim
        instance[0].setStyle({ A1: 'background-image:url(http://example.com/a.png);color:red' });
        expect(instance[0].getStyle('A1', 'background-image')).to.equal('url(http://example.com/a.png)');
        expect(instance[0].getStyle('A1', 'color')).to.equal('red');

        // Undo and redo read the style from the history
        instance[0].undo();
        expect(instance[0].getStyle('A1', 'background-image')).to.equal('');
        instance[0].redo();
        expect(instance[0].getStyle('A1', 'background-image')).to.equal('url(http://example.com/a.png)');
    });

    it('keeps a value that has a semicolon, also after undo and redo (issue #48)', () => {
        const worksheet = jspreadsheet(root, {
            worksheets: [{ minDimensions: [2, 2], style: { B1: 'color:red;background-image:url(data:image/png;base64,AAA=)' } }],
        })[0];
        const url = 'url(data:image/png;base64,iVBORw0KGgo=)';

        expect(worksheet.getStyle('B1', 'background-image')).to.contain('data:image/png;base64,AAA=');
        expect(worksheet.getStyle('B1', 'color')).to.equal('red');

        worksheet.setStyle({ A1: `background-image:${url};color:blue` });

        expect(worksheet.getStyle('A1', 'background-image')).to.contain('data:image/png;base64,iVBORw0KGgo=');
        expect(worksheet.getStyle('A1', 'color')).to.equal('blue');

        worksheet.undo();
        expect(worksheet.getStyle('A1', 'background-image')).to.equal('');

        worksheet.redo();
        expect(worksheet.getStyle('A1', 'background-image')).to.contain('data:image/png;base64,iVBORw0KGgo=');
        expect(worksheet.getStyle('A1', 'color')).to.equal('blue');
    });

    it('applies a value with a space after the colon, and toggles it off as a value without one (issue #67)', () => {
        const worksheet = jspreadsheet(root, { worksheets: [{ minDimensions: [2, 2], style: { A1: 'color: red', B1: 'color:red' } }] })[0];

        expect(worksheet.getStyle('A1', 'color')).to.equal('red');
        expect(worksheet.getStyle('B1', 'color')).to.equal('red');

        worksheet.setStyle({ A1: 'color: red', B1: 'color:red' });

        expect(worksheet.getStyle('A1', 'color')).to.equal('');
        expect(worksheet.getStyle('B1', 'color')).to.equal('');
        expect(worksheet.history.records[0].newValue).to.eql({ A1: 'color:', B1: 'color:' });
        expect(worksheet.history.records[0].oldValue).to.eql({ A1: 'color:red', B1: 'color:red' });
    });
});
