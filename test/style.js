const { expect } = require('chai');

const jspreadsheet = require('../dist/index.js');

describe('Style', () => {
    it('Initial worksheet styles are not in the history', () => {
        const instance = jspreadsheet(root, {
            worksheets: [{ minDimensions: [2, 2], style: { A1: 'font-weight:bold' } }],
        });

        expect(instance[0].history.length).to.equal(0);

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
        expect(instance[0].history.length).to.equal(0);
    });

    it('setStyle adds the change to the history by default', () => {
        const instance = jspreadsheet(root, {
            worksheets: [{ minDimensions: [2, 2] }],
        });

        instance[0].setStyle('A1', 'color', 'red');
        expect(instance[0].history.length).to.equal(1);

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
        expect(instance[0].history.length).to.equal(0);
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
});
