import { expect } from 'chai';

import jspreadsheet, { type Jspreadsheet } from './support/jspreadsheet';

describe('Sorting tests', () => {
    it('Default sorting', () => {
        const instance = jspreadsheet(root, {
            worksheets: [
                {
                    data: [
                        ['Mazda', 2001, 2000, '2006-01-01', '453.00', '2', '=E1*F1'],
                        ['Peugeot', 2010, 5000, '2005-01-01', '23.00', '5', '=E2*F2'],
                        ['Honda Fit', 2009, 3000, '2004-01-01', '214.00', '3', '=E3*F3'],
                        ['Honda CRV', 2010, 6000, '2003-01-01', '56.11', '2', '=E4*F4'],
                    ],
                },
            ],
        });

        instance[0].orderBy(5);

        expect(instance[0].options.data).to.eql([
            ['Mazda', 2001, 2000, '2006-01-01', '453.00', '2', '=E1*F1'],
            ['Honda CRV', 2010, 6000, '2003-01-01', '56.11', '2', '=E2*F2'],
            ['Honda Fit', 2009, 3000, '2004-01-01', '214.00', '3', '=E3*F3'],
            ['Peugeot', 2010, 5000, '2005-01-01', '23.00', '5', '=E4*F4'],
        ]);

        instance[0].orderBy(5);

        expect(instance[0].options.data).to.eql([
            ['Peugeot', 2010, 5000, '2005-01-01', '23.00', '5', '=E1*F1'],
            ['Honda Fit', 2009, 3000, '2004-01-01', '214.00', '3', '=E2*F2'],
            ['Mazda', 2001, 2000, '2006-01-01', '453.00', '2', '=E3*F3'],
            ['Honda CRV', 2010, 6000, '2003-01-01', '56.11', '2', '=E4*F4'],
        ]);
    });

    it('Custom sorting', () => {
        const instance = jspreadsheet(root, {
            worksheets: [
                {
                    data: [
                        ['Mazda', 2001, 2000, '2006-01-01', '453.00', '2', '=E1*F1'],
                        ['Peugeot', 1800, 5000, '2005-01-01', '23.00', '5', '=E2*F2'],
                        ['test', 2009, 3000, '2004-01-01', '214.00', '3', '=E3*F3'],
                        ['Honda CRV', 1900, 6000, '2003-01-01', '56.11', '2', '=E4*F4'],
                    ],
                    // @ts-expect-error -- sorting is an option of the spreadsheet, so the library ignores it on a worksheet and sorts with its default
                    sorting(direction: 0 | 1) {
                        return function (a: Jspreadsheet.SortingItem, b: Jspreadsheet.SortingItem) {
                            const valueA = a[1];
                            const valueB = b[1];

                            if (valueA === 'test') {
                                return direction ? 1 : -1;
                            }

                            if (valueB === 'test') {
                                return direction ? -1 : 1;
                            }

                            // Consider blank rows in the sorting
                            if (!direction) {
                                return valueA > valueB ? 1 : valueA < valueB ? -1 : 0;
                            } else {
                                return valueA > valueB ? -1 : valueA < valueB ? 1 : 0;
                            }
                        };
                    },
                },
            ],
        });

        instance[0].orderBy(0, 1);

        expect(instance[0].options.data).to.eql([
            ['test', 2009, 3000, '2004-01-01', '214.00', '3', '=E1*F1'],
            ['Peugeot', 1800, 5000, '2005-01-01', '23.00', '5', '=E2*F2'],
            ['Mazda', 2001, 2000, '2006-01-01', '453.00', '2', '=E3*F3'],
            ['Honda CRV', 1900, 6000, '2003-01-01', '56.11', '2', '=E4*F4'],
        ]);
    });

    it('orderBy history', () => {
        const instance = jspreadsheet(root, {
            worksheets: [
                {
                    data: [
                        ['Mazda', 2001, 2000, '2006-01-01', '453.00', '2', '=E1*F1'],
                        ['Peugeot', 2010, 5000, '2005-01-01', '23.00', '5', '=E2*F2'],
                        ['Honda Fit', 2009, 3000, '2004-01-01', '214.00', '3', '=E3*F3'],
                        ['Honda CRV', 2010, 6000, '2003-01-01', '56.11', '2', '=E4*F4'],
                    ],
                },
            ],
        });

        instance[0].orderBy(5);

        expect(instance[0].options.data).to.eql([
            ['Mazda', 2001, 2000, '2006-01-01', '453.00', '2', '=E1*F1'],
            ['Honda CRV', 2010, 6000, '2003-01-01', '56.11', '2', '=E2*F2'],
            ['Honda Fit', 2009, 3000, '2004-01-01', '214.00', '3', '=E3*F3'],
            ['Peugeot', 2010, 5000, '2005-01-01', '23.00', '5', '=E4*F4'],
        ]);

        // @ts-expect-error -- the test passes an argument, which undo ignores
        instance[0].undo(5);

        expect(instance[0].options.data).to.eql([
            ['Mazda', 2001, 2000, '2006-01-01', '453.00', '2', '=E1*F1'],
            ['Peugeot', 2010, 5000, '2005-01-01', '23.00', '5', '=E2*F2'],
            ['Honda Fit', 2009, 3000, '2004-01-01', '214.00', '3', '=E3*F3'],
            ['Honda CRV', 2010, 6000, '2003-01-01', '56.11', '2', '=E4*F4'],
        ]);

        // @ts-expect-error -- the test passes an argument, which redo ignores
        instance[0].redo(5);

        expect(instance[0].options.data).to.eql([
            ['Mazda', 2001, 2000, '2006-01-01', '453.00', '2', '=E1*F1'],
            ['Honda CRV', 2010, 6000, '2003-01-01', '56.11', '2', '=E2*F2'],
            ['Honda Fit', 2009, 3000, '2004-01-01', '214.00', '3', '=E3*F3'],
            ['Peugeot', 2010, 5000, '2005-01-01', '23.00', '5', '=E4*F4'],
        ]);
    });
});
