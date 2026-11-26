import { expect } from 'chai';

import jspreadsheet from './support/jspreadsheet';
import type { Dropdown } from 'jsuites';

// The element of a jSuites dropdown, which keeps the dropdown
type DropdownElement = HTMLElement & { dropdown: Dropdown };

describe('Use search', () => {
    it('search and resetSearch methods', () => {
        const instance = jspreadsheet(root, {
            tabs: true,
            worksheets: [
                {
                    search: true,
                    minDimensions: [7, 7],
                    data: [
                        ['Mazda', 2001, 2000, '2006-01-01 12:00:00'],
                        ['Peugeot', 2010, 5000, '2005-01-01 13:00:00'],
                        ['Honda Fit', 2009, 3000, '2004-01-01 14:01:00'],
                        ['Honda CRV', 2010, 6000, '2003-01-01 23:30:00'],
                    ],
                },
            ],
        });

        instance[0].search('Honda');

        expect(instance[0].searchInput.value).to.equal('Honda');

        const bodyTag = root.querySelector('tbody')!;

        expect(bodyTag.children.length).to.equal(2);

        expect(bodyTag.children[0].getAttribute('data-y')).to.equal('2');
        expect(bodyTag.children[1].getAttribute('data-y')).to.equal('3');

        instance[0].resetSearch();

        expect(instance[0].searchInput.value).to.equal('');

        expect(bodyTag.children[0].getAttribute('data-y')).to.equal('0');
        expect(bodyTag.children[1].getAttribute('data-y')).to.equal('1');
    });
});

describe('The view of a worksheet', () => {
    const data = [
        ['Honda', 1],
        ['Mazda', 2],
        ['Honda', 3],
        ['Peugeot', 4],
    ];

    it('shows every row on page 0 before a search, a filter or a page change', () => {
        const worksheet = jspreadsheet(root, { worksheets: [{ data }] })[0];

        expect(worksheet.view.results).to.equal(null);
        expect(worksheet.view.filters).to.eql([]);
        expect(worksheet.view.pageNumber).to.equal(0);
    });

    it('keeps the rows that a search finds until the search is reset', () => {
        const worksheet = jspreadsheet(root, { worksheets: [{ search: true, data }] })[0];

        worksheet.search('Honda');
        expect(worksheet.view.results).to.eql([0, 2]);

        worksheet.resetSearch();
        expect(worksheet.view.results).to.equal(null);
    });

    it('keeps the values of a filter and the rows it finds until the filters are reset', () => {
        const worksheet = jspreadsheet(root, { worksheets: [{ filters: true, data }] })[0];

        worksheet.openFilter(0);
        const dropdown = worksheet.filter!.children[1].querySelector<DropdownElement>('.jdropdown')!.dropdown;
        dropdown.setValue(['Honda']);
        dropdown.close();

        expect(worksheet.view.filters[0]).to.eql(['Honda']);
        expect(worksheet.view.results).to.eql([0, 2]);

        worksheet.resetFilters();

        expect(worksheet.view.filters[0]).to.equal(null);
        expect(worksheet.view.results).to.equal(null);
    });

    it('keeps the page that the worksheet shows', () => {
        const worksheet = jspreadsheet(root, { worksheets: [{ pagination: 2, data }] })[0];

        worksheet.page(1);

        expect(worksheet.view.pageNumber).to.equal(1);
        expect(root.querySelector('tbody')!.children[0].getAttribute('data-y')).to.equal('2');
    });

    it('goes back to page 0 when a search starts', () => {
        const worksheet = jspreadsheet(root, { worksheets: [{ search: true, pagination: 2, data }] })[0];

        worksheet.page(1);
        worksheet.search('Peugeot');

        expect(worksheet.view.pageNumber).to.equal(0);
        expect(worksheet.view.results).to.eql([3]);
    });

    it('is not shared between worksheets', async () => {
        const worksheets = jspreadsheet(root, { worksheets: [{ search: true, data }, { data }] });
        await new Promise((resolve) => setTimeout(resolve));

        worksheets[0].search('Mazda');

        expect(worksheets[0].view).to.not.equal(worksheets[1].view);
        expect(worksheets[1].view.results).to.equal(null);
    });
    it('sorts the rows that a filter found and keeps the filter', () => {
        const worksheet = jspreadsheet(root, { worksheets: [{ filters: true, data }] })[0];

        worksheet.openFilter(0);
        const dropdown = worksheet.filter!.children[1].querySelector<DropdownElement>('.jdropdown')!.dropdown;
        dropdown.setValue(['Honda']);
        dropdown.close();

        worksheet.orderBy(1, 1);

        expect(worksheet.getColumnData(1)).to.eql([4, 3, 2, 1]);
        expect(worksheet.view.filters[0]).to.eql(['Honda']);
        expect(worksheet.view.results).to.eql([1, 3]);
        const shown = Array.from(root.querySelector('tbody')!.children).map((tr) => tr.getAttribute('data-y'));
        expect(shown).to.eql(['1', '3']);
    });

    it('tells the page of a row, among the rows that a search found', () => {
        const worksheet = jspreadsheet(root, { worksheets: [{ pagination: 1, search: true, data }] })[0];

        expect(worksheet.whichPage(2)).to.equal(2);

        worksheet.search('Honda');

        expect(worksheet.whichPage(2)).to.equal(1);
        expect(worksheet.quantiyOfPages()).to.equal(2);
    });
});
