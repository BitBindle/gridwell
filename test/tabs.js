const { expect } = require('chai');

const jspreadsheet = require('../dist/index.js');

const wait = function (ms) {
    return new Promise((resolve) => setTimeout(resolve, ms));
};

// The constructor builds the worksheets after it returns. Resolve when all worksheets are ready.
const create = function (options) {
    return new Promise((resolve) => {
        const worksheets = jspreadsheet(
            root,
            Object.assign(
                {
                    worksheets: [
                        { minDimensions: [3, 3], worksheetName: 'First' },
                        { minDimensions: [3, 3], worksheetName: 'Second' },
                        { minDimensions: [3, 3], worksheetName: 'Third' },
                    ],
                    onload: () => resolve(worksheets),
                },
                options
            )
        );
    });
};

const getTabs = function () {
    return Array.from(root.querySelectorAll('.jtabs-headers > :not(.jtabs-border)'));
};

const getPanels = function () {
    return Array.from(root.querySelector('.jtabs-content').children);
};

const keyCodes = { Enter: 13, ' ': 32, End: 35, Home: 36, ArrowLeft: 37, ArrowRight: 39 };

// The grid reads the key code, and the tabs read the key name
const keyDown = function (element, key) {
    const e = new KeyboardEvent('keydown', { key: key, keyCode: keyCodes[key], which: keyCodes[key], bubbles: true, cancelable: true });
    element.dispatchEvent(e);
    return e;
};

// Select a cell with the mouse, so that the grid gets the keys
const selectCell = function (worksheet, x, y) {
    const cell = worksheet.element.querySelector('td[data-x="' + x + '"][data-y="' + y + '"]');
    cell.dispatchEvent(new MouseEvent('mousedown', { bubbles: true, buttons: 1 }));
    cell.dispatchEvent(new MouseEvent('mouseup', { bubbles: true, buttons: 1 }));
};

describe('Worksheet tabs', () => {
    it('the tab list has a name, and each tab is a button with aria-selected and aria-controls', async () => {
        await create({ tabs: true });

        const tablist = root.querySelector('.jtabs-headers');
        expect(tablist.getAttribute('role')).to.eq('tablist');
        expect(tablist.getAttribute('aria-label')).to.be.a('string').and.not.empty;

        const tabs = getTabs();
        const panels = getPanels();

        expect(tabs.length).to.eq(3);
        expect(tabs.map((tab) => tab.textContent)).to.eql(['First', 'Second', 'Third']);

        tabs.forEach((tab, i) => {
            expect(tab.tagName).to.eq('BUTTON');
            expect(tab.getAttribute('type')).to.eq('button');
            expect(tab.getAttribute('role')).to.eq('tab');
            expect(tab.getAttribute('aria-selected')).to.eq(i === 0 ? 'true' : 'false');
            expect(tab.getAttribute('aria-controls')).to.eq(panels[i].id);
            expect(panels[i].getAttribute('role')).to.eq('tabpanel');
            expect(panels[i].getAttribute('aria-labelledby')).to.eq(tab.id);
        });
    });

    it('only the selected tab is in the tab sequence', async () => {
        const worksheets = await create({ tabs: true });

        let tabs = getTabs();
        expect(tabs.map((tab) => tab.getAttribute('tabindex'))).to.eql(['0', '-1', '-1']);

        worksheets[0].openWorksheet(2);

        tabs = getTabs();
        expect(tabs.map((tab) => tab.getAttribute('tabindex'))).to.eql(['-1', '-1', '0']);
        expect(tabs.map((tab) => tab.getAttribute('aria-selected'))).to.eql(['false', 'false', 'true']);
    });

    it('the arrow keys, Home and End move the focus and open the tab', async () => {
        const worksheets = await create({ tabs: true });

        const tabs = getTabs();
        const panels = getPanels();

        tabs[0].focus();

        expect(keyDown(tabs[0], 'ArrowRight').defaultPrevented).to.eq(true);
        expect(document.activeElement).to.eq(tabs[1]);
        expect(worksheets[0].getWorksheetActive()).to.eq(1);
        expect(panels[1].classList.contains('jtabs-selected')).to.eq(true);
        expect(tabs[1].getAttribute('aria-selected')).to.eq('true');

        keyDown(tabs[1], 'End');
        expect(document.activeElement).to.eq(tabs[2]);
        expect(worksheets[0].getWorksheetActive()).to.eq(2);

        // Wraps from the last tab to the first
        keyDown(tabs[2], 'ArrowRight');
        expect(document.activeElement).to.eq(tabs[0]);
        expect(worksheets[0].getWorksheetActive()).to.eq(0);

        // Wraps from the first tab to the last
        keyDown(tabs[0], 'ArrowLeft');
        expect(document.activeElement).to.eq(tabs[2]);
        expect(worksheets[0].getWorksheetActive()).to.eq(2);

        keyDown(tabs[2], 'Home');
        expect(document.activeElement).to.eq(tabs[0]);
        expect(worksheets[0].getWorksheetActive()).to.eq(0);
        expect(tabs.map((tab) => tab.getAttribute('tabindex'))).to.eql(['0', '-1', '-1']);
    });

    it('the arrow keys on a tab do not move the selected cell', async () => {
        const worksheets = await create({ tabs: true });

        selectCell(worksheets[0], 0, 0);
        expect(worksheets[0].selection.range.map(Number)).to.eql([0, 0, 0, 0]);

        const tabs = getTabs();
        tabs[0].focus();
        keyDown(tabs[0], 'ArrowRight');

        expect(worksheets[0].getWorksheetActive()).to.eq(1);
        expect(worksheets[0].selection.range.map(Number)).to.eql([0, 0, 0, 0]);
    });

    it('the button to add a worksheet is a button with a name, and Enter and Space reach it', async () => {
        const worksheets = await create({ tabs: true });

        const add = root.querySelector('.jtabs-add');
        expect(add.tagName).to.eq('BUTTON');
        expect(add.getAttribute('type')).to.eq('button');
        expect(add.getAttribute('aria-label')).to.be.a('string').and.not.empty;

        // The grid must not take Enter and Space, because the browser clicks the button on these keys
        selectCell(worksheets[0], 0, 0);
        add.focus();
        expect(keyDown(add, 'Enter').defaultPrevented).to.eq(false);
        expect(keyDown(add, ' ').defaultPrevented).to.eq(false);
        expect(worksheets[0].selection.range.map(Number)).to.eql([0, 0, 0, 0]);
        expect(worksheets.length).to.eq(3);

        add.click();
        // The new worksheet is built after an await in buildWorksheet
        await wait(10);

        const tabs = getTabs();
        expect(worksheets.length).to.eq(4);
        expect(tabs.length).to.eq(4);
        expect(tabs[3].tagName).to.eq('BUTTON');
        expect(tabs[3].getAttribute('aria-selected')).to.eq('true');
        expect(tabs[3].getAttribute('tabindex')).to.eq('0');
        expect(worksheets[0].getWorksheetActive()).to.eq(3);
    });

    it('hideTabs hides the tabs with two or more worksheets, and openWorksheet still works', async () => {
        const worksheets = await create({ tabs: true, hideTabs: true });

        const container = root.querySelector('.jtabs-headers-container');
        expect(container.style.display).to.eq('none');
        expect(getPanels().length).to.eq(3);

        worksheets[0].openWorksheet(1);

        expect(worksheets[0].getWorksheetActive()).to.eq(1);
        expect(getPanels()[1].classList.contains('jtabs-selected')).to.eq(true);
        expect(getPanels()[0].classList.contains('jtabs-selected')).to.eq(false);
        expect(container.style.display).to.eq('none');
    });

    it('shows no tabs for one worksheet without the option tabs, and shows them for two worksheets', async () => {
        await create({ worksheets: [{ minDimensions: [3, 3] }] });

        expect(root.querySelector('.jtabs-headers-container').style.display).to.eq('none');

        root.innerHTML = '';

        await create({ worksheets: [{ minDimensions: [3, 3] }, { minDimensions: [3, 3] }] });

        expect(root.querySelector('.jtabs-headers-container').style.display).to.eq('');
        expect(root.querySelector('.jtabs-add')).to.eq(null);
    });

    it('the object element.tabs keeps the methods and properties of jSuites.tabs', async () => {
        const changes = [];
        const worksheets = await create({ tabs: true });
        const tabs = root.tabs;

        expect(tabs.headers).to.eq(root.querySelector('.jtabs-headers'));
        expect(tabs.content).to.eq(root.querySelector('.jtabs-content'));
        expect(tabs.headers.lastElementChild.classList.contains('jtabs-border')).to.eq(true);
        expect(tabs.content.getAttribute('role')).to.eq('region');
        expect(tabs.content.getAttribute('aria-label')).to.eq('Tab Panels');

        const methods = ['open', 'selectIndex', 'rename', 'create', 'remove', 'nextNumber', 'deleteElement', 'appendElement'];
        methods.push('getActive', 'updateContent', 'updatePosition', 'move', 'setBorder', 'init');
        methods.forEach((method) => {
            expect(tabs[method], method).to.be.a('function');
        });

        const onchange = tabs.options.onchange;
        tabs.options.onchange = function (el, instance, index, header, content) {
            changes.push([index, header.textContent, content === getPanels()[index]]);
            return onchange.apply(this, arguments);
        };

        tabs.open(1);
        expect(tabs.getActive()).to.eq(1);
        expect(changes).to.eql([[1, 'Second', true]]);

        expect(tabs.selectIndex(getTabs()[2])).to.eq(2);
        expect(tabs.getActive()).to.eq(2);

        // create adds a worksheet with the name of the new tab
        const panel = tabs.create('Extra');
        // The new worksheet is built after an await in buildWorksheet
        await wait(10);
        expect(panel).to.eq(getPanels()[3]);
        expect(worksheets.length).to.eq(4);
        expect(worksheets[3].options.worksheetName).to.eq('Extra');
        expect(worksheets[3].element).to.eq(panel);
        expect(tabs.getActive()).to.eq(3);

        // remove and deleteElement remove the tab and its panel
        tabs.remove(3);
        expect(getTabs().length).to.eq(3);
        expect(getPanels().length).to.eq(3);
        expect(tabs.getActive()).to.eq(0);

        tabs.deleteElement(2);
        expect(getTabs().map((tab) => tab.textContent)).to.eql(['First', 'Second']);
        expect(getPanels().length).to.eq(2);
    });

    it('the methods of element.tabs are on its prototype', async () => {
        await create({ tabs: true });
        const tabs = root.tabs;

        expect(Object.prototype.hasOwnProperty.call(tabs, 'open')).to.eq(false);
        expect(Object.getPrototypeOf(tabs).open).to.be.a('function');

        // A method assigned to the object replaces it for the calls of the tabs too, as with jSuites.tabs
        const opened = [];
        const open = tabs.open;
        tabs.open = function (index) {
            opened.push(index);
            return open.call(this, index);
        };

        getTabs()[1].dispatchEvent(new MouseEvent('click', { bubbles: true }));
        expect(opened).to.eql([1]);
        expect(tabs.getActive()).to.eq(1);

        tabs.selectIndex(getTabs()[2]);
        expect(opened).to.eql([1, 2]);
        expect(tabs.getActive()).to.eq(2);
    });

    it('createWorksheet, openWorksheet and deleteWorksheet update the tabs', async () => {
        const worksheets = await create({ tabs: true });

        worksheets[0].createWorksheet({ minDimensions: [2, 2], worksheetName: 'Fourth' });
        // The new worksheet is built after an await in buildWorksheet
        await wait(10);

        let tabs = getTabs();
        expect(tabs.map((tab) => tab.textContent)).to.eql(['First', 'Second', 'Third', 'Fourth']);
        expect(tabs[3].getAttribute('aria-selected')).to.eq('true');
        expect(tabs[3].getAttribute('aria-controls')).to.eq(worksheets[3].element.id);

        worksheets[0].openWorksheet(1);
        expect(worksheets[0].getWorksheetActive()).to.eq(1);

        worksheets[0].deleteWorksheet(1);

        tabs = getTabs();
        expect(tabs.map((tab) => tab.textContent)).to.eql(['First', 'Third', 'Fourth']);
        expect(tabs.map((tab) => tab.getAttribute('aria-selected'))).to.eql(['true', 'false', 'false']);
        expect(tabs.map((tab) => tab.getAttribute('tabindex'))).to.eql(['0', '-1', '-1']);
        expect(worksheets.length).to.eq(3);
    });

    it('the context menu of a tab has the role tabs and the index of the tab', async () => {
        let args = null;

        await create({
            tabs: true,
            contextMenu: function (worksheet, x, y, e, items, role) {
                args = [x, role];
                return false;
            },
        });

        const tabs = getTabs();
        tabs[2].dispatchEvent(new MouseEvent('contextmenu', { bubbles: true, cancelable: true }));

        expect(args).to.eql([2, 'tabs']);
        expect(root.tabs.getActive()).to.eq(2);
    });
});
