import { expect } from 'chai';

import jspreadsheet, { type Worksheet } from './support/jspreadsheet';

// The tests run in Node, from the folder of the repository, so they read its files
const fs = require('fs') as {
    readdirSync(path: string): string[];
    readFileSync(path: string, encoding: 'utf8'): string;
    statSync(path: string): { isDirectory(): boolean };
};

const folder = 'resources/lang';

/** The translations of a file in resources/lang, by English text. */
const read = (name: string) => JSON.parse(fs.readFileSync(`${folder}/${name}`, 'utf8')) as Record<string, string>;

const files = fs.readdirSync(folder).filter((name) => name.endsWith('.json'));
const english = read('en_GB.json');

/** The texts of the modules under src that the library translates with translate('...'). */
const translatedTexts = () => {
    const texts = new Set<string>();
    const visit = (path: string) => {
        for (const name of fs.readdirSync(path)) {
            const file = `${path}/${name}`;
            if (fs.statSync(file).isDirectory()) {
                visit(file);
            } else if (name.endsWith('.ts')) {
                for (const match of fs.readFileSync(file, 'utf8').matchAll(/\btranslate\('((?:[^'\\]|\\.)*)'\)/g)) {
                    texts.add(match[1].replace(/\\(.)/g, '$1'));
                }
            }
        }
    };
    visit('src');
    return texts;
};

describe('The translations in resources/lang (issue #153)', () => {
    it('en_GB.json has each text that the library translates, as itself', () => {
        for (const text of translatedTexts()) {
            expect(english, text).to.have.property(text);
        }
        for (const [text, translation] of Object.entries(english)) {
            expect(translation).to.equal(text);
        }
    });

    it('each file translates the texts of en_GB.json, and only those', () => {
        expect(files).to.include.members(['de_DE.json', 'fr_FR.json', 'zh_CN.json']);
        for (const name of files) {
            const translations = read(name);
            expect(Object.keys(translations), name).to.eql(Object.keys(english));
            for (const [text, translation] of Object.entries(translations)) {
                expect(translation, `${name}: ${text}`).to.be.a('string').and.not.equal('');
            }
        }
    });

    describe('A file given to setDictionary', () => {
        afterEach(() => {
            // An empty translation shows the English text again
            jspreadsheet.setDictionary(Object.fromEntries(Object.keys(english).map((text) => [text, ''])));
        });

        it('translates the texts of the spreadsheet', async () => {
            jspreadsheet.setDictionary(read('de_DE.json'));
            const worksheet = await new Promise<Worksheet>((resolve) => {
                const instance = jspreadsheet(root, {
                    worksheets: [{ minDimensions: [2, 6], search: true, pagination: 2, paginationOptions: [2, 5] }],
                    onload: () => resolve(instance[0]),
                });
            });

            const filter = root.querySelector('.jss_filter')!;
            expect(filter.querySelector('label')!.textContent).to.equal('Suchen: ');
            expect(filter.children[0].textContent).to.equal('Zeige 25 Einträge');
            expect(worksheet.pagination.children[0].textContent).to.equal('Seite 1 von 3');
            expect(worksheet.headerContainer.children[0].getAttribute('aria-label')).to.equal('Alles auswählen');

            worksheet.headers[0].dispatchEvent(new window.MouseEvent('contextmenu', { bubbles: true, cancelable: true }));
            const titles = [...root.querySelectorAll('.jss_contextmenu > div:not(.header) a')].map((item) => item.textContent);
            expect(titles.slice(0, 2)).to.eql(['Eine neue Spalte davor einfügen', 'Eine neue Spalte danach einfügen']);
            expect(titles).to.include('Kopieren...');
        });
    });
});
