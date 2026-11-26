import { expect } from 'chai';

import gridwell, { type Worksheet } from './support/gridwell';

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

/**
 * The texts of the modules under src that the library translates with translate('...') or translate('...', '...'), each
 * with its English text. A text with a context is in the dictionary as "text (context)".
 */
const translatedTexts = () => {
    const texts = new Map<string, string>();
    const unescape = (text: string) => text.replace(/\\(.)/g, '$1');
    const visit = (path: string) => {
        for (const name of fs.readdirSync(path)) {
            const file = `${path}/${name}`;
            if (fs.statSync(file).isDirectory()) {
                visit(file);
            } else if (name.endsWith('.ts')) {
                const source = fs.readFileSync(file, 'utf8');
                for (const match of source.matchAll(/\btranslate\('((?:[^'\\]|\\.)*)'(?:, '((?:[^'\\]|\\.)*)')?\)/g)) {
                    const text = unescape(match[1]);
                    texts.set(match[2] === undefined ? text : `${text} (${unescape(match[2])})`, text);
                }
            }
        }
    };
    visit('src');
    return texts;
};

/** The names of the months and the days, full and short, which the dates translate from lists. */
const dateNames = () => {
    const names = new Map<string, string>();
    for (const name of ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December']) {
        names.set(name, name);
    }
    for (const name of ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec']) {
        // The short name of May is its full name in English
        names.set(name === 'May' ? 'May (short)' : name, name);
    }
    for (const name of ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday', 'Sun', 'Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat']) {
        names.set(name, name);
    }
    return names;
};

describe('The translations in resources/lang (issue #153)', () => {
    it('en_GB.json has each text that the library translates, as itself', () => {
        const texts = new Map([...translatedTexts(), ...dateNames()]);
        for (const [key, text] of texts) {
            expect(english, key).to.have.property(key, text);
        }
        for (const [key, translation] of Object.entries(english)) {
            expect(translation, key).to.equal(texts.get(key) ?? key);
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
            gridwell.setDictionary(Object.fromEntries(Object.keys(english).map((text) => [text, ''])));
        });

        it('translates the texts of the spreadsheet', async () => {
            gridwell.setDictionary(read('de_DE.json'));
            const worksheet = await new Promise<Worksheet>((resolve) => {
                const instance = gridwell(root, {
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
