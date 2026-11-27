import { expect } from 'chai';

import jspreadsheet from './support/jspreadsheet';

describe('Version', () => {
    it('gives the version in package.json', () => {
        const { version } = require('../package.json') as { version: string };

        expect(jspreadsheet.version().version).to.equal(version);
        expect(jspreadsheet.version().print()[0]).to.equal(['Gridwell', version, 'https://github.com/BitBindle/gridwell', 'MIT'].join('\r\n'));
    });
});
