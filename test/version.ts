import { expect } from 'chai';

import gridwell from './support/gridwell';

describe('Version', () => {
    it('gives the version in package.json', () => {
        const { version } = require('../package.json') as { version: string };

        expect(gridwell.version().version).to.equal(version);
        expect(gridwell.version().print()[0]).to.equal(['Gridwell', version, 'https://github.com/BitBindle/jspreadsheet-ce', 'MIT'].join('\r\n'));
    });
});
