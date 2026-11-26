import { expect } from 'chai';

import gridwell from './support/gridwell';

describe('The script', () => {
    it('sets the global gridwell when there is no module loader', () => {
        const fs = require('fs') as { readFileSync: (path: string, encoding: string) => string };
        const code = fs.readFileSync('dist/index.js', 'utf8');
        const global: Record<string, unknown> = {};

        // mocha runs in the folder of the package. The script runs with this as the global object, and without module, exports and define
        new Function('module', 'exports', 'define', code).call(global);

        expect(Object.keys(global)).to.eql(['gridwell']);
        expect((global.gridwell as typeof gridwell).version().version).to.equal(gridwell.version().version);
    });
});
