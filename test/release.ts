import { expect } from 'chai';

// The script of the publish workflow
const { nextVersion } = require('../.github/scripts/next-version.cjs') as { nextVersion: (packageVersion: string, published: string[]) => string };

describe('The version that the publish workflow publishes', () => {
    it('starts a minor version on patch 0', () => {
        expect(nextVersion('6.0.0', [])).to.equal('6.0.0');
        expect(nextVersion('6.1.0', ['6.0.0', '6.0.1'])).to.equal('6.1.0');
    });

    it('gives the patch after the last patch of the minor version', () => {
        expect(nextVersion('6.0.0', ['6.0.0', '6.0.2', '6.0.1', '6.1.0', '7.0.4'])).to.equal('6.0.3');
        expect(nextVersion('6.0.0', ['6.0.9', '6.0.10'])).to.equal('6.0.11');
    });

    it('takes only the major and minor version from package.json', () => {
        expect(nextVersion('6.2.7', ['6.2.0'])).to.equal('6.2.1');
    });

    it('does not count the prereleases and the other packages', () => {
        expect(nextVersion('6.0.0', ['6.0.0-beta.1', '5.0.4', '60.0.1'])).to.equal('6.0.0');
    });

    it('fails on a version that is not a version', () => {
        expect(() => nextVersion('next', [])).to.throw('not a version');
    });
});
