'use strict';

// The version that the publish workflow publishes: the major and minor version in package.json, and the patch after
// the last patch of that minor version on npm. Run it with node to print the version.

const { execFileSync } = require('child_process');

/**
 * The next version of a minor version.
 * @param packageVersion - The version in package.json, such as 6.0.0. Only its major and minor version count.
 * @param published - The versions on npm. Prereleases do not count.
 */
function nextVersion(packageVersion, published) {
    const match = /^(\d+)\.(\d+)\.\d+/.exec(packageVersion);
    if (!match) {
        throw new Error(`The version in package.json is not a version: ${packageVersion}`);
    }
    const major = Number(match[1]);
    const minor = Number(match[2]);

    let patch = -1;
    for (const version of published) {
        const parts = /^(\d+)\.(\d+)\.(\d+)$/.exec(version);
        if (parts && Number(parts[1]) === major && Number(parts[2]) === minor) {
            patch = Math.max(patch, Number(parts[3]));
        }
    }

    return `${major}.${minor}.${patch + 1}`;
}

/**
 * The versions of a package on npm, or none when it was never published.
 * @param name - The name of the package.
 */
function publishedVersions(name) {
    try {
        const output = execFileSync('npm', ['view', name, 'versions', '--json'], { encoding: 'utf8', stdio: ['ignore', 'pipe', 'pipe'] });
        // npm gives a string when there is one version
        const versions = JSON.parse(output || '[]');
        return Array.isArray(versions) ? versions : [versions];
    } catch (error) {
        if (/E404/.test(String(error.stderr))) {
            return [];
        }
        throw error;
    }
}

if (require.main === module) {
    const { name, version } = require('../../package.json');
    console.log(nextVersion(version, publishedVersions(name)));
}

module.exports = { nextVersion };
