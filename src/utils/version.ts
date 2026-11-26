// Basic version information. The build takes the version from package.json.
export default {
    version: __VERSION__,
    host: 'https://github.com/BitBindle/jspreadsheet-ce',
    license: 'MIT',
    print(): [string] {
        return [['Gridwell', this.version, this.host, this.license].join('\r\n')];
    },
};
