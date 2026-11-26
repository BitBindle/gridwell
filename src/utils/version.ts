// Basic version information
export default {
    version: '5.0.0',
    host: 'https://github.com/BitBindle/jspreadsheet-ce',
    license: 'MIT',
    print(): [string] {
        return [['Gridwell', this.version, this.host, this.license].join('\r\n')];
    },
};
