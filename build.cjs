module.exports = function (source, map) {
    const result = source.replace(/import\s+jSuites\s+from\s+(?:'|")jsuites(?:'|");?/gm, '');

    // The import line becomes an empty line, so the source map of the coverage build still fits
    this.callback(null, result, map);
};
