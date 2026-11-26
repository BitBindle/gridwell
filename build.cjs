module.exports = function (source, map) {
    let result = source.replace(/import\s+jSuites\s+from\s+(?:'|")jsuites(?:'|");?/gm, '');

    result = result.replace(/import\s+formula\s+from\s+(?:'|")@jspreadsheet\/formula(?:'|");?/gm, '');

    // The import lines become empty lines, so the source map of the coverage build still fits
    this.callback(null, result, map);
};
