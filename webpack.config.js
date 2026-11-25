const path = require('path');
const MiniCssExtractPlugin = require('mini-css-extract-plugin');
const CopyPlugin = require('copy-webpack-plugin');

class MyPlugin {
    apply(compiler) {
        compiler.hooks.emit.tap('MyPlugin', (compilation) => {
            // Get the bundled file name
            const fileName = 'index.js';

            // Get the bundled file content
            const fileContent = compilation.assets[fileName].source();

            const header = `if (! jSuites && typeof(require) === 'function') {
    var jSuites = require('jsuites');
}

if (! formula && typeof(require) === 'function') {
    var formula = require('@jspreadsheet/formula');
}

;(function (global, factory) {
    typeof exports === 'object' && typeof module !== 'undefined' ? module.exports = factory() :
    typeof define === 'function' && define.amd ? define(factory) :
    global.jspreadsheet = factory();
}(this, (function () {`;

            const footer = `    return jspreadsheet;
})));`;

            // Updated file content with custom content added
            const updatedFileContent = header + '\n\n' + fileContent + '\n\n' + footer;

            // Replace the bundled file content with updated content
            compilation.assets[fileName] = {
                source: () => updatedFileContent,
                size: () => updatedFileContent.length,
            };
        });
    }
}

let isProduction = process.env.NODE_ENV === 'production';

const webpack = {
    target: ['web', 'es5'],
    entry: isProduction ? './src/index' : './src/test.js',
    mode: isProduction ? 'production' : 'development',
    externals: {},
    output: {
        filename: 'index.js',
        path: path.resolve(__dirname, 'dist'),
        library: 'jspreadsheet',
        libraryExport: 'default',
    },
    optimization: {
        minimize: true,
    },
    devServer: {
        static: {
            directory: path.join(__dirname, '/public'),
        },
        headers: {
            'Access-Control-Allow-Origin': '*',
            'Access-Control-Allow-Methods': 'GET, POST, PUT, DELETE, PATCH, OPTIONS',
            'Access-Control-Allow-Headers': 'X-Requested-With, content-type, Authorization',
        },
        port: 8000,
        devMiddleware: {
            publicPath: 'https://localhost:3000/',
        },
        hot: 'only',
    },
    plugins: [],
    resolve: {
        extensions: ['.ts', '.js'],
        // The modules import each other with the .js extension, also when the file is a .ts file
        extensionAlias: {
            '.js': ['.ts', '.js'],
        },
    },
    module: {
        rules: [
            isProduction
                ? {
                      test: /\.[jt]s$/,
                      use: [
                          {
                              loader: path.resolve('build.cjs'),
                              options: {},
                          },
                      ],
                  }
                : null,
            // Remove the types. TypeScript checks them separately (npm run typecheck).
            {
                test: /\.ts$/,
                loader: 'esbuild-loader',
                options: {
                    loader: 'ts',
                    // Keep the syntax as it is, as for the .js files
                    target: 'esnext',
                },
            },
            {
                test: /\.css$/,
                use: [isProduction ? MiniCssExtractPlugin.loader : 'style-loader', 'css-loader'],
            },
        ],
    },
    stats: {
        warnings: false,
    },
};

if (isProduction) {
    webpack.plugins.push(new MyPlugin());
    // The type definitions and stylesheets are maintained by hand in src and shipped next to the bundle
    webpack.plugins.push(
        new CopyPlugin({
            patterns: [
                { from: 'src/index.d.ts', to: 'index.d.ts' },
                { from: 'src/jspreadsheet.css', to: 'jspreadsheet.css' },
                { from: 'src/jspreadsheet.themes.css', to: 'jspreadsheet.themes.css' },
            ],
        })
    );
    webpack.plugins.push(
        new MiniCssExtractPlugin({
            filename: 'jspreadsheet.css',
        })
    );
}

module.exports = webpack;
