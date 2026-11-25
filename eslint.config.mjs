import js from '@eslint/js';
import globals from 'globals';
import tseslint from 'typescript-eslint';

export default tseslint.config(
    js.configs.recommended,
    {
        ignores: ['dist/**', 'node_modules/**'],
    },
    {
        files: ['**/*.{js,mjs,cjs}'],
        languageOptions: {
            ecmaVersion: 2022,
            sourceType: 'module',
            globals: {
                ...globals.browser,
                ...globals.node,
                ...globals.es2022,
                jspreadsheet: 'readonly',
            },
        },
        rules: {
            'no-unused-vars': 'off',
        },
    },
    {
        files: ['**/*.ts'],
        extends: [tseslint.configs.recommended],
        languageOptions: {
            globals: {
                ...globals.browser,
                ...globals.es2022,
            },
        },
        rules: {
            // The modules use the worksheet as this, and most functions start with const obj = this
            '@typescript-eslint/no-this-alias': 'off',
            // any is allowed while the code moves to TypeScript
            '@typescript-eslint/no-explicit-any': 'off',
            '@typescript-eslint/no-unused-vars': 'off',
            // Off while the modules move to TypeScript, so that the move changes no code. Turn them on afterwards.
            'prefer-spread': 'off',
        },
    },
    {
        files: ['test/**/*.js', 'mocha.config.js'],
        languageOptions: {
            globals: {
                ...globals.node,
                ...globals.mocha,
                describe: 'readonly',
                it: 'readonly',
                before: 'readonly',
                after: 'readonly',
                beforeEach: 'readonly',
                afterEach: 'readonly',
                root: 'readonly',
            },
        },
    },
    {
        files: ['webpack.config.js', 'build.cjs'],
        languageOptions: {
            globals: {
                ...globals.node,
                __dirname: 'readonly',
                __filename: 'readonly',
                module: 'readonly',
                require: 'readonly',
                exports: 'readonly',
                process: 'readonly',
            },
        },
    }
);
