import js from '@eslint/js';
import globals from 'globals';
import tseslint from 'typescript-eslint';

export default tseslint.config(
    js.configs.recommended,
    {
        ignores: ['dist/**', 'node_modules/**', 'coverage/**'],
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
            // any is allowed while the code moves to TypeScript
            '@typescript-eslint/no-explicit-any': 'off',
            '@typescript-eslint/no-unused-vars': 'off',
            // Functions are declarations, callbacks arrow functions and functions in objects methods
            'func-style': ['error', 'declaration'],
            'prefer-arrow-callback': 'error',
            'object-shorthand': ['error', 'always'],
            // A loop reads the items of an array with for...of, and a list is searched with includes
            '@typescript-eslint/prefer-for-of': 'error',
            // Strings are built with template literals
            'prefer-template': 'error',
            // Comparisons are strict, except x == null, which is true for null and undefined
            eqeqeq: ['error', 'always', { null: 'ignore' }],
        },
    },
    {
        // The rules that need the types
        files: ['src/**/*.ts', 'test/**/*.ts'],
        languageOptions: {
            parserOptions: {
                projectService: true,
                tsconfigRootDir: import.meta.dirname,
            },
        },
        rules: {
            '@typescript-eslint/prefer-includes': 'error',
            '@typescript-eslint/prefer-optional-chain': 'error',
        },
    },
    {
        files: ['test/**/*.ts'],
        languageOptions: {
            globals: {
                ...globals.mocha,
            },
        },
        rules: {
            // A test keeps its helpers in arrow functions next to the tests that use them
            'func-style': ['error', 'declaration', { allowArrowFunctions: true }],
            // The bundle is a UMD module, which a test requires
            '@typescript-eslint/no-require-imports': 'off',
            // Some chai assertions are properties, such as expect(x).to.be.ok
            '@typescript-eslint/no-unused-expressions': 'off',
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
        files: ['webpack.config.js', 'build.cjs', 'check-cycles.cjs'],
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
