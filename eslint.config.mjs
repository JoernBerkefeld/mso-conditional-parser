import eslintPluginPrettierRecommended from 'eslint-plugin-prettier/recommended';
import eslintPluginUnicorn from 'eslint-plugin-unicorn';
import globals from 'globals';
import jsdoc from 'eslint-plugin-jsdoc';
import js from '@eslint/js';

export default [
    {
        ignores: ['**/node_modules/**', '**/dist/**', '**/out/**', '**/coverage/**'],
    },
    js.configs.recommended,
    eslintPluginPrettierRecommended,
    jsdoc.configs['flat/recommended'],
    eslintPluginUnicorn.configs['recommended'],
    {
        languageOptions: {
            globals: {
                ...globals.nodeBuiltin,
                Atomics: 'readonly',
                SharedArrayBuffer: 'readonly',
            },
            ecmaVersion: 2022,
            sourceType: 'module',
        },
        settings: {
            jsdoc: {
                mode: 'typescript',
                preferredTypes: {
                    array: 'Array',
                    'array.<>': '[]',
                    'Array.<>': '[]',
                    'array<>': '[]',
                    'Array<>': '[]',
                    Object: 'object',
                    'object.<>': 'Object.<>',
                    'object<>': 'Object.<>',
                    'Object<>': 'Object.<>',
                },
            },
        },
        rules: {
            'logical-assignment-operators': ['error', 'always'],
            'unicorn/no-null': 'off',
            'unicorn/prevent-abbreviations': 'off',
            'arrow-body-style': ['error', 'as-needed'],
            curly: 'error',
            'no-console': 'warn',
            'jsdoc/check-line-alignment': 2,
            'jsdoc/require-jsdoc': 'off',
            'jsdoc/require-param-type': 'off',
            'jsdoc/require-returns-type': 'off',
            'jsdoc/tag-lines': ['warn', 'any', { startLines: 1 }],
            'jsdoc/no-undefined-types': 'off',
            'jsdoc/valid-types': 'off',
            'spaced-comment': [
                'warn',
                'always',
                {
                    block: {
                        exceptions: ['*'],
                        balanced: true,
                    },
                },
            ],
        },
    },
    {
        files: ['**/*.{js,mjs,cjs}'],
        rules: {
            'no-var': 'error',
            'prefer-const': 'error',
            'prettier/prettier': 'warn',
            'prefer-arrow-callback': 'warn',
            'no-unused-vars': ['error', { argsIgnorePattern: '^_', varsIgnorePattern: '^_' }],
        },
    },
    {
        files: ['tests/**/*.mjs'],
        rules: {
            'no-console': 'off',
        },
    },
];
