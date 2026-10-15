// @ts-check
import path from 'node:path';
import { fileURLToPath } from 'node:url';

import { FlatCompat } from '@eslint/eslintrc';
import eslint from '@eslint/js';
import { defineConfig } from 'eslint/config';
import eslintConfigPrettier from 'eslint-config-prettier';
import eslintPluginPrettierRecommended from 'eslint-plugin-prettier/recommended';
import pluginVue from 'eslint-plugin-vue';
import globals from 'globals';
import tseslint from 'typescript-eslint';
import vueParser from 'vue-eslint-parser';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

const compat = new FlatCompat({
  baseDirectory: __dirname,
  recommendedConfig: eslint.configs.recommended,
});

const airbnbBase = compat.extends('airbnb-base');

export default defineConfig(
  {
    ignores: [
      '**/dist/**',
      '**/node_modules/**',
      '**/coverage/**',
      '**/*.config.{js,cjs,mjs,ts}',
      'apps/api/prisma/migrations/**',
      'apps/web/public/**',
      'data/**',
    ],
  },
  eslint.configs.recommended,
  ...airbnbBase,
  ...tseslint.configs.recommended,
  {
    languageOptions: {
      ecmaVersion: 'latest',
      sourceType: 'module',
    },
    settings: {
      'import/resolver': {
        typescript: {
          alwaysTryTypes: true,
          noWarnOnMultipleProjects: true,
          project: [
            'apps/api/tsconfig.json',
            'apps/web/tsconfig.json',
            'apps/web/tsconfig.app.json',
            'packages/shared/tsconfig.json',
          ],
        },
        node: true,
      },
      'import/parsers': {
        '@typescript-eslint/parser': ['.ts', '.tsx'],
        'vue-eslint-parser': ['.vue'],
      },
      'import/extensions': ['.js', '.mjs', '.cjs', '.ts', '.tsx', '.vue'],
      // Avoid TS parser trying to load SFC when resolving imports from .ts
      'import/ignore': ['\\.(coffee|scss|css|less|hbs|svg|json|vue)$'],
    },
    rules: {
      'import/prefer-default-export': 'off',
      'import/extensions': 'off',
      'import/no-extraneous-dependencies': [
        'error',
        {
          devDependencies: [
            '**/*.{test,spec}.{ts,js}',
            '**/test/**',
            '**/tests/**',
            '**/vite.config.*',
            '**/eslint.config.*',
          ],
        },
      ],
      'class-methods-use-this': 'off',
      'no-underscore-dangle': [
        'error',
        { allow: ['_id', '__dirname', '__filename', '__v'] },
      ],
      '@typescript-eslint/no-explicit-any': 'off',
      '@typescript-eslint/no-unused-vars': [
        'error',
        {
          argsIgnorePattern: '^_',
          varsIgnorePattern: '^_',
          caughtErrorsIgnorePattern: '^_',
        },
      ],
      'no-unused-vars': 'off',
      'no-shadow': 'off',
      '@typescript-eslint/no-shadow': 'error',
      'no-use-before-define': 'off',
      '@typescript-eslint/no-use-before-define': [
        'error',
        { functions: false, classes: true, variables: true },
      ],
    },
  },
  {
    files: ['apps/api/**/*.{ts,js}'],
    languageOptions: {
      globals: {
        ...globals.node,
        ...globals.jest,
      },
      parserOptions: {
        projectService: {
          allowDefaultProject: ['*.js', '*.mjs', '*.cjs'],
        },
        tsconfigRootDir: path.join(__dirname, 'apps/api'),
      },
    },
  },
  {
    files: ['packages/shared/**/*.{ts,js}'],
    languageOptions: {
      globals: {
        ...globals.node,
      },
      parserOptions: {
        projectService: {
          allowDefaultProject: ['*.js', '*.mjs', '*.cjs'],
        },
        tsconfigRootDir: path.join(__dirname, 'packages/shared'),
      },
    },
  },
  {
    files: ['packages/shared/scripts/**/*.{mjs,js}'],
    rules: {
      'no-console': 'off',
    },
  },
  {
    files: ['apps/web/**/*.{ts,vue}'],
    languageOptions: {
      globals: {
        ...globals.browser,
      },
    },
  },
  {
    files: ['**/*.{ts,tsx}'],
    languageOptions: {
      parser: tseslint.parser,
      parserOptions: {
        ecmaVersion: 'latest',
        sourceType: 'module',
      },
    },
  },
  ...pluginVue.configs['flat/recommended'].map((config) => ({
    ...config,
    files: ['apps/web/**/*.vue'],
  })),
  {
    files: ['apps/web/**/*.vue'],
    languageOptions: {
      parser: vueParser,
      parserOptions: {
        parser: tseslint.parser,
        ecmaVersion: 'latest',
        sourceType: 'module',
        extraFileExtensions: ['.vue'],
        // Airbnb sets globalReturn:true which breaks vue-eslint-parser v-for scope
        ecmaFeatures: {
          globalReturn: false,
        },
      },
      globals: {
        ...globals.browser,
      },
    },
    rules: {
      'vue/multi-word-component-names': 'off',
      'vue/require-default-prop': 'off',
      // script-setup symbols used only in template are invisible to TS unused-vars
      '@typescript-eslint/no-unused-vars': 'off',
      'no-unused-vars': 'off',
    },
  },
  // Nest/Vue: allow fire-and-forget in event handlers (`void promise`)
  {
    files: ['apps/web/**/*.{ts,vue}'],
    rules: {
      'no-void': 'off',
      'no-alert': 'off',
      'no-console': ['warn', { allow: ['warn', 'error'] }],
      'vue/no-v-html': 'off',
    },
  },
  // Web: non-type-checked (Vue SFC + projectService breaks vue-eslint-parser)
  // Prettier last — disables conflicting Airbnb/Vue stylistic rules
  eslintConfigPrettier,
  eslintPluginPrettierRecommended,
  {
    rules: {
      'prettier/prettier': ['error', { endOfLine: 'auto' }],
    },
  },
  // Ensure Vue parser wins after prettier/other merges
  {
    files: ['apps/web/**/*.vue'],
    languageOptions: {
      parser: vueParser,
      parserOptions: {
        parser: tseslint.parser,
        ecmaVersion: 'latest',
        sourceType: 'module',
        extraFileExtensions: ['.vue'],
        ecmaFeatures: {
          globalReturn: false,
        },
      },
    },
  },
);
