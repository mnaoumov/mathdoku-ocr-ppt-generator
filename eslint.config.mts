// ESLint configuration -- based on obsidian-dev-utils strict config,
// Adapted for Google Apps Script (no modules, global scope).
// Omitted from obsidian-dev-utils: import-x, modules-newlines, obsidianmd (not applicable).

/* eslint-disable no-magic-numbers -- Config files use magic numbers for rule settings. */

import commentsConfigs from '@eslint-community/eslint-plugin-eslint-comments/configs';
import eslint from '@eslint/js';
import stylistic from '@stylistic/eslint-plugin';
import perfectionist from 'eslint-plugin-perfectionist';
import {
  defineConfig,
  globalIgnores
} from 'eslint/config';
import tseslint from 'typescript-eslint';

export default defineConfig(
  globalIgnores([
    '**/*.js',
    '**/node_modules/'
  ]),
  // ── Base configs ──
  eslint.configs.recommended,
  ...tseslint.configs.strictTypeChecked,
  ...tseslint.configs.stylisticTypeChecked,
  commentsConfigs.recommended,
  perfectionist.configs['recommended-alphabetical'],
  // ── Stylistic ──
  stylistic.configs.recommended,
  stylistic.configs.customize({
    arrowParens: true,
    braceStyle: '1tbs',
    commaDangle: 'never',
    semi: true
  }),
  // ── Project rules ──
  {
    languageOptions: {
      parserOptions: {
        projectService: true,
        tsconfigRootDir: import.meta.dirname
      }
    },
    plugins: {
      '@stylistic': stylistic
    },
    rules: {
      // ── Comments ──
      '@eslint-community/eslint-comments/require-description': 'error',
      // ── Stylistic overrides ──
      '@stylistic/indent': 'off',
      '@stylistic/indent-binary-ops': 'off',
      '@stylistic/jsx-one-expression-per-line': 'off',
      '@stylistic/no-extra-semi': 'error',
      '@stylistic/object-curly-newline': [
        'error',
        {
          ExportDeclaration: {
            minProperties: 2,
            multiline: true
          },
          ImportDeclaration: {
            minProperties: 2,
            multiline: true
          }
        }
      ],
      '@stylistic/operator-linebreak': [
        'error',
        'before',
        { overrides: { '=': 'after' } }
      ],
      '@stylistic/quotes': [
        'error',
        'single',
        { allowTemplateLiterals: 'never' }
      ],
      // ── TypeScript-ESLint ──
      '@typescript-eslint/explicit-function-return-type': 'error',
      '@typescript-eslint/explicit-member-accessibility': 'error',
      '@typescript-eslint/no-invalid-void-type': ['error', {
        allowAsThisParameter: true
      }],
      '@typescript-eslint/no-non-null-assertion': 'off',
      '@typescript-eslint/no-this-alias': ['error', {
        allowedNames: ['that']
      }],
      // Disabled: TypeScript handles via noUnusedLocals/noUnusedParameters.
      // ESLint can't see cross-file usage in Apps Script's global scope (module: "None").
      '@typescript-eslint/no-unused-vars': 'off',
      '@typescript-eslint/prefer-nullish-coalescing': [
        'error',
        { ignorePrimitives: { boolean: true } }
      ],
      '@typescript-eslint/prefer-readonly': 'error',
      // ── Core ESLint ──
      'accessor-pairs': 'error',
      'array-callback-return': 'error',
      'camelcase': 'error',
      'capitalized-comments': ['error', 'always', { block: { ignorePattern: 'v8' } }],
      'complexity': 'error',
      'consistent-this': 'error',
      'curly': 'error',
      'default-case': 'error',
      'default-case-last': 'error',
      'default-param-last': 'error',
      'eqeqeq': 'error',
      'func-name-matching': 'error',
      'func-names': 'error',
      'func-style': [
        'error',
        'declaration',
        { allowArrowFunctions: false }
      ],
      'grouped-accessor-pairs': ['error', 'getBeforeSet'],
      'guard-for-in': 'error',
      'no-alert': 'error',
      'no-array-constructor': 'error',
      'no-bitwise': 'error',
      'no-caller': 'error',
      'no-console': ['error', { allow: ['warn', 'error'] }],
      'no-constructor-return': 'error',
      'no-div-regex': 'error',
      'no-else-return': ['error', { allowElseIf: false }],
      'no-empty-function': 'error',
      'no-extend-native': 'error',
      'no-extra-bind': 'error',
      'no-extra-label': 'error',
      'no-implicit-coercion': ['error', { allow: ['!!'] }],
      'no-implied-eval': 'error',
      'no-inner-declarations': 'error',
      'no-iterator': 'error',
      'no-label-var': 'error',
      'no-labels': 'error',
      'no-lone-blocks': 'error',
      'no-lonely-if': 'error',
      'no-loop-func': 'error',
      'no-magic-numbers': [
        'error',
        {
          detectObjects: true,
          enforceConst: true,
          ignore: [-1, 0, 1]
        }
      ],
      'no-multi-assign': 'error',
      'no-multi-str': 'error',
      'no-negated-condition': 'error',
      'no-nested-ternary': 'error',
      'no-new-func': 'error',
      'no-new-wrappers': 'error',
      'no-object-constructor': 'error',
      'no-octal-escape': 'error',
      'no-promise-executor-return': 'error',
      'no-proto': 'error',
      'no-restricted-syntax': [
        'error',
        {
          message: 'Do not use definite assignment assertions (!). Initialize the field or make it optional.',
          selector: 'PropertyDefinition[definite=true]'
        },
        {
          message: 'Do not use definite assignment assertions (!) on abstract fields.',
          selector: 'TSAbstractPropertyDefinition[definite=true]'
        }
      ],
      'no-return-assign': 'error',
      'no-script-url': 'error',
      'no-self-compare': 'error',
      'no-sequences': 'error',
      'no-shadow': 'error',
      'no-template-curly-in-string': 'error',
      'no-throw-literal': 'error',
      'no-unmodified-loop-condition': 'error',
      'no-unneeded-ternary': 'error',
      'no-unreachable-loop': 'error',
      'no-unused-expressions': 'error',
      'no-useless-assignment': 'error',
      'no-useless-call': 'error',
      'no-useless-computed-key': 'error',
      'no-useless-concat': 'error',
      'no-useless-constructor': 'error',
      'no-useless-rename': 'error',
      'no-useless-return': 'error',
      'no-var': 'error',
      'no-void': 'error',
      'object-shorthand': 'error',
      'operator-assignment': 'error',
      'prefer-arrow-callback': 'error',
      'prefer-const': 'error',
      'prefer-exponentiation-operator': 'error',
      'prefer-named-capture-group': 'error',
      'prefer-numeric-literals': 'error',
      'prefer-object-has-own': 'error',
      'prefer-object-spread': 'error',
      'prefer-promise-reject-errors': 'error',
      'prefer-regex-literals': 'error',
      'prefer-rest-params': 'error',
      'prefer-spread': 'error',
      'prefer-template': 'error',
      'radix': 'error',
      'require-atomic-updates': 'error',
      'require-await': 'error',
      'symbol-description': 'error',
      'unicode-bom': 'error',
      'vars-on-top': 'error',
      'yoda': 'error'
    }
  }
);
/* eslint-enable no-magic-numbers -- end config file block. */
