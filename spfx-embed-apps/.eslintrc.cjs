/* Repo-wide ESLint config for TS + React + SPFx */
module.exports = {
  root: true,
  parser: '@typescript-eslint/parser',
  plugins: ['@typescript-eslint', 'react', 'react-hooks', 'jsx-a11y', 'import'],
  extends: [
    'eslint:recommended',
    'plugin:@typescript-eslint/recommended',
    'plugin:react/recommended',
    'plugin:react-hooks/recommended',
    'plugin:jsx-a11y/recommended',
    'prettier'
  ],
  settings: {
    react: { version: 'detect' }
  },
  parserOptions: {
    ecmaVersion: 2021,
    sourceType: 'module',
    ecmaFeatures: { jsx: true }
  },
  env: { browser: true, es2021: true, node: true, jest: true },
  rules: {
    'no-console': ['warn', { allow: ['warn', 'error'] }],
    'import/no-restricted-paths': [
      'error',
      {
        zones: [
          {
            target: './apps',
            from: './apps',
            except: [],
            message: 'Cross-app imports are not allowed. Use libs/*.'
          }
        ]
      }
    ],
    '@typescript-eslint/no-explicit-any': ['warn', { ignoreRestArgs: true }],
    '@typescript-eslint/explicit-function-return-type': 'warn',
    // A11y stricter checks
    'jsx-a11y/tabindex-no-positive': 'error',
    'jsx-a11y/anchor-is-valid': 'error',
    'jsx-a11y/no-static-element-interactions': 'warn',
    'jsx-a11y/click-events-have-key-events': 'error'
  },
  ignorePatterns: ['dist', 'lib', 'temp', 'coverage', '*.config.cjs']
};


