/** @type {import('jest').Config} */
module.exports = {
  preset: 'ts-jest',
  testEnvironment: 'node',
  roots: ['<rootDir>/src'],
  testMatch: ['**/__tests__/**/*.test.ts'],
  moduleFileExtensions: ['ts', 'js', 'json'],
  transform: {
    '^.+\\.(ts)$': [
      'ts-jest',
      {
        tsconfig: '<rootDir>/tsconfig.jest.json',
        isolatedModules: true,
      },
    ],
  },
  moduleNameMapper: {
    '^@pnp/sp/.*$': '<rootDir>/src/__mocks__/pnp-empty.js',
    '^@pnp/sp$': '<rootDir>/src/__mocks__/pnp-empty.js',
  },
  collectCoverageFrom: [
    'src/**/*.ts',
    '!src/**/__tests__/**',
    '!src/index.ts',
    '!src/spfiFactory.ts'
  ],
  coverageThreshold: {
    global: {
      lines: 80,
      statements: 80
    }
  }
};


