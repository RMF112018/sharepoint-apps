/** @type {import('jest').Config} */
module.exports = {
  testEnvironment: 'jsdom',
  setupFilesAfterEnv: ['<rootDir>/jest.setup.cjs'],
  roots: ['<rootDir>/src'],
  testMatch: ['**/__tests__/**/*.test.ts?(x)'],
  moduleFileExtensions: ['ts', 'tsx', 'js', 'jsx'],
  transform: {
    '^.+\\.[tj]sx?$': [
      'ts-jest',
      {
        tsconfig: '<rootDir>/tsconfig.json',
        isolatedModules: true,
      },
    ],
  },
  moduleNameMapper: {
    '\\.(css|scss)$': '<rootDir>/styleMock.cjs',
    '\\.(png|jpg|jpeg|gif|svg)$': '<rootDir>/fileMock.cjs'
  },
  collectCoverageFrom: [
    'src/webparts/**/components/**/*.tsx',
    '!src/**/__tests__/**'
  ],
  coverageThreshold: {
    global: {
      lines: 60,
      statements: 60,
    },
  },
};


