module.exports = {
  moduleFileExtensions: ['ts', 'tsx', 'js', 'jsx', 'json', 'node'],
  testEnvironment: 'node',
  testPathIgnorePatterns: ['/dist/', '/node_modules/'],
  testRegex: '/(src|tests)/.*\\.(test|spec)?\\.(ts|tsx)$',
  transform: {'^.+\\.ts?$': 'ts-jest'},
  verbose: true,
}
