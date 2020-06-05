// eslint-disable-next-line no-undef
module.exports = {
  env: {
    "googleappsscript/googleappsscript": true,
    es6: true,
  },
  plugins: ["googleappsscript"],
  extends: "eslint:recommended",
  globals: {
    Atomics: "readonly",
    SharedArrayBuffer: "readonly",
  },
  parserOptions: {
    ecmaVersion: 11,
    sourceType: "module",
  },
  rules: {
    quotes: ["error", "double"],
  },
};
