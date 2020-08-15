export default {
  extraBabelPlugins: [],
  cssModules: {
    generateScopedName: '[name]-[local]',
  },
  injectCSS: false,
  lessInBabelMode: false,
  esm: 'rollup',
  cjs: 'rollup',
};
