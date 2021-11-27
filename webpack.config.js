const path = require('path')

const PATHS = {
  src: path.join(__dirname, './src'),
  build: path.join(__dirname, './dist')
}

function buildConfig(prod, es6) {
  const outputFilename = `[name]${es6 ? '.es6' : ''}${prod ? '.min' : ''}.js`;
  const tsLoaderOptions = es6 ? { compilerOptions: { target: "es6" } } : {};

  return {
    mode: 'development',
    entry: {
      'docx-preview': PATHS.src + '/docx-preview.ts'
    },
    output: {
      path: PATHS.build,
      filename: outputFilename,
      library: 'docx',
      libraryTarget: 'umd'
    },
    devtool: 'source-map',
    module: {
      rules: [
        {
          test: /\.ts$/,
          use: [{
            loader: 'ts-loader',
            options: tsLoaderOptions
          }]
        }
      ]
    },
    resolve: {
      extensions: ['.ts', '.js']
    },
    externals: {
      "jszip": "JSZip",
    }
  }
}


module.exports = (env, argv) => {
  return buildConfig(argv.mode === 'production', false);
};
