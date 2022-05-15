const path = require('path')

const PATHS = {
  src: path.join(__dirname, './src'),
  build: path.join(__dirname, './dist')
}

function buildConfig(prod) {
  const outputFilename = `[name]${prod ? '.min' : ''}.js`;

  return {
    mode: 'development',
    entry: {
      'docx-preview': PATHS.src + '/docx-preview.ts'
    },
    output: {
      path: PATHS.build,
      filename: outputFilename,
      library: 'docx',
      libraryTarget: 'umd',
      globalObject: 'globalThis'
    },
    devtool: 'source-map',
    module: {
      rules: [
        {
          test: /\.ts$/,
          use: [{
            loader: 'ts-loader'
          }]
        },
        {
          test: /\.s[ac]ss$/i,
          use: [
            { loader: "css-loader", options: { exportType: "string" } },
            { loader: "sass-loader" },
          ],
        },
      ]
    },
    resolve: {
      extensions: ['.ts', '.js']
    },
    externals: {
      "jszip": {
        root: "JSZip",
        commonjs: "jszip",
        commonjs2: "jszip",
        amd: "jszip"
      },
    }
  }
}


module.exports = (env, argv) => {
  return buildConfig(argv.mode === 'production');
};
