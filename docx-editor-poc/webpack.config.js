const path = require('path')

const PATHS = {
  src: path.join(__dirname, './src'),
  build: path.join(__dirname, './build')
}

var config = {
  mode: 'development',
  entry: {
    'index': PATHS.src + '/index.ts'
  },
  output: {
    path: PATHS.build,
    filename: '[name].js',
    library: 'docx',
    libraryTarget: 'umd'
  },
  devtool: 'source-map',
  module: {
    rules: [
      {
        test: /\.ts$/,
        use: 'ts-loader'
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

module.exports = config;
