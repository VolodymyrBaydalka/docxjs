const path = require('path')

const PATHS = {
  src: path.join(__dirname, './src'),
  build: path.join(__dirname, './dist')
}

var config = {
  mode: 'development',
  entry: {
    'docx-preview': PATHS.src + '/docx-preview.ts'
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

module.exports = (env, argv) => {
  if (argv.mode === 'production') {
    config.output.filename = '[name].min.js'
  }

  return config;
};
