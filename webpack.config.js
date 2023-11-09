const path = require('path')

function buildConfig(prod, umd = false) {
  const config = {
    mode: 'development',
    entry: {
      'docx-preview': path.join(__dirname, './src/docx-preview.ts')
    },
    output: {
      path: path.join(__dirname, './dist'),
      filename: `[name]${prod ? '.min' : ''}.${umd ? '' : 'm'}js`,
      globalObject: 'globalThis'
    },
    devtool: 'source-map',
    module: {
      rules: [
        {
          test: /\.ts$/,
          use: [{ loader: 'ts-loader' }]
        }
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
        amd: "jszip",
        module: 'jszip'
      },
    }
  };

  if (umd) {
    config.output.library = { name: 'docx', type: 'umd', umdNamedDefine: true };
  } else {
    config.experiments = { outputModule: true };
    config.output.library = { type: 'module' };
  }

  return config;
}


module.exports = (env, argv) => {
  return buildConfig(argv.mode === 'production', env.umd);
};
