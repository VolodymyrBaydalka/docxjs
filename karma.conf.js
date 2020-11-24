module.exports = (config) => {
  config.set({
    basePath: '',
    frameworks: ['jasmine'],
    files: [
      'https://unpkg.com/jszip/dist/jszip.js',
      'https://unpkg.com/diff',
      'dist/docx-preview.js',
      'tests/**/*spec.js',
      { pattern: 'tests/**/*.docx', included: false },
      { pattern: 'tests/**/*.html', included: false }
    ],
    reporters: ['progress'],
    port: 9876,
    colors: true,
    logLevel: config.LOG_INFO,
    autoWatch: true,
    browsers: ['Chrome'],
    singleRun: false,
    concurrency: Infinity,
    crossOriginAttribute: false
  })
}
