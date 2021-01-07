const { merge } = require('webpack-merge');
const commonConfig = require('./webpack.common');

module.exports = merge(commonConfig, {
    mode: 'production',

    // Enable sourcemaps for debugging webpack's output.
    devtool: 'source-map'
});