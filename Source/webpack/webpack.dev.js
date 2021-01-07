const path = require("path");
const webpack = require("webpack");
const { merge } = require('webpack-merge');
const commonConfig = require('./webpack.common');

module.exports = merge(commonConfig, {
    mode: 'development',

    // Enable sourcemaps for debugging webpack's output.
    devtool: 'eval-cheap-module-source-map',

    // Dev server options
    devServer: {
        contentBase: path.resolve(__dirname, '../wwwroot'),
        port: 8080,
        historyApiFallback: true,
        hot: true,
        inline: true,
        writeToDisk: true,
    },

    plugins: [
        new webpack.HotModuleReplacementPlugin(),
    ],
});