const devCerts = require('office-addin-dev-certs');
const { CleanWebpackPlugin } = require('clean-webpack-plugin');
const CopyWebpackPlugin = require('copy-webpack-plugin');
const ExtractTextPlugin = require('extract-text-webpack-plugin');
const HtmlWebpackPlugin = require('html-webpack-plugin');
const webpack = require('webpack');

const urlDev = 'https://localhost:3000/';
const urlProd = 'https://icelam.github.io/excel-named-range-wizard/';

module.exports = async (env, options) => {
  const dev = options.mode === 'development';
  const buildType = dev ? 'dev' : 'prod';
  const config = {
    devtool: buildType === 'dev' ? 'source-map' : false,
    entry: {
      vendor: [
        'react',
        'react-dom',
        'core-js',
        'office-ui-fabric-react',
      ],
      taskpane: [
        'react-hot-loader/patch',
        './src/taskpane/index.tsx',
      ],
      commands: './src/commands/commands.ts',
    },
    output: {
      filename: 'assets/js/[name].[hash].js',
      chunkFilename: 'assets/js/[name].[hash].chunk.js',
    },
    resolve: {
      extensions: ['.ts', '.tsx', '.html', '.js'],
    },
    module: {
      rules: [
        {
          test: /\.tsx?$/,
          use: [
            'react-hot-loader/webpack',
            'ts-loader',
          ],
          exclude: /node_modules/,
        },
        {
          test: /\.css$/,
          use: ExtractTextPlugin.extract({
            fallback: 'style-loader',
            use: 'css-loader',
          }),
        },
        {
          test: /\.(png|jpg|jpeg|gif)$/,
          loader: 'file-loader',
          options: {
            name: '[path][name].[ext]',
          },
        },
      ],
    },
    plugins: [
      new CleanWebpackPlugin(),
      new CopyWebpackPlugin({
        patterns: [
          {
            to: `[name].${buildType}.[ext]`,
            from: 'manifest*.xml',
            transform(content) {
              if (dev) {
                return content;
              }
              return content.toString().replace(new RegExp(urlDev, 'g'), urlProd);
            },
          },
          {
            to: 'assets',
            from: 'assets',
          },
        ],
      }),
      new ExtractTextPlugin('[name].[hash].css'),
      new HtmlWebpackPlugin({
        filename: 'taskpane.html',
        template: './src/taskpane/taskpane.html',
        chunks: ['taskpane', 'vendor', 'polyfills'],
      }),
      new HtmlWebpackPlugin({
        filename: 'commands.html',
        template: './src/commands/commands.html',
        chunks: ['commands'],
      }),
      new webpack.ProvidePlugin({
        Promise: ['es6-promise', 'Promise'],
      }),
    ],
    devServer: {
      hot: true,
      headers: {
        'Access-Control-Allow-Origin': '*',
      },
      https: (options.https !== undefined) ? options.https : await devCerts.getHttpsServerOptions(),
      port: process.env.npm_package_config_dev_server_port || 3000,
    },
  };

  return config;
};
