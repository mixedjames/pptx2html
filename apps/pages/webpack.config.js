import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import CopyWebpackPlugin from 'copy-webpack-plugin';
import HtmlWebpackPlugin from 'html-webpack-plugin';
import webpack from 'webpack';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const demosDir = path.resolve(__dirname, 'src/demos');

// Demo presentations are just dropped into src/demos — this scan is how the picker in
// src/index.ts learns what's available, so adding a .pptx there is the only step needed.
const demoFiles = fs.existsSync(demosDir)
  ? fs
      .readdirSync(demosDir)
      .filter((name) => name.toLowerCase().endsWith('.pptx'))
      .sort()
  : [];

export default {
  entry: './src/index.ts',
  output: {
    path: path.resolve(__dirname, 'dist'),
    filename: '[name].[contenthash].js',
    // Resolves relative to the executing script's own URL, so the same build works whether it's
    // served from a GitHub Pages project subpath (/repo-name/...), a custom domain, or locally.
    publicPath: 'auto',
    clean: true,
  },
  resolve: {
    extensions: ['.ts', '.js'],
  },
  module: {
    rules: [
      {
        test: /\.ts$/,
        use: 'ts-loader',
        exclude: /node_modules/,
      },
    ],
  },
  plugins: [
    new HtmlWebpackPlugin({
      template: './src/index.html',
    }),
    new webpack.DefinePlugin({
      __DEMO_FILES__: JSON.stringify(demoFiles),
    }),
    new CopyWebpackPlugin({
      patterns: [{ from: demosDir, to: 'demos', noErrorOnMissing: true }],
    }),
  ],
  devServer: {
    static: path.resolve(__dirname, 'dist'),
    port: 8081,
  },
};
