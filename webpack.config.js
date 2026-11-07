/* eslint-disable no-undef */
const path = require("path");
const HtmlWebpackPlugin = require("html-webpack-plugin");
const CopyWebpackPlugin = require("copy-webpack-plugin");

module.exports = (env, options) => ({
  // ES5 runtime webpacku, aby launchevent.js bezel i ve starsim JS runtime klasickeho Outlooku
  target: ["web", "es5"],
  devtool: options.mode === "development" ? "source-map" : false,
  entry: {
    launchevent: "./src/launchevent.js",
    taskpane: "./src/taskpane.js"
  },
  output: {
    path: path.resolve(__dirname, "dist"),
    filename: "[name].js",
    publicPath: "",
    clean: true
  },
  module: {
    rules: [
      {
        test: /\.m?js$/,
        include: [path.resolve(__dirname, "src"), path.resolve(__dirname, "node_modules/@azure")],
        use: {
          loader: "babel-loader",
          options: {
            presets: [["@babel/preset-env", { targets: { browsers: ["since 2015"] } }]]
          }
        }
      }
    ]
  },
  plugins: [
    new HtmlWebpackPlugin({
      filename: "commands.html",
      template: "./src/commands.html",
      chunks: ["launchevent"],
      inject: "head",
      scriptLoading: "blocking"
    }),
    new HtmlWebpackPlugin({
      filename: "taskpane.html",
      template: "./src/taskpane.html",
      chunks: ["taskpane"],
      inject: "body",
      scriptLoading: "blocking"
    }),
    new CopyWebpackPlugin({
      patterns: [
        { from: "assets", to: "assets" },
        { from: ".nojekyll", to: ".nojekyll", toType: "file" }
      ]
    })
  ],
  performance: { hints: false }
});
