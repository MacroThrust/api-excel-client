const path = require("path");
const HtmlWebpackPlugin = require("html-webpack-plugin");
const CopyWebpackPlugin = require("copy-webpack-plugin");
const CustomFunctionsMetadataPlugin = require("custom-functions-metadata-plugin");
const { CleanWebpackPlugin } = require("clean-webpack-plugin");

const isProduction = process.env.NODE_ENV === "production";

module.exports = {
  entry: {
    taskpane: "./src/taskpane/taskpane.ts",
    functions: "./src/functions/functions.ts",
    commands: "./src/commands/commands.ts",
    "auth-dialog": "./src/auth/authDialog.ts",
  },
  output: {
    path: path.resolve(__dirname, "dist"),
    filename: "[name].[contenthash:8].js",
    clean: true,
  },
  resolve: {
    extensions: [".ts", ".js", ".json"],
  },
  module: {
    rules: [
      {
        test: /\.ts$/,
        use: "ts-loader",
        exclude: /node_modules/,
      },
      {
        test: /\.css$/,
        use: ["style-loader", "css-loader"],
      },
      {
        test: /\.(png|jpg|jpeg|gif|ico|svg)$/,
        type: "asset/resource",
        generator: {
          filename: "assets/[name][ext]",
        },
      },
    ],
  },
  plugins: [
    new CleanWebpackPlugin(),
    new CustomFunctionsMetadataPlugin({
      output: "functions.json",
      input: "./src/functions/functions.ts",
    }),
    new HtmlWebpackPlugin({
      filename: "taskpane.html",
      template: "./src/taskpane/taskpane.html",
      chunks: ["taskpane", "functions", "commands"],
    }),
    new HtmlWebpackPlugin({
      filename: "auth-dialog.html",
      template: "./src/auth/auth-dialog.html",
      chunks: ["auth-dialog"],
    }),
    new CopyWebpackPlugin({
      patterns: [
        { from: "assets", to: "assets" },
        { from: "manifest.xml", to: "manifest.xml" },
      ],
    }),
  ],
  devServer: {
    static: {
      directory: path.join(__dirname, "dist"),
    },
    headers: {
      "Access-Control-Allow-Origin": "*",
    },
    server: "https",
    port: 3000,
    hot: true,
  },
  devtool: isProduction ? "source-map" : "eval-source-map",
};
