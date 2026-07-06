const path = require("path");
const webpack = require("webpack");
const HtmlWebpackPlugin = require("html-webpack-plugin");
const CopyWebpackPlugin = require("copy-webpack-plugin");
const CustomFunctionsMetadataPlugin = require("custom-functions-metadata-plugin");
const { CleanWebpackPlugin } = require("clean-webpack-plugin");

const packageJson = require("./package.json");
const isProduction = process.env.NODE_ENV === "production";
const buildTimestamp = new Date().toISOString();
// Lowercase so OAuth redirect_uri matches Authentik strict redirect URIs.
const addinHost = (process.env.ADDIN_HOST || "https://localhost:3000")
  .replace(/\/+$/, "")
  .toLowerCase();

const versionManifest = JSON.stringify({
  version: packageJson.version,
  buildTimestamp,
  minimumSupported: packageJson.minimumSupportedVersion || packageJson.version,
  releaseNotesUrl: packageJson.homepage || "",
}, null, 2);

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
    extensionAlias: {
      ".js": [".js", ".ts"],
    },
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
    new webpack.DefinePlugin({
      __ADDIN_VERSION__: JSON.stringify(packageJson.version),
      __ADDIN_NAME__: JSON.stringify(packageJson.name),
      __BUILD_TIMESTAMP__: JSON.stringify(buildTimestamp),
      __ADDIN_HOST__: JSON.stringify(addinHost),
    }),
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
        { from: "docs", to: "docs" },
        {
          from: "manifest.xml",
          to: "manifest.xml",
          transform(content) {
            return content
              .toString()
              .replace(/https:\/\/localhost:3000/g, addinHost);
          },
        },
      ],
    }),
    {
      apply(compiler) {
        compiler.hooks.thisCompilation.tap("EmitVersionJson", (compilation) => {
          compilation.hooks.processAssets.tap(
            { name: "EmitVersionJson", stage: webpack.Compilation.PROCESS_ASSETS_STAGE_ADDITIONAL },
            () => {
              compilation.emitAsset("version.json", new webpack.sources.RawSource(versionManifest));
            }
          );
        });
      },
    },
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
