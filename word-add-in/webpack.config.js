/* eslint-disable no-undef */

// const devCerts = require("office-addin-dev-certs");
const CopyWebpackPlugin = require("copy-webpack-plugin");
const HtmlWebpackPlugin = require("html-webpack-plugin");
const webpack = require("webpack");
// const fs = require("fs");
// const path = require("path");

const urlDev = "http://localhost:3002/";
const urlProd = "https://www.anygenai.com/"; // CHANGE THIS TO YOUR PRODUCTION DEPLOYMENT LOCATION

// async function getHttpsOptions() {
//   // Check if custom certificates exist
//   const certPath = path.resolve(__dirname, "certs/anygen.ai.pem");
//   const keyPath = path.resolve(__dirname, "certs/domain.key");
//   const caPath = path.resolve(__dirname, "certs/SectigoRSADomainValidationSecureServerCA.pem");

//   if (fs.existsSync(certPath) && fs.existsSync(keyPath)) {
//     console.log("Using custom certificates from certs folder:");
//     console.log(`- Certificate: ${certPath}`);
//     console.log(`- Key: ${keyPath}`);

//     const options = {
//       cert: fs.readFileSync(certPath),
//       key: fs.readFileSync(keyPath),
//     };

//     // Add CA certificate if it exists
//     if (fs.existsSync(caPath)) {
//       console.log(`- CA Certificate: ${caPath}`);
//       options.ca = fs.readFileSync(caPath);
//     }

//     return options;
//   } else {
//     console.log("Custom certificates not found, falling back to dev certs");
//     const httpsOptions = await devCerts.getHttpsServerOptions();
//     return { ca: httpsOptions.ca, key: httpsOptions.key, cert: httpsOptions.cert };
//   }
// }

module.exports = async (env, options) => {
  const dev = options.mode === "development";
  const config = {
    devtool: "source-map",
    entry: {
      polyfill: ["core-js/stable", "regenerator-runtime/runtime"],
      vendor: ["react", "react-dom", "core-js", "@fluentui/react-components", "@fluentui/react-icons"],
      taskpane: ["./src/taskpane/index.jsx", "./src/taskpane/taskpane.html"],
      commands: "./src/commands/commands.js",
    },
    output: {
      clean: true,
    },
    resolve: {
      extensions: [".js", ".jsx", ".html"],
    },
    module: {
      rules: [
        {
          test: /\.jsx?$/,
          use: {
            loader: "babel-loader",
          },
          exclude: /node_modules/,
        },
        {
          test: /\.html$/,
          exclude: /node_modules/,
          use: "html-loader",
        },
        {
          test: /\.(png|jpg|jpeg|ttf|woff|woff2|gif|ico)$/,
          type: "asset/resource",
          generator: {
            filename: "assets/[name][ext][query]",
          },
        },
      ],
    },
    plugins: [
      new HtmlWebpackPlugin({
        filename: "taskpane.html",
        template: "./src/taskpane/taskpane.html",
        chunks: ["polyfill", "vendor", "taskpane"],
      }),
      new CopyWebpackPlugin({
        patterns: [
          {
            from: "assets/*",
            to: "assets/[name][ext][query]",
          },
          {
            from: "manifest*.xml",
            to: "[name]" + "[ext]",
            transform(content) {
              if (dev) {
                return content;
              } else {
                return content.toString().replace(new RegExp(urlDev, "g"), urlProd);
              }
            },
          },
        ],
      }),
      new HtmlWebpackPlugin({
        filename: "commands.html",
        template: "./src/commands/commands.html",
        chunks: ["polyfill", "commands"],
      }),
      new webpack.ProvidePlugin({
        Promise: ["es6-promise", "Promise"],
      }),
    ],
    devServer: {
      hot: true,
      headers: {
        "Access-Control-Allow-Origin": "*",
      },
      // server: {
      //   type: "https",
      //   options: env.WEBPACK_BUILD || options.https !== undefined ? options.https : await getHttpsOptions(),
      // },
      port: process.env.npm_package_config_dev_server_port || 3000,
    },
  };

  return config;
};
