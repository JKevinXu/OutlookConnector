/* eslint-disable no-undef */

const devCerts = require("office-addin-dev-certs");
const CopyWebpackPlugin = require("copy-webpack-plugin");
const HtmlWebpackPlugin = require("html-webpack-plugin");

const urlDev = "https://localhost:3000/";
const urlProd = "https://jkevinxu.github.io/OutlookConnector/"; // Updated for GitHub Pages

async function getHttpsOptions() {
  const httpsOptions = await devCerts.getHttpsServerOptions();
  return { ca: httpsOptions.ca, key: httpsOptions.key, cert: httpsOptions.cert };
}

module.exports = async (env, options) => {
  const dev = options.mode === "development";
  const publicPath = dev ? "/" : "/OutlookConnector/";
  
  const config = {
    devtool: "source-map",
    entry: {
      polyfill: ["core-js/stable", "regenerator-runtime/runtime"],
      taskpane: ["./src/taskpane/taskpane.ts", "./src/taskpane/taskpane.html"],
      commands: "./src/commands/commands.ts",
    },
    output: {
      clean: true,
      publicPath: publicPath,
    },
    resolve: {
      extensions: [".ts", ".html", ".js"],
    },
    module: {
      rules: [
        {
          test: /\.ts$/,
          exclude: /node_modules/,
          use: {
            loader: "babel-loader"
          },
        },
        {
          test: /\.html$/,
          exclude: /node_modules/,
          use: "html-loader",
        },
        {
          test: /\.(png|jpg|jpeg|gif|ico)$/,
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
        chunks: ["polyfill", "taskpane"],
        publicPath: publicPath,
      }),
      new CopyWebpackPlugin({
        patterns: [
          {
            from: "assets/*",
            to: "assets/[name][ext][query]",
          },
          {
            from: "manifest*.json",
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
        publicPath: publicPath,
      }),
    ],
    devServer: {
      headers: {
        "Access-Control-Allow-Origin": "*",
        "Content-Security-Policy": "frame-ancestors 'self' https://localhost:* http://localhost:* https://*.amazon.com https://*.office.com https://*.office365.com https://*.microsoftonline.com https://*.officeapps.live.com https://*.outlook.com https://outlook.live.com https://outlook.office.com https://outlook.office365.com; default-src 'self' https://localhost:* http://localhost:* https://*.amazon.com https://*.office.com https://*.office365.com https://appsforoffice.microsoft.com; script-src 'self' 'unsafe-inline' 'unsafe-eval' https://localhost:* http://localhost:* https://appsforoffice.microsoft.com; style-src 'self' 'unsafe-inline' https://localhost:* http://localhost:* https://res-1.cdn.office.net; connect-src 'self' https://localhost:* http://localhost:* https://*.amazon.com https://*.microsoft.com https://*.office.com https://*.office365.com https://*.microsoftonline.com https://graph.microsoft.com https://login.microsoftonline.com https://appsforoffice.microsoft.com;"
      },
      server: {
        type: "https",
        options: env.WEBPACK_BUILD || options.https !== undefined ? options.https : await getHttpsOptions(),
      },
      port: process.env.npm_package_config_dev_server_port || 3000,
      historyApiFallback: {
        index: '/taskpane.html',
        rewrites: [
          { from: /^\/taskpane\/callback$/, to: '/taskpane.html' },
        ]
      },
      proxy: [
        {
          context: ['/api'],
          target: 'https://bwzo9wnhy3.execute-api.us-west-2.amazonaws.com/beta',
          secure: true,
          changeOrigin: true,
          logLevel: 'debug',
          pathRewrite: {
            '^/api': ''
          },
          onProxyReq: (proxyReq, req, res) => {
            console.log('Proxying request:', req.method, req.url);
          },
          onProxyRes: (proxyRes, req, res) => {
            console.log('Proxy response:', proxyRes.statusCode, req.url);
          }
        }
      ]
    },
  };

  return config;
};
