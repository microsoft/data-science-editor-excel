const path = require("path");
const { CleanWebpackPlugin } = require("clean-webpack-plugin");
const HtmlWebpackPlugin = require("html-webpack-plugin");
const devCerts = require("office-addin-dev-certs");

async function getHttpsOptions() {
    const options = await devCerts.getHttpsServerOptions();
    
    return {
        cacert: options.ca,
        key: options.key,
        cert: options.cert
    };
}

module.exports = async () => ({
    mode: "development",
    entry: "./src/index.ts",
    module: {
        rules: [
            {
                test: /\.tsx?$/,
                use: "ts-loader",
                exclude: /node-modules/
            }
        ]
    },
    resolve: {
        extensions: [".tsx", ".ts", ".js"]
    },
    output: {
        filename: "bundle.js",
        path: path.resolve(__dirname, "dist")
    },
    plugins: [
        new CleanWebpackPlugin(),
        new HtmlWebpackPlugin({
          title: "Data Science Editor",
          template: "index.ejs"
        }),
    ],
    devServer: {
        headers: {
            "Access-Control-Allow-Origin": "*"
        },
        https: await getHttpsOptions(),
        port: 8080
    }
});
