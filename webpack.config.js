const path = require("path");
const CopyWebpackPlugin = require("copy-webpack-plugin");
const { CleanWebpackPlugin } = require("clean-webpack-plugin");
const HtmlWebpackPlugin = require("html-webpack-plugin");
const devCerts = require("office-addin-dev-certs");

const localhost = false;
const domain = localhost
    ? "http://127.0.0.1:8000/"
    : "https://microsoft.github.io/data-science-editor/";
module.exports = async () => ({
    mode: "development",
    entry: "./src/index.ts",
    module: {
        rules: [
            {
                test: /\.tsx?$/,
                use: "ts-loader",
                exclude: /node-modules/,
            },
        ],
    },
    resolve: {
        extensions: [".tsx", ".ts", ".js"],
    },
    output: {
        filename: "bundle.js",
        path: path.resolve(__dirname, "dist"),
    },
    plugins: [
        new CleanWebpackPlugin(),
        new HtmlWebpackPlugin({
            title: "Data Science Editor",
            template: "index.ejs",
            domain,
        }),
        new CopyWebpackPlugin({
            patterns: [
                { from: "listing/about.md", to: "" },
                { from: "hosted_files/*", to: "" },
                { from: "assets/*.png", to: "" },
            ],
        }),
    ],
    devServer: {
        headers: {
            "Access-Control-Allow-Origin": "*",
        },
        https: async () => {
            // wrapped in a function so this is only called when running the dev-server
            return await getHttpsOptions();
        },
        port: 8080,
    },
});

async function getHttpsOptions() {
    const options = await devCerts.getHttpsServerOptions();

    return {
        cacert: options.ca,
        key: options.key,
        cert: options.cert,
    };
}
