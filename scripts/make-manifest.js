// create the production manifest from the local manifest.

/**
 * clean data
 * @param {string} data
 * @returns
 */
function clean(data) {
    // remove the BOM
    // https://en.wikipedia.org/wiki/Byte_order_mark
    // The BOM is generally unexpected in text files and causes JSON.parse to fail.
    // U+FEFF is the Byte Order Mark for UTF-8
    data = data.replace(/^\uFEFF/, "");

    // standardize newlines to proper unix line endings
    data = data.replace(/\r/gm, "");
    return data;
}

/**
 * make manifest for production
 * @param {string} manifest
 */
function production(data) {
    // replace
    data = data.replaceAll(
        "https://localhost:8080/",
        "https://microsoft.github.io/data-science-editor-excel/"
    );
    return clean(data);
}

function main() {
    const fs = require("fs");

    const localManifestPath = "manifest-local.xml";
    const prodManifestPath = "./hosted_files/manifest.xml";

    const data = fs.readFileSync(localManifestPath, { encoding: "utf-8" });
    fs.writeFileSync(prodManifestPath, production(data));
}

main();
