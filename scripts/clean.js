/* eslint-env node */
/* eslint @typescript-eslint/no-var-requires: "off" */

// cross platform replacement for
// "clean":"rm --dir --recursive --verbose --force dist temp",
// "clean-windows": "if exist dist (rmdir /S /Q dist) && if exist temp (rmdir /S /Q temp)",

const fs = require("fs");

const parameters = process.argv.slice(2);

if (parameters.length !== 1) {
    console.log("usage: [delete directory path]");
    process.exit(1);
}

const [target] = parameters;

// check that target exists before removing it.
if (fs.existsSync(target)) {
    fs.rmSync(target, { recursive: true });
}
