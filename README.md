# Data Science Editor

Prototype integration of the [Data Science Editor](https://microsoft.github.io/data-science-editor/excel/) in Office for the Web.

![A gif showcasing the app in action.](https://microsoft.github.io/data-science-editor-excel/hosted_files/editorHowTo.gif)

## Developer Zone

These instructions will show you how to build and debug the Office App locally.
To file issues, use https://github.com/microsoft/data-science-editor/issues

## Getting Started

These instructions have been tested on Windows only.

1. Download the latest LTS version of [node.js](https://nodejs.org/en/download/).
1. Install all dependencies.

> npm run install

### Build

The following script will build and place assets in the dist directory:

> npm run build

### Lint

Runs prettier over all typescript files

> npm run lint

### Manual Test

1. Run the following script to start the dev server:
    - `npm run server`
1. [Manually sideload the add-in to Office on the web](https://learn.microsoft.com/en-us/office/dev/add-ins/testing/sideload-office-add-ins-for-testing#manually-sideload-an-add-in-to-office-on-the-web)
1. select `manifest-local.xml`

## Testing the hosted data science editor locally

-   clone https://github.com/microsoft/data-science-editor and follow instructions to launch dev server
-   update `localhost` to true in webpack.config.js and rebuild


## Architecture

This add-in is a host to an iframe that holds the data science editor. This Add-In provides the interface to allow the data-science editor to interact with Excel.

This add-in is complete static and hosted on a github pages site.

### Layout

Layout of folders

- **assets**
    - image assets
- **src**
    - source code for the add-in
- **listing**
    - descriptions for the add-in store listing
- **hosted_files**
    - additional hosted files
- **scripts**
    - development scripts
- **config**
    - tooling configuration files

- **dist**
    - the build site, this is the exact layout hosted

## Add-In Manifest

### Generate production manifest.xml from manifest-local.xml

Make all manifest changes to `manifest-local.xml`.

When the local manifest changes run:

> npm run manifest

The command:

- check that the local manifest is valid
- generates the production `manifest.xml`
- checks the production manifest is valid

### Requirement Set

The Manifest is set to require a specific Excel version to avoid having to support specific outdated browser versions.

[ExcelApi Requirement Sets and Supported Office Versions](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/excel/excel-api-requirement-sets#requirement-set-availability)
