# Data Science Editor

Contains prototype of Data Science Editor Office.js add-in.

## Getting Started

1. Download the latest LTS version of [node.js](https://nodejs.org/en/download/).
1. Install all dependencies.

> npm run install


## Build

The following script will build and place assets in the dist directory:

> npm run build


## Lint

Runs prettier over all typescript files

> npm run lint

## Manual Test

1. Run the following script to start the dev server:
    - `npm run dev-server`
1. [Manually sideload the add-in to Office on the web](https://learn.microsoft.com/en-us/office/dev/add-ins/testing/sideload-office-add-ins-for-testing#manually-sideload-an-add-in-to-office-on-the-web)
1. select `manifest-local.xml`

## Add-In Manifest

### Generate production manifest.xml from manifest-local.xml
Make all manifest changes to `manifest-local.xml`.

When the local manifest changes run:

> npm run update-manifest

The command:

- check that the local manifest is valid
- generates the production `manifest.xml`
- checks the production manifest is valid



### Requirement Set

The Manifest is set to require a specific Excel version to avoid having to support specific outdated browser versions.

[ExcelApi Requirement Sets and Supported Office Versions](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/excel/excel-api-requirement-sets#requirement-set-availability)

## Layout

Layout of folders

- __assets__
    - image assets
- __src__
    - source code for the add-in
- __listing__
    - descriptions for the add-in store listing
    - __statements__
    - statements that must be linked to the add-in store listing
- __hosted_files__
    - additional hosted files
- __scripts__
    - development scripts

- __dist__
    - the build site, this is the exact layout hosted

## Architecture

This add-in is a host to an iframe that holds the data science editor. This Add-In provides the interface to allow the data-science editor to interact with Excel.

This add-in is complete static and hosted on a github pages site.

## Testing the hosted data science editor locally

- clone https://github.com/microsoft/data-science-editor and follow instructions to launch dev server
- update `localhost` to true in webpack.config.js and rebuild

## TODOs

- [ ] detect changes in worksheet and notify blocks to recompute
- [ ] somehow 1 block workspace per worksheet (low pri)
- [ ] fix the editor CSS so that it uses the whole screen
- [ ] match the color scheme looks to Excel?
