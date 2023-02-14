# Data Science Editor

Prototype integration of the [Data Science Editor](https://microsoft.github.io/data-scienc-editor/) in Office for the Web.

![A gif showcasing the app in action.](https://microsoft.github.io/data-science-editor-excel/hosted_files/editorHowTo.gif)

## Developer Zone

These instructions will show you how to build and debug the Office App locally.

### Getting Started

These instructions have been tested on Windows only.

1. Download the latest LTS version of [node.js](https://nodejs.org/en/download/).
1. Install all dependencies.

```bash
npm run install
```

### Build

The following script will build and place assets in the dist directory:

```bash
npm run build
```

### Lint

Runs prettier over all typescript files

```back
npm run lint
```

### Manual Test

1. Run the following script to start the dev server:
    - `npm run dev-server`
1. [Manually sideload the add-in to Office on the web](https://learn.microsoft.com/en-us/office/dev/add-ins/testing/sideload-office-add-ins-for-testing#manually-sideload-an-add-in-to-office-on-the-web)
1. select `manifest-local.xml`

### Layout

Layout of folders

-   **src**
    -   source code for the add-in
-   **listing**
    -   descriptions for the add-in store listing
    -   **statements**
    -   statements that must be linked to the add-in store listing

## Testing the hosted data science editor locally

-   clone https://github.com/microsoft/data-science-editor and follow instructions to launch dev server
-   update `localhost` to true in webpack.config.js and rebuild

## Architecture

This add-in is a host to an iframe that holds the data science editor. This Add-In provides the interface to allow the data-science editor to interact with Excel.

This add-in is complete static and hosted on a github pages site.

## TODOs

-   [ ] detect changes in worksheet and notify blocks to recompute
-   [ ] somehow 1 block workspace per worksheet (low pri)
-   [ ] fix the editor CSS so that it uses the whole screen
-   [ ] match the color scheme looks to Excel?
