# Data Science Editor

Contains prototype of Data Science Editor Office.js add-in.

## Getting Started

1. Download the latest LTS version of [node.js](https://nodejs.org/en/download/).
1. Install all dependencies.

```bash
npm install
```

## Build

The following script will build and place assets in the dist directory:

```bash
npm build
```

## Lint

Runs prettier over all typescript files

```back
npm lint
```

## Manual Test

1. Run the following script to start the dev server:
    - `npm run dev-server`
1. [Manually sideload the add-in to Office on the web](https://learn.microsoft.com/en-us/office/dev/add-ins/testing/sideload-office-add-ins-for-testing#manually-sideload-an-add-in-to-office-on-the-web)
1. select `manifest.xml`

## Layout

Layout of folders

- __src__
    - source code for the add-in
- __listing__
    - descriptions for the add-in store listing
    - __statements__
    - statements that must be linked to the add-in store listing

## Architecture

This add-in is a host to an iframe that holds the data science editor. This Add-In provides the interface to allow the data-science editor to interact with Excel.

This add-in is complete static and hosted on a github pages site.

## Hosted Site

A pipeline is configured to build and deploy to an Azure Static Web App site on each merge into `main`.  The pipeline file is `pipelines\build_and_deploy_to_static_app.yaml`.

### Changing the Azure Static Web App Resource to which the pipeline deploys

You can create a new Azure Static Web App through the Azure portal.  Once the resource is created, on the Overview page, click on `Manage deployment token` and copy the value.  Next, go to [edit the deployment pipeline](https://office.visualstudio.com/Hackathon/_apps/hub/ms.vss-build-web.ci-designer-hub?pipelineId=17281&branch=main).  Click on the `Variables` button, then click on the `deployment_token` variable, replace the value, and click `Ok`.

### Changing the Domain

 A custom URL is generated when you create the Azure Static Web App resource and different domains can be associated through the [Azure portal for the resource](https://ms.portal.azure.com/?websitesextension_ext=asd.featurePath%3Danalysis%2FLinuxAppDown#@microsoft.onmicrosoft.com/resource/subscriptions/55163ab1-9de3-49af-888e-c06f7e2c3a49/resourceGroups/excel-data-science-editor-addin/providers/Microsoft.Web/staticSites/excel-data-science-editor-addin-app/customDomains).  If you change the domain, you will need to update the URLs in `hosted_files\hosted_manifest.xml`.

## TODOs

- [ ] detect changes in worksheet and notify blocks to recompute
- [ ] somehow 1 block workspace per worksheet (low pri)
- [ ] fix the editor CSS so that it uses the whole screen
- [ ] match the color scheme looks to Excel?
