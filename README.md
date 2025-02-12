<!--Copyright (c) Laserfiche.
Licensed under the MIT License. See LICENSE.md in the project root for license information.-->

# laserfiche-sharepoint-integration

## Summary

This project, built with React, contains 3 SharePoint web parts and a command set that can be used to communicate with Laserfiche. To learn more about web parts, consult Microsoft's documentation for [using them](https://support.microsoft.com/en-us/office/using-web-parts-on-sharepoint-pages-336e8e92-3e2d-4298-ae01-d404bbe751e0) and [building them](https://learn.microsoft.com/en-us/sharepoint/dev/spfx/web-parts/get-started/build-a-hello-world-web-part).

Project contains:

- 3 Web Parts (UI Components that are placed on SharePoint pages)
  - [Laserfiche Administrator Configuration](./src/webparts/laserficheAdminConfiguration/)
  - [Laserfiche Repository Explorer](./src/webparts/LaserficheRepositoryAccessWebPart/)
  - [Laserfiche Sign In](./src/webparts/sendToLaserficheLoginComponent/)
- 1 Command Set (Action available on items in lists and libraries)
  - [Send to Laserfiche](./src/extensions/savetoLaserfiche/)

Admin and User Documentation is available on [GitHub pages](https://laserfiche.github.io/laserfiche-sharepoint-integration/)

## Prerequisites

See .github/workflows/main.yml for Node and NPM version used.

## Change Log

See CHANGELOG [here](./jekyll_files/CHANGELOG.md).

## Contribution

We welcome contributions and feedback. Please follow our [contributing guidelines](./CONTRIBUTING.md).

---

## To run locally in your SharePoint Workbench

- Ensure that you are at the solution folder
  - run **npm install**
  - run **npm run gulp-trust-dev-cert**
  - Replace `REPLACE_WITH_YOUR_SHAREPOINT_SITE` in serve.json with your sharepoint site
  - run **npm run serve**
    - this should open up a window in the browser called a SharePoint workbench.
  - To use a.clouddev.laserfiche.com: Open browser dev tools and go to site Local Storage: set 'spDevMode' to true

## To test in a SharePoint site using localhost

- run **npm install**
- **npm run build**
- **npm run package**
- this should result in the creation of a file with the path `/sharepoint/solution/laserfiche-sharepoint-integration.sppkg` from the root folder.
- Navigate to the solution folder
- run **npm run gulp-trust-dev-cert** (one-time only)
- Replace `REPLACE_WITH_YOUR_SHAREPOINT_SITE` in serve.json with your sharepoint site
- run **npm run serve** to host the code for the integration
- reference the [Admin Documentation](https://laserfiche.github.io/laserfiche-sharepoint-integration/docs/admin-documentation) for instructions on how to use the solution file you generated in the first steps to test your changes to the web parts in SharePoint sites.

## To test in a SharePoint site with files hosted by SharePoint

- run **npm install**
- **npm run bundle**
- **npm run package-solution**
- This will create the solution file at /sharepoint/solution/laserfiche-sharepoint-integration.sppkg.
- Once you've built and packaged the solution file, you can use it as a production package and upload it in the SharePoint admin center (see [Admin Docs](https://laserfiche.github.io/laserfiche-sharepoint-integration/docs/admin-documentation))

## To build documentation locally

1. Follow the instructions [here](https://jekyllrb.com/docs/) to install jekyll and bundler.
   - if any dependency installation fails, try downloading the package manually.
   - navigate to the download directory.
   - run `gem install problematic-dependency-name`, where the last word is replaced by the package you downloaded manually.
   - retry installing the original package now that its dependency is installed.
1. Run `bundle install` to install all the dependencies needed to serve.
1. Navigate to the `jekyll_files` directory.
1. Run `bundle exec jekyll serve` to serve the documentation.
1. Open `localhost:4000` in a browser.
