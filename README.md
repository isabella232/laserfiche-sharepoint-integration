# laserfiche-sharepoint-integration

## Summary

This project, built with React, contains 3 Sharepoint WebParts that can be used to communicate with Laserfiche. To learn more about webparts, consult Microsoft's documentation for [Using them](https://support.microsoft.com/en-us/office/using-web-parts-on-sharepoint-pages-336e8e92-3e2d-4298-ae01-d404bbe751e0) and [Building them](https://learn.microsoft.com/en-us/sharepoint/dev/spfx/web-parts/get-started/build-a-hello-world-web-part).

## Prerequisites

See .github/workflows/main.yml for Node and NPM version used.

## Change Log

See CHANGELOG [here](./CHANGELOG.md).

## Contribution

We welcome contributions and feedback. Please follow our [contributing guidelines](./CONTRIBUTING.md).

---

## one-time setup
- clone this repo
- run **npm install**

## To run locally
- Ensure that you are at the solution folder
  - run **npm run gulp-trust-dev-cert**
  - Replace `REPLACE_WITH_YOUR_SHAREPOINT_SITE` in serve.json with your sharepoint site
  - run **npm run serve**
    - this should open up a window in the browser called a SharePoint workbench. 
  - To use a.clouddev.laserfiche.com: Open browser dev tools and go to site Local Storage: set 'spDevMode' to true

## To build solution for development/testing changes
- **npm run build**
- **npm run package**
- this should result in the creation of a file with the path `/sharepoint/solution/laserfiche-sharepoint-integration.sppkg` from the root folder.
- reference the [Admin Documentation](https://laserfiche.github.io/laserfiche-sharepoint-integration/) for instructions on how to use the solution file to test your changes to the WebParts in SharePoint Sites.

## To build solution for production
- **npm run build --ship**
- **npm run package --ship**
- Once you've built the solution file, we don't quite know how to distribute it to customers... yet.
