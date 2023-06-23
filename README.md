# laserfiche-sharepoint-integration

## Summary

This project, built with React, contains 3 Sharepoint WebParts that can be used to communicate with Laserfiche.

## Prerequisites

See .github/workflows/main.yml for Node and NPM version used.

## Change Log

See CHANGELOG [here](./CHANGELOG.md).

## Contribution

We welcome contributions and feedback. Please follow our [contributing guidelines](./CONTRIBUTING.md).

---

## To run locally

- Clone this repository
- Ensure that you are at the solution folder
  - **npm install**
  - **npm run gulp-trust-dev-cert**
  - Replace `REPLACE_WITH_YOUR_SHAREPOINT_SITE` in serve.json with your sharepoint site
  - **npm run serve**
  - To use a.clouddev.laserfiche.com: Open browser dev tools and go to site Local Storage: set 'spDevMode' to true
