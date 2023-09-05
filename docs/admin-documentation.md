---
layout: default
title: Laserfiche SharePoint Integration Administration Guide
nav_order: 2
---

# Laserfiche SharePoint Integration Administration Guide

## PRE-RELEASE DOCUMENTATION - SUBJECT TO CHANGE

## Deploy Laserfiche SharePoint Integration to a SharePoint Sites

1. Prerequisites
  - A SharePoint Online account with "Site owner" permission
  - Download the Laserfiche SharePoint Integration [latest package file](./assets/laserfiche-sharepoint-integration.sppkg)
1. [Click here](https://go.microsoft.com/fwlink/?linkid=2185219) to go to the SharePoint Admin Center or find the same link at [learn.microsoft.com](https://learn.microsoft.com/en-us/sharepoint/sharepoint-admin-role#about-the-sharepoint-administrator-role-in-microsoft-365).
1. In the navigation menu, select the "More features" item.
1. Open "Apps".
1. Click Upload and select the Laserfiche SharePoint package file (.sppkg). 
  - NOTE: It's possible to build a new SharePoint Integration package file directly from source code by following instructions in [README.md](https://github.com/Laserfiche/laserfiche-sharepoint-integration#readme)
solution.
1. In your SharePoint Site (Not the Admin Center), navigate to your
site’s App catalog by clicking on the "Site Contents" item in the
navigation bar on the left side of the page.
1. Open the "New" Dropdown menu by clicking on the "+" icon.
1. Add the App named “laserfiche-sharepoint-integration-client-side-solution”.
1. Enable the app if you are asked to do so.
1. Navigate to your SharePoint site. On successful installation "Laserfiche SharePoint Integration" app is listed under the “Site Contents” tab.


## Use Laserfiche Apps on SharePoint Pages

- The Repository Explorer Page:
    1. In your SharePoint Site, select the "Pages" item in the navigation bar on the left side of the page.
    1. Create and open a new site page by clicking the blue "+ New" button and selecting "Site Page" from the dropdown.
    1. Title the page “LaserficheSpApp”.
    1. Move your cursor just below the title area to the white space beneath. This should reveal a hidden "+" button. If you hover over it, it should display the message "Add a new web part in column one”.
    1. Click on that button and Search for “Repository Explorer.
    1. Click on the search result with a white L on an orange square. The Repository Explorer WebPart should now appear on your Page. Before using the Webpart, make sure to [Register Your App in the Laserfiche Developer Console](https://laserfiche.github.io/laserfiche-sharepoint-integration/docs/admin-documentation.html#how-to-register-your-app-in-the-developer-console)
    1. After you register your App and give it the right redirect URI, you should be able to log in and use the component. For Documentation on how to use the components, reference the [User Documentation](./user-documentation.html)

- The Admin Configuration Page:
    - Follow the same steps as above, but title the page “LaserficheSpAdministration”, and add the “Admin Configuration” web part instead of “Repository Access".
- The Laserfiche Sign In Page:
    - Follow the same steps as above, but title the Page LaserficheSignIn, and add the “Laserfiche Sign In” web part.
    - NOTE: If you do not create the LaserficheSignInPage, you will not be able to save documents from the SharePoint Site to Laserfiche.

## Register Your App in the Developer Console
1. Open the [Developer Console](https://developer.laserfiche.com/developer-console.html)
1. Attempt to Create a New App from Manifest, and copy-paste the manifest provided [here](https://github.com/Laserfiche/laserfiche-sharepoint-integration/blob/1.x/UserDocuments/Laserfiche%20SharePoint%20Integration%20AppManifest.json)
1. If the attempt fails because an app with that client ID already exists, find the app with that client id by opening the following url in a new tab: https://app.laserfiche.com/devconsole/apps/<b>{your_client_id_goes_here}</b>/config, where the part in brackets should be replaced by the client_id you copied.
1. One way or another, an app with that client ID should now exist. Open the app in devconsole and switch from the general tab to the authentication tab.
1. Add the url of your SharePoint Page with the Laserfiche Web Part as a new redirect URI.
1. You should now be able to Sign In.
