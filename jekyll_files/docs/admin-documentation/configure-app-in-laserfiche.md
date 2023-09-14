---
layout: default
title: Configure App in Laserfiche
nav_order: 3
parent: Laserfiche SharePoint Integration Administration Guide
---
# Configure App in Laserfiche
## PRE-RELEASE DOCUMENTATION - SUBJECT TO CHANGE

1. Open the [Developer Console](https://developer.laserfiche.com/developer-console.html).
<a href="./assets/images/createAppFromManifest.png"><img src="./assets/images/createAppFromManifest.png"></a>
1. Attempt to Create a New App from Manifest, and upload the manifest provided [here](https://github.com/Laserfiche/laserfiche-sharepoint-integration/blob/1.x/UserDocuments/Laserfiche%20SharePoint%20Integration%20AppManifest.json).
<a href="./assets/images/createApplication.png"><img src="./assets/images/createApplication.png"></a>
1. If the attempt fails because an app with that client ID already exists, find the app with that client id by opening the following url in a new tab: https://app.laserfiche.com/devconsole/apps/<b>{your_client_id_goes_here}</b>/config, where the part enclosed in braces should be replaced by the client_id of the manifest linked in the previous step.
<a href="./assets/images/clientIdRegistered.png"><img src="./assets/images/clientIdRegistered.png"></a>
1. One way or another, an app with that client ID should now exist. Open the app and switch from the general tab to the authentication tab.
<a href="./assets/images/redirectUri.png"><img src="./assets/images/redirectUri.png"></a>
1. Add the URL of your SharePoint Page with the Laserfiche web part as a new redirect URI.
1. You should now be able to Sign In.
