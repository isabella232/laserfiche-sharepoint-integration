---
layout: default
title: Add App to SharePoint Site
nav_order: 2
parent: Laserfiche SharePoint Integration Administration Guide
---


# Add App to SharePoint Site
## PRE-RELEASE DOCUMENTATION - SUBJECT TO CHANGE
### Prerequisites
  - Be the owner of a SharePoint Site in an organization which has the Laserfiche integration installed. (see [Add App to Organization](./add-app-organization)])
###  The Laserfiche Sign In Page:
1. In your SharePoint Site, select the "Pages" item in the navigation bar on the left side of the page.
<a href="../assets/images/newSitePage.png"><img src="../assets/images/newSitePage.png"></a>
1. Create and open a new site page by clicking the blue "+ New" button and selecting "Site Page" from the dropdown.
1. Title the page “LaserficheSignIn”.
1. Move your cursor just below the title area to the white space beneath. This should reveal a hidden "+" button. If you hover over it, it should display the message "Add a new web part in column one”.
<a href="../assets/images/hiddenPlusButton.png"><img src="../assets/images/hiddenPlusButton.png"></a>
1. Click on that button and Search for “Laserfiche Sign In".
<a href="../assets/images/searchRepositoryExplorer.png"><img src="../assets/images/searchRepositoryExplorer.png"></a>
1. Click on the search result with a white L on an orange square. The Laserfiche Sign In web part should now appear on your Page. Before using the web part, make sure to [Register Your App in the Laserfiche Developer Console](https://laserfiche.github.io/laserfiche-sharepoint-integration/docs/admin-documentation.html#register-your-app-in-the-developer-console)
1. After you register your App and give it the right redirect URI, you should be able to log in and use the component. For Documentation on how to use the components, reference the [User Documentation](./user-documentation/).

### The Admin Configuration Page:
 Follow the same steps as above, but title the page whatever you wish and add the “Admin Configuration” web part instead of “Laserfiche Sign In".

### The Repository Explorer Page:
Follow the same steps as above, but title the Page whatever you wish and add the “Repository Explorer” web part.
