---
layout: default
title: Add App to SharePoint Site
nav_order: 2
parent: Laserfiche SharePoint Integration Administration Guide
---


# Add App to SharePoint Site
## PRE-RELEASE DOCUMENTATION - SUBJECT TO CHANGE
### Prerequisites
  - Be the owner of a SharePoint Site in an organization which has the Laserfiche integration installed. (see [Add App to Organization](./add-app-organization))

1. Create a new SharePoint site by going to the Home Page and selecting the Create Site button
<a href="../assets/images/createSite.png"><img src="../assets/images/createSite.png"></a>
1. Choose Team Site or Communication Site according to your preference and finish initializing the site.
1. In your newly created site, Navigate to the app catalog by clicking on the "Site Contents" item in the
navigation bar.
<a href="../assets/images/sharePointSiteContents.png"><img src="../assets/images/sharePointSiteContents.png"></a>
1. Open the "New" Dropdown menu by clicking on the "+" icon.
<a href="../assets/images/NewDropDown.png"><img src="../assets/images/NewDropDown.png"></a>
1. Add the App named “laserfiche-sharepoint-integration-client-side-solution”.
<a href="../assets/images/addTheApp.png"><img src="../assets/images/addTheApp.png"></a>
1. Enable the app if you are asked to do so.
1. Navigate to your SharePoint site. On successful installation "Laserfiche SharePoint Integration" app is listed under the “Site Contents” tab.
<a href="../assets/images/appInstalled.png"><img src="../assets/images/appInstalled.png"></a>

###  The Laserfiche Sign In Page:
1. In your SharePoint Site, select the "Pages" item in the navigation bar on the left side of the page.
<a href="../assets/images/newSitePage.png"><img src="../assets/images/newSitePage.png"></a>
1. Create and open a new site page by clicking the "+ New" button and selecting "Site Page" from the dropdown.
1. Title the page “LaserficheSignIn”.
1. Move your cursor just below the title area to the white space beneath. This should reveal a hidden "+" button. If you hover over it, it should display the message "Add a new web part in column one”.
<a href="../assets/images/hiddenPlusButton.png"><img src="../assets/images/hiddenPlusButton.png"></a>
1. Click on that button and Search for “Laserfiche Sign In".
<a href="../assets/images/searchRepositoryExplorer.png"><img src="../assets/images/searchRepositoryExplorer.png"></a>
1. Click on the search result with a white L on an orange square. The Laserfiche Sign In web part should now appear on your Page. Before creating subsequent pages, make sure to click the 'Publish' button to save the page. 

### The Repository Explorer Page:
Follow the same steps as above, but title the Page whatever you wish and add the “Repository Explorer” web part.

### The Admin Configuration Page:
 Follow the same steps as above, but title the page whatever you wish and add the “Admin Configuration” web part instead of “Laserfiche Sign In".

### Next Steps
Before you can log in and use the web pages you just created, you will need to [Register them in the Laserfiche Developer Console](../admin-documentation/register-app-in-laserfiche). After you register your Apps, you should be able to log in and use the web parts. For Documentation on how to use the integration, reference the [User Documentation](./user-documentation/).
