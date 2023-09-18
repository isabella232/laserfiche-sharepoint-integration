---
layout: default
title: Add App to Organization
nav_order: 1
parent: Laserfiche SharePoint Integration Administration Guide
---

# Add App to Organization

## PRE-RELEASE DOCUMENTATION - SUBJECT TO CHANGE


### Prerequisites
  - Have a SharePoint Online account with administrator privileges for the tenant app catalog.
  - Download the latest Laserfiche SharePoint Integration [package](./assets/laserfiche-sharepoint-integration.sppkg)

### Steps
1. Navigate to the following url: https://<b>{your-full-subdomain.and-domain.com}</b>/sites/appcatalog/AppCatalog/Forms/AllItems.aspx, where the part in curly braces is replaced by the domain and subdomain of your SharePoint-related websites.
1. If you can see the Add and Upload buttons, congratulations - you may proceed. If not, ask an administrator to [add you as an admin to the SharePoint Online App Catalog](https://learn.microsoft.com/en-us/office365/customlearning/addappadmin#add-an-administrator).
1. Click Upload and select the Laserfiche SharePoint package file (.sppkg).
<a href="./assets/images/uploadSppkgFile.png"><img src="./assets/images/uploadSppkgFile.png"></a>
  - NOTE: It's possible to build a new SharePoint Integration package file directly from source code by following the instructions in this [README.md](https://github.com/Laserfiche/laserfiche-sharepoint-integration#readme).
1. Back in your SharePoint Site, navigate to the app catalog by clicking on the "Site Contents" item in the
navigation bar.
<a href="./assets/images/sharePointSiteContents.png"><img src="./assets/images/sharePointSiteContents.png"></a>
1. Open the "New" Dropdown menu by clicking on the "+" icon.
<a href="./assets/images/NewDropDown.png"><img src="./assets/images/NewDropDown.png"></a>
1. Add the App named “laserfiche-sharepoint-integration-client-side-solution”.
<a href="./assets/images/addTheApp.png"><img src="./assets/images/addTheApp.png"></a>
1. Enable the app if you are asked to do so.
1. Navigate to your SharePoint site. On successful installation "Laserfiche SharePoint Integration" app is listed under the “Site Contents” tab.
<a href="./assets/images/appInstalled.png"><img src="./assets/images/appInstalled.png"></a>
