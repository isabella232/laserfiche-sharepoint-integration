# Test Plan for Laserfiche SharePoint Integration

## Objective

We wish to verify that changes made to the Laserfiche SharePoint Integration do not disrupt
existing functionality in the product. This plan should be executed prior to each new
release, and no changes should be included in the release until they have been tested. As
new functionality is added to the integration, new tests should be added to the plan
to ensure adequate coverage.

## Test Cases

### Use Documentation to Install the Laserfiche SharePoint Integration sppkg

Prerequisistes:

- None

Steps:

1. Follow the instructions in the [Documentation](https://laserfiche.github.io/laserfiche-sharepoint-integration/docs/admin-documentation.html#deploy-laserfiche-sharepoint-integration-to-a-sharepoint-site) for adding the sppkg file to SharePoint.

Expected Results:

- Instructions in documentation are clear and effective
- sppkg file ends up added to SharePoint

### Use Documentation to set up Laserfiche Sign In Page

Prerequisites:

- the SharePoint package (.sppkg) file must already be installed to SharePoint.
- follow the instructions in the README.md for running locally

Steps:

1. Follow the instructions in the [Documentation](https://laserfiche.github.io/laserfiche-sharepoint-integration/docs/admin-documentation/add-app-to-sp-site) for setting up the Laserfiche Sign In Page.

Expected Results:

- Instructions in documentation are clear and effective
- the Laserfiche Sign In Page is created with the sign-in web part.

### Test Functionality of Save To Laserfiche

Prerequisites:

- Laserfiche Sign In Page must already Exist

Steps:

1. Upload a document of some kind with some text to the Document's tab of a SharePoint site
1. Right-click on the document.
1. Select the Save To Laserfiche option
1. Select View File in Laserfiche
1. Return to the original tab and select close

Expected Results:

1. Does not test Integration behavior
1. The `Save to Laserfiche` option should exist in the resulting drop down.
1. A dialog should immediately open, and eventually display a success message and a button saying `View File in Laserfiche`
1. The file should be opened in a new tab with a `Back` button. Clicking the `Back` button should display a folder view containing the document saved to Laserfiche.
1. The dialog should disappear.

### Use Documentation to set up Repository Explorer page

Prerequisites

- the SharePoint package (.sppkg) file must already be installed to SharePoint.
- follow the instructions in the README.md for running locally

Steps:

1. Follow the instructions in the [Documentation](https://laserfiche.github.io/laserfiche-sharepoint-integration/docs/admin-documentation/add-app-to-sp-site) for setting up the Repository Explorer Page.

Expected Results:

- Instructions in documentation are clear and effective
- a SharePoint page containing the Repository Explorer web part is created

### Test functionality of repository explorer web part

Prerequisites

- Repository Explorer web part must exist

Steps:

1. Create a new folder by clicking on the folder icon with the plus sign inside
1. Double-click on the folder to go into it.
1. Create another folder within the first and click into it.
1. Click on the upload button and choose a file to upload.
1. Click on the uploaded file to select it and then click the icon of the the northeast arrow in a square to open the file in Laserfiche
1. Use the breadcrumb navigation above the column titles to navigate back to the top level of the repository.

Expected Results:

1. We should be able to see the new folder without needing to reload anything
1. We should see an empty set of children and the breadcrumb navigation should appear
1. We should be able to see the new folder without needing to reload anything. Upon double-clicking, we should see an empty set of children and the breadcrumb navigation should still allow us to move back up to the ancestor folders.
1. A dialog should pop up allowing you to choose the file. After you've chosen, the file should be immediately visible.
1. the selected file row should be visibly distinguished from the other non-selected rows. Additionally, clicking on the arrow-in-square should open a new tab to view the file in Laserfiche.
1. the breadcrumb navigation links should take us to the named locations.

### Use Documentation to set up Admin Configuration Page

Prerequisites

- the SharePoint package (.sppkg) file must already be installed to SharePoint.
- follow the instructions in the README.md for running locally

Steps:

1. Follow the instructions in the [Documentation](https://laserfiche.github.io/laserfiche-sharepoint-integration/docs/admin-documentation/add-app-to-sp-site) for setting up the Admin Configuration Page.

Expected Results:

- Instructions in documentation are clear and effective
- a SharePoint page containing the admin configuration web part is created

### Test Functionality of admin configuration web part

Prerequisites:

- the admin configuration web part must exist in a SharePoint Page
- finish testing the functionality of the repository explorer web part

Steps:

1. Go to the Profiles tab and click the `Add Profile` button.
1. Name the Profile `Example Profile Name`, do not select a template, select the Folder which you created in the functionality test of the Repository Explorer web part as the destination folder, and choose `Leave a copy of the file in SharePoint` for the `After import` behavior. Click the Save button.
1. Go to the Profile Mapping tab and click the `Add` button.
1. Select `Document` for the SharePoint Content Type and select `Example Profile Name` for the `Laserfiche Profile`. Click the floppy disk icon to save.

Expected Results:

1. Something resembling the following Profile Editor appears: [Could Not Display Image](./assets/profileCreator.png)
1. You should get a Success dialog, and then get returned to the `Profiles Tab`, where the new profile should be visible.
