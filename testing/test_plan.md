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

Expected Result
- Instructions in documentation are clear and effective
- sppkg file ends up added to SharePoint
### Use Documentation to set up Laserfiche Sign In Page

Prerequisites:
- the SharePoint package (.sppkg) file must already be installed to SharePoint.
- follow the instructions in the README.md for running locally

Steps:
1. Follow the instructions in the [Documentation](https://laserfiche.github.io/laserfiche-sharepoint-integration/docs/admin-documentation/add-app-to-sp-site#) for setting up the Laserfiche Sign In Page.

Expected Result:
- Instructions in documentation are clear and effective
- the Laserfiche Sign In Page is created

### Verify Save To Laserfiche Works
Prerequisites:
- Laserfiche Sign In Page must already Exist

Steps:
1. Upload a word document with some text to a SharePoint site
1. Right click on the document
1. Select the Save To Laserfiche option
1. Select View File in Laserfiche
1. Return to the original tab and select close

Expected Result(s):
1. Each of the steps above can be done (e.g., the Save To Laserfiche Option exists for Step #3)
1. A word document is saved to Laserfiche, and its contents are identical to the original in SharePoint

### Set up Repository Explorer web part using Documentation
Prerequisites
- the SharePoint package (.sppkg) file must already be installed to SharePoint.
- follow the instructions in the README.md for running locally

Steps:
1. Follow the instructions in the [Documentation](https://laserfiche.github.io/laserfiche-sharepoint-integration/docs/admin-documentation/add-app-to-sp-site) for setting up the Repository Explorer Page.

Expected Result:
- Instructions in documentation are clear and effective
- the Laserfiche Repository Explorer is created

### Test functionality of Repository Explorer
Prerequisites
- Repository Explorer web part must exist

Steps:
1. Create a new folder by clicking on the folder icon with the plus sign inside
1. Double Click on the folder to go into it.
1. Create another folder within the first and click into it.
1. Click on the upload button and choose a file to upload.
1. Click on the uploaded file to select it and then click the icon of the the northeast arrow in a square to open the file in Laserfiche
1. Use the breadcrumb navigation above the Column titles to navigate back to the top level of the repository.

Expected Results:
1. We should be able to see the new folder without needing to reload anything
1. We should see an empty set of children and the breadcrumb navigation should appear
1. We should be able to see the new folder without needing to reload anything. Upon double clicking, we should see an empty set of children and the breadcrumb navigation should be still allow us to move back up to the ancestor folders.
1. A dialog should pop up allowing you to choose the file. After you've chose, the file should be immediately visible.
1. the file row should be visibly selected after clicking and clickingon the arrow-in-square should open a new tab to view the file in Laserfiche
1. the breadcrumb navigation links should take us to the named locations.

