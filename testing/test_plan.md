# Test Plan for Laserfiche SharePoint Online Integration

## Objective

Verify that changes made to the Laserfiche SharePoint Online Integration do not disrupt
existing functionality in the product. This plan should be executed prior to each new
release, and no changes should be included in the release until they have been tested. As
new functionality is added to the integration, new tests should be added to the plan
to ensure adequate coverage.

## Test Cases

- [Installation](#installation)
- [Site Configuration](#site-configuration)
- [Integration Configuration](#integration-configuration)
- [Save to Laserfiche](#save-to-laserfiche)
- [Repository View](#repository-view)

### **Installation**

#### **Use Documentation to Add the Integration to your Tenant App Catalog**

Steps:

1. Follow the instructions in the [Adding App to Organization Documentation](https://laserfiche.github.io/laserfiche-sharepoint-integration/docs/admin-documentation/adding-app-organization.html)

Expected Results:

- Instructions in documentation are clear and effective
- `Laserfiche SharePoint Online Integration` is available in your tenant app catalog

#### **Use Documentation to Add the Integration to your SharePoint Site**

Steps:

1. Follow steps 1-5 in the [Adding App to SharePoint Site Documentation](https://laserfiche.github.io/laserfiche-sharepoint-integration/docs/admin-documentation/adding-app-to-sp-site.html)

Expected Results:

- Instructions in documentation are clear and effective
- `Laserfiche SharePoint Online Integration` is available in your SharePoint Site site contents

### Site Configuration

Prerequisites:

- Follow the [Installation](#installation) steps successfully

#### Use Documentation to set up Laserfiche Sign In Page

Steps:

1. Follow the instructions for Laserfiche Sign in page in the [Add Add to SP Site Documentation](https://laserfiche.github.io/laserfiche-sharepoint-integration/docs/admin-documentation/add-app-to-sp-site)

Expected Results:

- Instructions in documentation are clear and effective
- You have a page in your site called LaserficheSignIn that contains the Laserfiche Sign In Web Part

#### Use Documentation to set up Laserfiche Repository Explorer

Steps:

1. Follow the instructions for the Repository Explorer page in the [Add Add to SP Site Documentation](https://laserfiche.github.io/laserfiche-sharepoint-integration/docs/admin-documentation/add-app-to-sp-site)

Expected Results:

- Instructions in documentation are clear and effective
- You have a page in your site that contains the Laserfiche Repository Explorer Web Part

#### Use Documentation to set up Laserfiche Admin Configuration

Steps:

1. Follow the instructions for Admin Configuration page in the [Add Add to SP Site Documentation](https://laserfiche.github.io/laserfiche-sharepoint-integration/docs/admin-documentation/add-app-to-sp-site)

Expected Results:

- Instructions in documentation are clear and effective
- You have a page in your site that contains the Laserfiche Administrator Configuration Web Part

#### Use Documentation to register app in dev console

Steps:

1. Follow the instructions for registering application in the [Register App in Laserfiche Documentation](https://laserfiche.github.io/laserfiche-sharepoint-integration/docs/admin-documentation/register-app-in-laserfiche.html)

Expected Results:

- Verify manifest is valid (i.e. SPA, correct clientId, etc.)
- App is registered successfully in Laserfiche dev console
- You can sign in on all three pages you created above
### Repository View

Prerequisites

- Repository Explorer web part must exist

Steps:

1. Create a new folder by clicking on the folder icon with the plus sign inside
   - Expected Result:
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

### Integration Configuration

Prerequisites:

- Follow the [Installation](#installation) and [Site Configuration](#site-configuration) steps successfully

#### Test Access Rights of Web Part
Prerequisites:
- the admin configuration web part must exist in a SharePoint Page
- you must NOT BE a site owner of the site containing that page.

Steps:
1. Attempt to open the admin configuration web part on the protected page.
  - Expected Results: You should not be able to do any configuring, and there should
  be an error message explaining that you don't have the necessary rights.

#### Create standard profile

Prerequisites:

- the admin configuration web part must exist in a SharePoint Page
- you must BE a site owner of the site containing that page.
- finish testing the functionality of the repository explorer web part

Steps:

1. Go to the Profiles tab and click the `Add Profile` button.
  - Expected Results: Something resembling the following Profile Editor appears: [Could Not Display Image](./assets/profileCreator.png)
1. Name the Profile `Example Profile Name`, do not select a template, select the Folder which you created in the functionality test of the Repository Explorer web part as the destination folder, and choose `Leave a copy of the file in SharePoint` for the `After import` behavior. Click the Save button.
  - Expected Results: You should see a success dialog, and then get returned to the `Profiles Tab`, where the new profile should be visible.
1. Go to the Profile Mapping tab and click the `Add` button.
1. Select `Document` for the SharePoint Content Type and select `Example Profile Name` for the `Laserfiche Profile`. Click the floppy disk icon to save.


#### Test Profile Error Handling
Prerequisites:

- the admin configuration web part must exist in a SharePoint Page
- you must BE a site owner of the site containing that page.

Steps:
1. Go to the Profiles tab and click the `Add Profile` button.
1. Name the Profile, `Bad Profile`, select the Folder which you created in the functionality test of the Repository Explorer web part as the destination folder, and select the `General` template. In the Mapping section, Click `Add Field`, and choose `Actual Work` for the SharePoint Column and `Date (2)` for the Laserfiche Field.
  - Expected Results: 
    - You should get a warning/error that the data types don't match.
    - You should't be able to save
1. Delete the SP Column/LF field pair.
  - Expected Results:
    - You should be able to save (button not disabled)
1. Save the Profile
1. Add a New Profile, and name it `Bad Profile` as well. Attempt to Save.
  - Expected Results:
      - The Profile should not be added.
      - The page should not indicate that the profile was added
      - The page should explain that the profile was not added because a profile with that name already exists.
1. Add a New Profile, and attempt to Save.
  - Expected Results:
      - The Profile should not be added.
      - The page should not indicate that the profile was added
      - The page should explain that the profile was not added because the profile lacked a name.
#### Test Edit Profile

Steps:
1. Click the pencil button to edit a profile and add some compatible metadata mappings like a text type for the SharePoint Column and a String type for the Laserfiche Field, for example. Click Save.
  - Expected Results:
    - The page should indicate that the profile was saved.
    - If you edit the profile, you should find that it saved your edits.

#### Test Creating Different Profiles

Steps:


#### Test Default Profile
Steps:



### Save to Laserfiche

Prerequisites:

- Laserfiche Sign In Page must already exist

#### Test happy path save

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
### Begin Additions
Admin page (https://lfdevm365.sharepoint.com/sites/TestSiteAlex/SitePages/Admin-Configuration.aspx) – On site you are an owner 


Attempt to mismatch field types in metadata mappings 

Ex/ Map SharePoint Content Type of Number to LF Field Type Date/Time 

Should prevent you from saving 

Attempt to save two profiles with same name 

Save ‘Test Profile’ 

Save ‘Test Profile’ 

Attempt to save with no name – Should prevent you from saving 

Edit Profile 

Update the metadata/action/etc. 

Save and click edit again to verify changes 

Add multiple profiles with: 

Different actions on save (delete, replace) 

Different metadata, things you know won’t exist at save time and things you know will exist (i.e. Author) 

Profile Mapping 

Map document to first test profile (test good case first) 

Map another content type to something else 

Test mapping default 

Admin Page – on site you are not owner 

Test that you cannot access page, it says you do not have rights 

Document save to Laserfiche 

Test expected “good” case 

With default mapping 

Test no content type on document 

Test unmapped content type 

Test mapped content type 

With NO default mapping 

Test no content type on document 

Test unmapped content type 

Test mapped content type 

Test metadata constraint failed case  

Add SharePoint Column “Actual Work” to SharePoint Library 

Add value for “Actual Work” for specific document to be very large number 

Use mapping that maps the “Actual Work” field to a number field in Laserfiche (i.e. Amount) 

Test different content type (ex/ form) 

Test required field doesn’t exist case 

Use mapping that maps a required field to a property that doesn’t exist on the document (most besides Author, Date Created, etc.) 

Test file type is URL case 

Attempt to save an item that is a URL (i.e one of the ones that has been replaced once saved to Laserfiche) 

Test replace with URL action 

Use mapping that has this action selected 

Test delete after save action 

Use mapping that has this action selected 

### End Additions

