---
layout: default
title: User Documentation
nav_order: 3
---

## The Repository Access Webpart

### Description
The Laserfiche Repository Access Webpart displays a table whose rows
correspond to the folders and documents within the parent folder and whose columns describe metadata.

### Usage
- Sign In: To view your folders and documents, authenticate yourself
by clicking on the button labeled "Sign in to Laserfiche" and
subsequently logging in.
- Buttons
    - Open: Clicking this button will open the selected entry, if one exists
    - Upload File: Clicking this button will prompt you to select a file to upload to the currently open folder in the Repository Access webpart
    - New Folder: Clicking this button will prompt you to input a name for a new folder within the currently open folder.
- Navigation
    - Navigation Down
        - Double-click on a folder to view its contents within the webpart
        - Double-click on a file to view it in Laserfiche Repository
    - Navigation Up
        - Use the links in the Breadcrumb Navigation to jump back up to a higher level folder

## The Save to Laserfiche Webpart

### Description
The Save to Laserfiche Webpart allows you to export files directly from SharePoint to Laserfiche. It provides the essential functionality of the SharePoint/Laserfiche Integration. Advanced Settings can be configured in the [Admin Configuration Web Part](goes nowhere).

### Usage
- First-Time Setup - Authenticate youself by clicking on the button labeled "Sign in to Laserfiche".
- Typical Use - Right-click on any document in SharePoint and find the “Save to Laserfiche” item in the drop-down. If it does not exist, consult the [Admin Documentation](doesnt yet exist)

## The Admin Configuration Webpart

### Description
The Admin Configuration Web Part is used to define how the Save to
Laserfiche web part should map SharePoint metadata to Laserfiche
Metadata when sending Documents from SharePoint to Laserfiche, as well
as where to save it in Laserfiche. This webpart contains Three Tabs.

### About Tab
Provides information about the Web Part.

### Profiles Mapping Tab
- Open: Clicking this button will open the selected entry, if one
exists
- Upload File: Clicking this button will prompt you to select a file to upload to the currently open folder in the Repository Access webpart
- New Folder: Clicking this button will prompt you to input a name for a new folder within the currently open folder.

### Profiles Tab
Displays the list of currently defined Profiles. Click on the “Add Profile” Button, or on the Pencil Icon to open the Profile Editor to add a new profile or edit an Existing Profile, respectively.

### Profile Editor
- Name: this is the identifier used to associate SharePoint content types with this profile in the Profile Mapping tab
- Laserfiche Template: If a profile is assigned a template, then all files saved to Laserfiche through that profile will be assigned that template in Laserfiche. [Learn more about templates](https://doc.laserfiche.com/laserfiche.documentation/en-us/Content/Fields_and_Templates.html)
- After Import: This option specifies what to do with the
SharePoint file after exporting it to Laserfiche
- Mappings from SharePoint Column to Laserfiche Field Values
    - This is where the actual metadata transfer is configured.
    - Each Field in the template can be assigned a SharePoint column, so that when files are exported from SharePoint to Laserfiche, the file in Laserfiche will have a field with the same value as the Column of the file in SharePoint
    - templates with required fields MUST have columns assigned to them.
    - The association between SharePoint columns and Laserfiche fields should be one-to-one, i.e., you should not attempt to map multiple SharePoint columns to the same Laserfiche field.
