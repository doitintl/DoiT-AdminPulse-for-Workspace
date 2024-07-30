# Security Checklist for Workspace Admins

A tool by DoiT for Workspace administrators to review their security posture and inventory the admin SDK.

## Overview:

This is a free and open-source tool by DoiT developed for Google Workspace Administrators to conduct a comprehensive security review of their organization's Google Workspace or Cloud Identity environments.

The tool has two main components: The checklist of security settings with links to documentation as well as to the relevant setting in the admin console, and read-only access to the Admin SDK to inventory settings for users, groups, OAuth tokens and more.

## Features:

* Security Checklist: A comprehensive list of security controls, links to Google's documentation, best practice recommendations to the relevant area of the admin console.
* Note-taking Section: Keep track of your progress by taking notes directly within the checklist. Prepare questions to be discussed during a review session with support engineers.
* Inventory scripts (Admin SDK): Utilize the scripts in the Extensions menu to run app script code that inventories the environment and adds reports for Users, License assignments, Google Group Settings, Groups Membership, Mobile Devices, Shared Drive settings, OAuth Tokens, App Passwords, Organizational units, and Domains and DNS records.

## Prerequisites:

This tool requires super admin access to your Google Workspace or Cloud Identity account and access to the following APIs: Admin SDK, Groups Settings API, Google Sheets API, Enterprise License Manager API and Drive API.

## Getting Started:

1. Make a copy of the [Security Checklist for Workspace Admins](https://docs.google.com/spreadsheets/d/1rbgKhzDYDmPDKuyx9_qR3CWpTX_ouacEKViuPwAUAf8/copy), which will also copy the App Script code.
2. Open the workbook and begin the security review process, starting with the Google Workspace Security Checklist.
3. Use the provided links for guidance and take notes as needed.
4. Optionally, run the script by navigating to Extensions > [Public] Security Checklist for Workspace Admins > Run all scripts to inventory users, group memberships, and more.
5. The completed workbook will be helpful to identify areas of Google Workspace where your organization may be able to improve security.
![Fetch Info menu button\](image.png)](<Fetch Info.png>)

## Contribution:

Contributions are welcome! Please refer to the How to Contribute.md in the GitHub project for guidelines on how to contribute, submit issues, or propose new features. If you have ideas for improvements or new features, follow the outlined process in the contribution guide.


## Data privacy / What could go wrong:

* The sheet is made available to the public via Google Sheets View Only access.
* When the workbook is copied, you are also prompted to copy the associated appscript project. Your copy of the workbook and appscript project are then owned entirely by your account and unlinked from the public copy. Any changes to the public copy of the workbook or the public appscript will not be reflected in your copy of the workbook.
* When a copy is made and scripts are run, the Google Sheet will contain sensitive information about your organization’s directory including user email addresses, names, Google Group email addresses, memberships, and more. For this reason, it is recommended to only share the document with other Google Workspace administrators that are a part of your organization or trusted partners.
* The app script code does not transmit, share or otherwise log the returned data to anywhere else besides your copy of the Google Sheet. 
* The scripts will not work for non-super admin users. The scripts cannot be run against an external organization to which your Google account does not belong. 
* The scripts are designed as a read-only reporting mechanism. Once the workbook is copied, the your copy of the app script code could be modified to provide addtional functionoality to Google Workspace Admin SDK. Read/write access can be dangerous and outside of the scope of this tool.
