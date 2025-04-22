# DoiT AdminPulse for Workspace

A tool by DoiT for Workspace administrators to review their security posture and inventory the admin SDK.


<a href="https://workspace.google.com/marketplace/app/doit_adminpulse_for_workspace/639424393187?pann=b">
  <img src="https://workspace.google.com/static/img/marketplace/en/gwmBadge.svg?style=flat-square" alt="Google Workspace Marketplace" height="68px">
</a>

[![License: MIT](https://img.shields.io/badge/License-MIT-green.svg)](https://github.com/doitintl/DoiT-AdminPulse-for-Workspace/blob/main/LICENSE)

## Overview:

This is a free and open-source tool by DoiT developed for Google Workspace™ Administrators to conduct a comprehensive security review of their organization's Google Workspace™ or Cloud Identity environments.

The tool has two main components: The checklist of security settings with links to documentation as well as to the relevant setting in the admin console, and read-only access to the Admin SDK to inventory settings for users, groups, OAuth tokens and more.

## Features:

* Security Checklist: A comprehensive list of security controls, links to Google's documentation, best practice recommendations to the relevant area of the admin console.
* Integration with the [Cloud Identity Policy API](https://cloud.google.com/identity/docs/concepts/overview-policies) to list OU and group policies.
* Inventory scripts (Admin SDK): Utilize the scripts in the Extensions menu to run app script code that inventories the environment and adds reports for Users, License assignments, Google Groups™ Settings, Groups Membership, Mobile Devices, Shared Drive settings, OAuth Tokens, App Passwords, Organizational units, and Domains and DNS records.

## Prerequisites:

This tool requires super admin access to your Google Workspace™ or Cloud Identity account and access to the following APIs: Admin SDK, Groups Settings API, Google Sheets™ API, Enterprise License Manager API and Drive API.

## Getting Started:

1. Use your Super Admin account to install the [DoiT AdminPulse for Workspace](https://workspace.google.com/marketplace/app/doit_adminpulse_for_workspace/639424393187) Google Sheets Editor Add-on.
2. Open a new Google Sheet and use **Extensions > DoiT AdminPulse for Workspace > Setup or Refresh Sheet** populate the checklist.
3. Use the Cloud Identity Policies API by navigating to **Extensions > DoiT AdminPulse for Workspace > Inventory Workspace Settings > Check all policies**
4. Use the provided documentation links for guidance and take notes as needed while reviewing the enviorment.
5. Optionally, run the inventory scripts by navigating to **Extensions > DoiT AdminPulse for Workspace > Run all scripts** to inventory users, group memberships, and more.
6. The completed workbook will be helpful to identify areas of Google Workspace™ where your organization may be able to improve security.
![Fetch Info menu button\](image.png)](<Fetch Info.png>)

## Contribution:

Contributions are welcome! Please refer to the How to Contribute.md in the GitHub project for guidelines on how to contribute, submit issues, or propose new features. If you have ideas for improvements or new features, follow the outlined process in the contribution guide.


## Data privacy / What could go wrong:

* The sheet is made available to the public via Google Sheets™ View Only access.
* When the application completes it's run, the Google Sheet will contain sensitive information about your organization’s directory including user email addresses, names, Google Groups™ email addresses, memberships, and more. For this reason, it is recommended to only share the document with other Google Workspace™ administrators that are a part of your organization or trusted partners.
* The app script code does not transmit, share or otherwise log the returned data to anywhere else besides your copy of the Google Sheet. 
* The scripts will not work for non-super admin users. The scripts cannot be run against an external organization to which your Google account does not belong. 
* The scripts are designed as a read-only reporting mechanism.
