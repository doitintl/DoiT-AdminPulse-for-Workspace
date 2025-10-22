# Self-Hosting DoiT AdminPulse for Workspace

This guide provides instructions on how to self-host the DoiT AdminPulse for Workspace application on Google Sheets and Google Cloud.

## Prerequisites

* A Google Cloud Platform (GCP) project.
* Access to the Google Workspace Admin SDK.

## Steps

### 1. Make a copy of the Google Sheet

1.  Click [here](https://docs.google.com/spreadsheets/d/1rbgKhzDYDmPDKuyx9_qR3CWpTX_ouacEKViuPwAUAf8/copy) to make a copy of the public Google Sheet and associated Apps Script.

### 2. Set up your Google Cloud Project

1.  Go to the [Google Cloud Console](https://console.cloud.google.com/).
2.  Create a new project or select an existing one.
3.  Enable the following APIs:
    *   Admin SDK API
    *   Google Drive API
    *   Google Sheets API
    *   Cloud Identity API
    *   Enterprise License Manager API
    *   Gmail API
    *   Apps Script API
    *   Group Settings API
4.  Create an **OAuth consent screen**.
    *   Select **Internal** for the user type.
    *   Add the following scopes, separated by commas:
        `https://www.googleapis.com/auth/admin.directory.user.readonly,https://www.googleapis.com/auth/admin.directory.domain.readonly,https://www.googleapis.com/auth/admin.directory.orgunit.readonly,https://www.googleapis.com/auth/admin.directory.group.readonly,https://www.googleapis.com/auth/admin.directory.group.member.readonly,https://www.googleapis.com/auth/admin.reports.audit.readonly,https://www.googleapis.com/auth/apps.licensing.readonly,https://www.googleapis.com/auth/drive.readonly,https://www.googleapis.com/auth/spreadsheets,https://www.googleapis.com/auth/script.container.ui,https://www.googleapis.com/auth/script.external_request,https://www.googleapis.com/auth/admin.directory.user.security,https://www.googleapis.com/auth/admin.directory.device.mobile.readonly,https://www.googleapis.com/auth/userinfo.email,https://www.googleapis.com/auth/apps.groups.settings,https://www.googleapis.com/auth/drive,https://www.googleapis.com/auth/apps.licensing,https://www.googleapis.com/auth/script.send_mail,https://www.googleapis.com/auth/admin.directory.customer.readonly,https://www.googleapis.com/auth/cloud-identity.policies.readonly,https://www.googleapis.com/auth/script.scriptapp`

### 3. Connect your Apps Script to your Google Cloud Project

1.  Open the Apps Script editor.
2.  Click on **Project Settings**.
3.  Under **Google Cloud Platform (GCP) Project**, click on **Change project**.
4.  Enter your GCP project number and click on **Set project**.

### 4. Run the Application

1.  Open the Google Sheet.
2.  Click on **Extensions**, click on **DoiT AdminPulse for Workspace**.
3.  Click on **Inventory Workspace Settings** -> **Check all policies**.
4.  Authorize the script to access your Google Workspace data.

Once the script has finished running, you will see a new sheet with the security checklist.