# Report_used_spaces

## Description
This Powershell script downloads a file from SharePoint Online to a local path, connects to SharePoint, OneDrive, Exchange to collect data about used spaces and also connects to local servers to collect data about used disk spaces then exports the collected data to an Excel file and adds the file to a SharePoint location.

## Prerequisites
- PowerShell version 5.1 or later
- Installed SharePointPnPPowerShellOnline module
- Installed ExchangeOnlineManagement
- Installed ImportExcel module
- The user who runs the script should be a member of the Remote Management Users group in AD

## Configuration
Before running the script, ensure the following variables are properly configured:

- `$filename`: Provide the desired filename for the Excel report (e.g., `Raport_used_spaces.xlsx`).
- `$localPath`: Specify the local path where the Excel report will be saved (e.g., `C:\Raporty\`).
- `$siteUrl`: Set the URL of the SharePoint site where data will be collected (e.g., `https://company.sharepoint.com/sites/it-dep`).
- `$onlinePath`: Specify the path on SharePoint where the file should be added (e.g., `Shared Documents/Global/`).
- `$tenant`: Provide the name of the tenant (e.g., `company.onmicrosoft.com`).
- `$appId`: Set the Client ID of the Azure AD application registered for this script.
- `$thumbprint`: Provide the certificate thumbprint associated with the Azure AD application.

Also if you want collect data from servers ensure the following variables are properly configured:
- $servers1 = @('`server`', '`server2`') #Windows 2012
- $servers2 = @('`server3`', '`server4`', '`server5`', '`server6`', '`server7`', '`server8`', '`server9`', '`server10`') #Windows 2016-2019 / Windows 10
- $servers3 = @('`server11`', '`server12`', '`server13`') #Windows 2022 / Windows 11
