# CopyPstFilesToStorageAccount.ps1

## Overview

The `CopyPstFilesToStorageAccount.ps1` script automates the process of handling PST files stored in Azure. It performs the following tasks:

1. Downloads PST files from a specified Azure location.
2. Processes the downloaded PST files.
3. Uploads the processed PST files to a specified Azure Storage container.

## Features

- **Authentication Methods**: Supports two authentication methods: using a certificate or using credentials.
- **Logging**: Logs the status of PST file uploads to a CSV file.
- **Error Handling**: Includes robust error handling and logging mechanisms.
- **Modular Functions**: Contains several helper functions to perform specific tasks like connecting to SharePoint, downloading files, and uploading files.

## Parameters

- **WorkingDirectory**: The directory where temporary files and logs will be stored.
- **ExcelFilePath**: The path to the Excel file containing the list of PST files to be processed. This us obtained by using ShareGate.
- **RequiredColumns**: An array of column names that are required in the Excel file. Default is `@("Name", "Location", "Size (MB)", "Site Address")`.
- **StorageAccountName**: The name of the Azure Storage account.
- **StorageAccountKey**: The key for the Azure Storage account.
- **StorageAccountContainer**: The name of the Azure Storage container where the PST files will be uploaded.
- **ClientID**: The client ID used for authentication.
- **TenantName**: The tenant name used for authentication (required for the 'Certificate' parameter set).
- **CertificateThumbprint**: The thumbprint of the certificate used for authentication (required for the 'Certificate' parameter set).
- **Credentials**: The credentials used for authentication (required for the 'Credentials' parameter set).

## Usage

### Example with Certificate-Based Authentication

```powershell
CopyPstFilesToStorageAccount.ps1 -WorkingDirectory "C:\PSTFiles" -ExcelFilePath "C:\PSTFiles\pst_list.xlsx" -StorageAccountName "mystorageaccount" -StorageAccountKey "myaccountkey" -StorageAccountContainer "pst-container" -ClientID "myclientid" -TenantName "mytenantname" -CertificateThumbprint "mycertthumbprint"
```
