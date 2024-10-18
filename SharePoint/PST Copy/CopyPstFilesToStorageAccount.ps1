<#
.SYNOPSIS
Automates the process of downloading PST files from Azure, processing them, and uploading them to an Azure Storage container.

.DESCRIPTION
This script automates the workflow of handling PST files stored in Azure. It performs the following tasks:
1. Downloads PST files from a specified Azure location.
2. Processes the downloaded PST files.
3. Uploads the processed PST files to a specified Azure Storage container.

The script supports two authentication methods: using a certificate or using credentials. It requires the user to specify various parameters such as the working directory, Excel file path, storage account details, and client ID.

.PARAMETER WorkingDirectory
The directory where temporary files and logs will be stored.

.PARAMETER ExcelFilePath
The path to the Excel file containing the list of PST files to be processed.

.PARAMETER RequiredColumns
An array of column names that are required in the Excel file. Default is @("Name", "Location", "Size (MB)", "Site Address").

.PARAMETER StorageAccountName
The name of the Azure Storage account.

.PARAMETER StorageAccountKey
The key for the Azure Storage account.

.PARAMETER StorageAccountContainer
The name of the Azure Storage container where the PST files will be uploaded.

.PARAMETER ClientID
The client ID used for authentication.

.PARAMETER TenantName
The tenant name used for authentication (required for the 'Certificate' parameter set).

.PARAMETER CertificateThumbprint
The thumbprint of the certificate used for authentication (required for the 'Certificate' parameter set).

.PARAMETER Credentials
The credentials used for authentication (required for the 'Credentials' parameter set).

.NOTES
Ensure that you have the necessary permissions to access the Azure locations and the specified directories.

.EXAMPLE
.\PSTCopy - New.ps1 -WorkingDirectory "C:\PSTFiles" -ExcelFilePath "C:\PSTFiles\pst_list.xlsx" -StorageAccountName "mystorageaccount" -StorageAccountKey "myaccountkey" -StorageAccountContainer "pst-container" -ClientID "myclientid" -TenantName "mytenantname" -CertificateThumbprint "mycertthumbprint"

This example runs the script using certificate-based authentication.

.EXAMPLE
.\PSTCopy - New.ps1 -WorkingDirectory "C:\PSTFiles" -ExcelFilePath "C:\PSTFiles\pst_list.xlsx" -StorageAccountName "mystorageaccount" -StorageAccountKey "myaccountkey" -StorageAccountContainer "pst-container" -ClientID "myclientid" -Credentials (Get-Credential)

This example runs the script using credentials-based authentication.
#>
[CmdletBinding(DefaultParameterSetName = 'Certificate',
    SupportsShouldProcess = $true)]
param
(
    [Parameter(Mandatory = $true)]
    [string]$WorkingDirectory,

    [Parameter(Mandatory = $true)]
    [string]$ExcelFilePath,

    [Parameter(Mandatory = $false)]
    [string[]]$RequiredColumns = @("Name", "Location", "Size (MB)", "Site Address"),

    [Parameter(Mandatory = $true)]
    [string]$StorageAccountName,

    [Parameter(Mandatory = $true)]
    [string]$StorageAccountKey,

    [Parameter(Mandatory = $true)]
    [string]$StorageAccountContainer,

    [Parameter(Mandatory = $true)]
    [string]$ClientID,

    [Parameter(ParameterSetName = 'Certificate',
        Mandatory = $true)]
    [string]$TenantName,

    [Parameter(ParameterSetName = 'Certificate',
        Mandatory = $true)]
    [string]$CertificateThumbprint,

    [Parameter(ParameterSetName = 'Credentials',
        Mandatory = $true)]
    [pscredential]$Credentials
)
Clear-Host
$ErrorActionPreference = "Stop"
$Script:WorkingDirectory = $WorkingDirectory
$env:PNPPOWERSHELL_UPDATECHECK = "off"
$env:PNPPOWERSHELL_DISABLETELEMETRY = $true

Import-Module -Name "Az.Storage" -Force
Import-Module -Name "ImportExcel" -Force
Import-Module -Name "PnP.PowerShell" -Force


<#
.SYNOPSIS
Logs the status of PST file uploads to a CSV file.
.DESCRIPTION
This function logs the status of PST file uploads to a CSV file, including the file name, file path, upload time, and status.
.PARAMETER PstFileName
The name of the PST file.
.PARAMETER FilePath
The local file path of the PST file.
.PARAMETER Status
The status of the upload operation.
.PARAMETER CsvPath
The path to the CSV log file. Default is "C:\PSTFilesUploadLog.txt".
.EXAMPLE
Write-UploadLog -PstFileName "example.pst" -FilePath "C:\PSTFiles\example.pst" -Status "Uploaded"
#>
function Write-UploadLog {
    param (
        [Parameter(Mandatory = $true)]
        [string]$PstFileName,
        [Parameter(Mandatory = $true)]
        [string]$FilePath,
        [Parameter(Mandatory = $true)]
        [string]$Size,
        [Parameter(Mandatory = $true)]
        [string]$Status,
        [Parameter(Mandatory = $false)]
        [string]$CsvPath
    )

    $csvEntry = [pscustomobject]@{
        FileName   = $PstFileName
        FilePath   = $FilePath
        Size       = $Size
        UploadTime = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
        Status     = $Status
    }

    if (-not (Test-Path $csvPath)) {
        $csvEntry | Export-Csv -Path $csvPath -NoTypeInformation
    }
    else {
        $csvEntry | Export-Csv -Path $csvPath -NoTypeInformation -Append


    }
}

<#
.SYNOPSIS
Gets the size of a file and returns it as a formatted string.

.DESCRIPTION
The Get-FileSizeString function takes a file path as input and returns the size of the file as a formatted string. The size is returned in KB, MB, or GB depending on the file size.

.PARAMETER FilePath
The path to the file whose size is to be retrieved.

.RETURNS
A string representing the size of the file in KB, MB, or GB.

.EXAMPLE
$size = Get-FileSizeString -FilePath "C:\path\to\file.txt"
Write-Output $size

This example retrieves the size of the specified file and outputs it as a formatted string.

.NOTES
If the file does not exist, the function returns "0.00 KB".
#>
function Get-FileSizeString {
    param (
        [Parameter(Mandatory = $true)]
        [string]$FilePath
    )
    if (-not (Test-Path $FilePath)) {
        Write-Verbose "File not found: $FilePath"
        return "0.00 KB"
    }

    $FileSize = (Get-Item $FilePath).Length

    if ($FileSize -lt 1GB) {
        $FileSizeString = "{0:N2} MB" -f ($FileSize / 1MB)
        if ($FileSize -lt 1MB) {
            $FileSizeString = "{0:N2} KB" -f ($FileSize / 1KB)
        }
    }
    else {
        $FileSizeString = "{0:N2} GB" -f ($FileSize / 1GB)
    }

    return $FileSizeString
}

<#
.SYNOPSIS
Stops the Excel process using the specified file path.
.DESCRIPTION
The Stop-ExcelProcess function stops the Excel process that is using the specified file path. It retrieves the process ID of the Excel process using the file and forcibly stops the process.
.PARAMETER ExcelFilePath
The file path of the Excel file being used by the Excel process.
.EXAMPLE
Stop-ExcelProcess -ExcelFilePath "C:\path\to\excel.xlsx"
#>
function Stop-ExcelProcess {
    [CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'High')]
    param (
        [string]$ExcelFilePath
    )

    Write-Warning "The file is being used by another process. Attempting to close the process..."
    $processes = Get-Process -Name EXCEL | Where-Object { $_.MainWindowTitle -like "*$(Split-Path $ExcelFilePath -Leaf)*" }
    if ($processes) {
        foreach ($process in $processes) {
            if ($PSCmdlet.ShouldProcess($process.Name, "Close Process")) {
                Write-Host "Closing process: $($process.Name) (ID: $($process.Id))"
                Stop-Process -Id $process.Id -Force
            }
            else {
                Write-Error "Failed to close the process using the file. Please close it manually and try again."
                exit 1
            }
        }
    }
    else {
        Write-Error "Failed to close the process using the file. Please close it manually and try again."
        exit 1
    }
}

<#
.SYNOPSIS
Checks if a property exists in an object.

.DESCRIPTION
The Test-PropertyExistance function checks if a specified property exists in a given object. It returns a boolean value indicating whether the property exists or not.

.PARAMETER Object
The object to check for the existence of the property.

.PARAMETER PropertyName
The name of the property to check for.

.OUTPUTS
System.Boolean
Returns $true if the property exists, otherwise returns $false.

.EXAMPLE
$object = [PSCustomObject]@{
    Name = "John"
    Age = 30
}

Test-PropertyExistance -Object $object -PropertyName "Name"
# Returns $true

Test-PropertyExistance -Object $object -PropertyName "Address"
# Returns $false
#>
function Test-PropertyExistance {
    [CmdletBinding()]
    [OutputType([bool])]
    param (
        [Parameter(Mandatory = $true)]
        [object]$TestObject,

        [Parameter(Mandatory = $true)]
        [string]$PropertyName
    )

    return $TestObject.PSObject.Properties.Name -contains $PropertyName
}

<#
.SYNOPSIS
    Retrieves a list of records from an Excel file based on the specified columns.

.DESCRIPTION
    The Get-PSTListFromExcel function retrieves a list of records from an Excel file. It takes the path of the Excel file and an array of required columns as input parameters. The function checks if the Excel file exists and if it contains any records. It also validates if all the required columns are present in the Excel file. If all the validations pass, the function returns the records with the specified columns.

.PARAMETER ExcelFilePath
    The path of the Excel file from which to retrieve the records.

.PARAMETER RequiredColumns
    An array of column names that are required in the Excel file.

.OUTPUTS
    Returns a collection of records with the specified columns.

.EXAMPLE
    Get-PSTListFromExcel -ExcelFilePath "C:\path\to\excel.xlsx" -RequiredColumns "Name", "Email"

    This example retrieves a list of records from the "excel.xlsx" file located at "C:\path\to\". It specifies that the "Name" and "Email" columns are required in the Excel file.

.NOTES
    This function requires the Import-Excel module to be installed. You can install it by running the following command:
    Install-Module -Name ImportExcel

    For more information about the Import-Excel module, visit: https://www.powershellgallery.com/packages/ImportExcel

#>
function Get-PSTListFromExcel {
    [CmdletBinding(SupportsShouldProcess)]
    [OutputType([PSCustomObject])]
    param (
        [Parameter(Mandatory = $true)]
        [string]$ExcelFilePath,
        [Parameter(Mandatory = $true)]
        [string[]]$RequiredColumns
    )

    begin {
        Write-Verbose "Reading Excel File: $ExcelFilePath"
        Write-Verbose "Testing Path: $ExcelFilePath"
        if (!(Test-Path $ExcelFilePath)) {
            Write-Error -Message "Excel file not found at $ExcelFilePath"
        }

        Write-Verbose "Importing Excel File"
        try {
            $Records = Import-Excel -Path $ExcelFilePath
        }
        catch [System.IO.IOException] {
            Stop-ExcelProcess -ExcelFilePath $ExcelFilePath
            $Records = Import-Excel -Path $ExcelFilePath
        }
        catch {
            Write-Error -Message "Failed to import Excel file. Error: $($_)"
        }

        if ($Records.Count -eq 0) {
            Write-Host "No PST Files to Process" -ForegroundColor Yellow
            exit 0
        }
        foreach ($column in $RequiredColumns) {
            Write-Verbose "Checking for Required Column: $column"
            if (!(Test-PropertyExistance -TestObject $Records[0] -PropertyName $column)) {
                Write-Error -Message "Required Column '$($column)' is missing."
            }
        }
    }
    process {
        return $Records | Select-Object $RequiredColumns
    }

    end {
        Write-Verbose "$($Records.Count) PST Files To Process"
    }
}

<#
.SYNOPSIS
Connects to a SharePoint site using either a certificate or credentials.

.DESCRIPTION
The Connect-ToSite function connects to a SharePoint site using either a certificate or credentials. It supports two parameter sets: 'Certificate' and 'Credentials'. Depending on the parameter set used, it connects to the site using the specified client ID, tenant name, and certificate thumbprint, or using the specified client ID and credentials.

.PARAMETER SiteUrl
The URL of the SharePoint site to connect to.

.PARAMETER ClientID
The client ID used for authentication.

.PARAMETER TenantName
The tenant name used for authentication (required for the 'Certificate' parameter set).

.PARAMETER CertificateThumbprint
The thumbprint of the certificate used for authentication (required for the 'Certificate' parameter set).

.PARAMETER Credentials
The credentials used for authentication (required for the 'Credentials' parameter set).

.EXAMPLE
Connect-ToSite -SiteUrl "https://example.sharepoint.com" -ClientID "your-client-id" -TenantName "your-tenant-name" -CertificateThumbprint "your-cert-thumbprint"

This example connects to the SharePoint site using a certificate.

.EXAMPLE
Connect-ToSite -SiteUrl "https://example.sharepoint.com" -ClientID "your-client-id" -Credentials (Get-Credential)

This example connects to the SharePoint site using credentials.
#>
function Connect-ToSite {
    [CmdletBinding()]
    param
    (
        [Parameter(ParameterSetName = 'Certificate',
            Mandatory = $true,
            Position = 0)]
        [Parameter(ParameterSetName = 'Credentials',
            Mandatory = $true,
            Position = 0)]
        [string]$SiteUrl,

        [Parameter(ParameterSetName = 'Certificate',
            Mandatory = $true)]
        [Parameter(ParameterSetName = 'Credentials',
            Mandatory = $true)]
        [string]$ClientID,

        [Parameter(ParameterSetName = 'Certificate',
            Mandatory = $true)]
        [string]$TenantName,

        [Parameter(ParameterSetName = 'Certificate',
            Mandatory = $true)]
        [string]$CertificateThumbprint,

        [Parameter(ParameterSetName = 'Credentials',
            Mandatory = $true)]
        [pscredential]$Credentials
    )

    try {
        Write-Verbose "Connecting to Site $SiteUrl Using $($PSCmdlet.ParameterSetName)"
        switch ($PsCmdlet.ParameterSetName) {
            'Certificate' {
                Connect-PnPOnline -Url $SiteUrl -ClientId $ClientID -Tenant $TenantName -Thumbprint $CertificateThumbprint -WarningAction Ignore
                break
            }
            'Credentials' {
                Connect-PnPOnline -Url $SiteUrl -ClientId $ClientID -Credentials $Credentials -WarningAction Ignore
                break
            }
        }
        Write-Verbose "Connected to Site $SiteUrl"
    }
    catch {
        Write-Error -Message "Failed to connect to the site. Error: $($_)"
    }
}

<#
.SYNOPSIS
Disconnects from the currently connected SharePoint site.

.DESCRIPTION
The Disconnect-FromSite function disconnects from the currently connected SharePoint site using the Disconnect-PnPOnline cmdlet.

.EXAMPLE
Disconnect-FromSite

This example disconnects from the currently connected SharePoint site.
#>
function Disconnect-FromSite {
    Write-Verbose "Disconnecting from Site"
    Disconnect-PnPOnline -ErrorAction SilentlyContinue
    Write-Verbose "Disconnected from Site"
}

<#
.SYNOPSIS
Downloads a PST file from Azure to a specified destination path.

.DESCRIPTION
The Get-PstFileFromAzure function downloads a PST file from Azure to a specified destination path. It takes a PSCustomObject representing the PST file and a string representing the destination path as parameters.

.PARAMETER PstFile
A PSCustomObject representing the PST file to be downloaded. This object should contain at least a Name and Location property.

.PARAMETER DestinationPath
The path where the PST file will be downloaded.

.EXAMPLE
$PstFile = [PSCustomObject]@{ Name = "example.pst"; Location = "https://example.com/path/to/example.pst" }
Get-PstFileFromAzure -PstFile $PstFile -DestinationPath "C:\PSTFiles"

This example downloads the PST file from the specified URL to the C:\PSTFiles directory.

.NOTES
Ensure that you have the necessary permissions to access the Azure location and the destination path.
#>
function Get-PstFileFromAzure {
    [CmdletBinding()]
    [OutputType([string])]
    param (
        [Parameter(Mandatory = $true)]
        [PSCustomObject]$PstFile,
        [Parameter(Mandatory = $true)]
        [string]$DestinationPath
    )

    try {
        $absolutePath = [System.Uri]::new($PstFile.Location).AbsolutePath
        $FilePath = "$($DestinationPath)$($absolutePath.Replace("/", "\").Replace("%20", " "))"
        if (-not (Test-Path (Split-Path $FilePath -Parent))) {
            Write-Verbose "Creating Directory: $(Split-Path $FilePath -Parent)"
            New-Item -Path (Split-Path $FilePath -Parent) -ItemType Directory | Out-Null
        }
        Write-Verbose "Downloading PST File: $($PstFile.Name)"
        Get-PnPFile -Url $absolutePath -AsFile -Path (Split-Path $FilePath -Parent) -FileName (Split-Path $FilePath -Leaf) -Force -ErrorAction Stop
        Write-Verbose "Downloaded PST File: $($PstFile.Name) to $FilePath with size $(Get-FileSizeString -FilePath $FilePath)"
        return $FilePath
    }
    catch {
        Write-Error "Failed to download PST file: $($PstFile.Name). Error: $($_)"
    }
}

<#
.SYNOPSIS
Creates a storage context for Azure Storage operations.

.DESCRIPTION
The Get-StorageContext function creates a storage context for Azure Storage operations using the provided storage account name and key.

.PARAMETER StorageAccountName
The name of the Azure Storage account.

.PARAMETER StorageAccountKey
The key for the Azure Storage account.

.RETURNS
An object representing the storage context.

.EXAMPLE
$context = Get-StorageContext -StorageAccountName "mystorageaccount" -StorageAccountKey "myaccountkey"

This example creates a storage context for the specified Azure Storage account.
#>
function Get-StorageContext {
    [CmdletBinding()]
    [OutputType([object])]
    param (
        [Parameter(Mandatory = $true)]
        [string]$StorageAccountName,
        [Parameter(Mandatory = $true)]
        [string]$StorageAccountKey
    )

    $StorageContext = New-AzStorageContext -StorageAccountName $StorageAccountName -SasToken $StorageAccountKey
    return $StorageContext
}

<#
.SYNOPSIS
Uploads PST files to Azure Storage.

.DESCRIPTION
The Set-PstFileToAzure function uploads PST files to Azure Storage using the provided storage context and container name.

.PARAMETER PstFiles
An array of file paths representing the PST files to be uploaded.

.PARAMETER StorageContext
The storage context for Azure Storage operations.

.PARAMETER ContainerName
The name of the Azure Storage container where the PST files will be uploaded.

.EXAMPLE
$context = Get-StorageContext -StorageAccountName "mystorageaccount" -StorageAccountKey "myaccountkey"
Set-PstFileToAzure -PstFiles @("C:\PSTFiles\file1.pst", "C:\PSTFiles\file2.pst") -StorageContext $context -ContainerName "pst-container"

This example uploads the specified PST files to the specified Azure Storage container.
#>
function Set-PstFileToAzure {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [string[]]$PstFiles,
        [Parameter(Mandatory = $true)]
        [object]$StorageContext,
        [Parameter(Mandatory = $true)]
        [string]$ContainerName
    )

    Write-Verbose "Uploading $($PstFiles.Count) PST Files to Azure"
    $Status = ""
    $CurrentCount = 0
    foreach ($File in $PstFiles) {
        Write-Progress -Id 2 -Activity "Uploading PST Files" -Status "Uploading PST File $($CurrentCount) of $($PstFiles.Count)" -PercentComplete ($CurrentCount / $PstFiles.Count * 100)
        $CurrentCount++
        Write-Verbose "Uploading PST File: $File"
        try {
            $urlFilePath = $File.Replace($Script:WorkingDirectory, "").Replace("\", "/")
            Set-AzStorageBlobContent -File $File -Container $ContainerName -Context $StorageContext -Blob $urlFilePath -Force -Verbose:$false | Out-Null
            $Status = "Uploaded"
        }
        catch {
            Write-Warning "Failed to upload PST file: $File. Error: $($_). Skipping..."
            $Status = "Failed"
            continue
        }
        finally {
            Write-UploadLog -PstFileName (Split-Path $File -Leaf) -FilePath $urlFilePath -Size (Get-FileSizeString -FilePath $File) -Status $Status -CsvPath "$($Script:WorkingDirectory)\PSTFilesUploadLog.csv"
            Write-Verbose "Removing PST File: $File"
            Remove-Item -Path $File -Force -ErrorAction SilentlyContinue
        }
    }
    Write-Progress -Id 2 -Activity "Uploading PST Files" -Completed
}

<#
.SYNOPSIS
Processes PST files by downloading them from Azure and uploading them to another Azure Storage container.

.DESCRIPTION
The Invoke-ProcessPstFiles function processes PST files by downloading them from Azure and uploading them to another Azure Storage container. It takes an array of PSCustomObject representing the PST files, a destination path, storage account name, container name, and storage account key as parameters.

.PARAMETER PstFiles
An array of PSCustomObject representing the PST files to be processed.

.PARAMETER DestinationPath
The path where the PST files will be downloaded.

.PARAMETER StorageAccountName
The name of the Azure Storage account.

.PARAMETER ContainerName
The name of the Azure Storage container where the PST files will be uploaded.

.PARAMETER StorageAccountKey
The key for the Azure Storage account.

.EXAMPLE
$PstFiles = @([PSCustomObject]@{ Name = "file1.pst"; Location = "https://example.com/file1.pst" }, [PSCustomObject]@{ Name = "file2.pst"; Location = "https://example.com/file2.pst" })
Invoke-ProcessPstFiles -PstFiles $PstFiles -DestinationPath "C:\PSTFiles" -StorageAccountName "mystorageaccount" -ContainerName "pst-container" -StorageAccountKey "myaccountkey"

This example processes the specified PST files by downloading them from Azure and uploading them to the specified Azure Storage container.
#>
function Invoke-ProcessPstFiles {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [PSCustomObject[]]$PstFiles,
        [Parameter(Mandatory = $true)]
        [string]$DestinationPath,
        [Parameter(Mandatory = $true)]
        [string]$StorageAccountName,
        [Parameter(Mandatory = $true)]
        [string]$ContainerName,
        [Parameter(Mandatory = $true)]
        [string]$StorageAccountKey
    )

    $PstCount = $PstFiles.Count
    $CurrentPst = 0
    Write-Verbose "Processing $PstCount PST Files"
    $DownloadedFiles = @()

    foreach ($Pst in $PstFiles) {
        Write-Progress -Id 1 -Activity "Downloading PST Files" -Status "Downloading PST File $($CurrentPst) of $PstCount" -PercentComplete ($CurrentPst / $PstCount * 100)
        $CurrentPst++
        try {
            $DownloadedFiles += Get-PstFileFromAzure -PstFile $Pst -DestinationPath $DestinationPath
        }
        catch {
            Write-Warning "Failed to Download PST file: $($Pst.Name). Error: $($_). Skipping..."
        }
    }

    Write-Progress -Id 1 -Activity "Downloading PST Files" -Completed
    Write-Verbose "Downloaded $PstCount PST Files"

    try {
        $Context = Get-StorageContext -StorageAccountName $StorageAccountName -StorageAccountKey $StorageAccountKey
        Set-PstFileToAzure -PstFiles $DownloadedFiles -StorageContext $Context -ContainerName $ContainerName
    }
    catch {
        Write-Warning "Failed to upload PST files to Azure. Error: $($_)"
    }
}

<#
.SYNOPSIS
    Extracts and returns a list of unique URLs from an array of PSCustomObject items.

.DESCRIPTION
    The Get-UniqueUrls function takes an array of PSCustomObject items, extracts the 'Site Address' property from each item, and returns a sorted list of unique URLs. It also provides a verbose output indicating the count of unique URLs found.

.PARAMETER PstFiles
    Specifies an array of PSCustomObject items from which the 'Site Address' property will be extracted.
    This parameter is mandatory.

.OUTPUTS
    System.String[]
        Returns an array of unique URLs.

.EXAMPLE
    # Example of how to call the function with an array of PSCustomObject items
    $PstFiles = @(
        [pscustomobject]@{ 'Site Address' = 'https://example.com/site1' },
        [pscustomobject]@{ 'Site Address' = 'https://example.com/site2' },
        [pscustomobject]@{ 'Site Address' = 'https://example.com/site1' }
    )

    $uniqueUrls = Get-UniqueUrls -PstFiles $PstFiles
    $uniqueUrls

    # This example demonstrates how to call the Get-UniqueUrls function with an array of PSCustomObject items and retrieve the unique URLs.

.NOTES
    The function uses the Sort-Object -Unique cmdlet to ensure that the URLs are unique and sorted.
    The Write-Verbose cmdlet is used to provide additional information about the number of unique URLs found. To see the verbose output, run the function with the -Verbose switch.

.LINK
    Sort-Object
    Write-Verbose
#>
function Get-UniqueUrls {
    param (
        [Parameter(Mandatory = $true)]
        [PSCustomObject[]]$PstFiles
    )

    $UniqueUrls = $PstFiles.'Site Address' | Sort-Object -Unique
    Write-Verbose "Unique URLs: $($UniqueUrls.Count)"
    return $UniqueUrls
}

if (-not (Test-Path $Script:WorkingDirectory)) {
    Write-Host "Creating Working Directory: $Script:WorkingDirectory" -ForegroundColor Yellow
    New-Item -Path $Script:WorkingDirectory -ItemType Directory | Out-Null
}

Write-Host "Starting PST Copy Script" -ForegroundColor Green

$PSTFiles = Get-PSTListFromExcel -ExcelFilePath $ExcelFilePath -RequiredColumns $RequiredColumns
$UniqueUrls = Get-UniqueUrls -PstFiles $PSTFiles

$CurrentUrl = 0
foreach ($url in $UniqueUrls) {
    $absolutePath = [System.Uri]::new($url).AbsolutePath
    Write-Progress -Id 0 -Activity "Processing Sites" -Status "Processing Site: $absolutePath - ($CurrentUrl of $($UniqueUrls.Count))" -PercentComplete ($CurrentUrl / $UniqueUrls.Count * 100)
    $CurrentUrl++
    Write-Host "Processing PST files for Site: $url" -ForegroundColor Cyan
    try {
        switch ($PsCmdlet.ParameterSetName) {
            'Certificate' {
                Connect-ToSite -SiteUrl $url -ClientID $ClientID -TenantName $TenantName -CertificateThumbprint $CertificateThumbprint
                break
            }
            'Credentials' {
                Connect-ToSite -SiteUrl $url -ClientID $ClientID -Credentials $Credentials
                break
            }
        }
        $PstList = $PstFiles | Where-Object { $_.'Site Address' -match $url }
        try {
            Invoke-ProcessPstFiles -PstFiles $PstList -DestinationPath $Script:WorkingDirectory -StorageAccountName $StorageAccountName -ContainerName $StorageAccountContainer -StorageAccountKey $StorageAccountKey
        }
        catch {
            Write-Warning "Failed to process PST files for site: $url. Error: $($_)"
            continue
        }
    }
    catch {
        Write-Warning "Failed to connect to site: $url, skipping..."
        continue
    }
    finally {
        Disconnect-FromSite -ErrorAction SilentlyContinue
    }
}

Write-Progress -Id 0 -Activity "Processing Sites" -Completed
