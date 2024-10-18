$minimumVersions = @(
    [System.Version]"6.0.35",
    [System.Version]"8.0.10"
)

$Uninstaller = "https://github.com/dotnet/cli-lab/releases/download/1.7.521001/dotnet-core-uninstall-1.7.521001.msi"
$versions = @("6.0", "8.0");
$NetCoreDirectories = Get-ChildItem -Path "$env:ProgramFiles\dotnet\shared\Microsoft.NetCore.App" -Directory -ErrorAction SilentlyContinue | Where-Object { $_.Name -notlike "*preview*" }
$NetCoreDirectories += Get-ChildItem -Path "$(${env:ProgramFiles(x86)})\dotnet\shared\Microsoft.NetCore.App" -Directory -ErrorAction SilentlyContinue | Where-Object { $_.Name -notlike "*preview*" }
[System.Version[]]$versionFolders = $NetCoreDirectories.Name | ForEach-Object { [System.Version]$_ }

<#
.SYNOPSIS
Generates a temporary file path for a given file name.

.DESCRIPTION
This function uses the .NET method System.IO.Path.GetTempPath() to get the path to the temporary files directory, and then appends the provided file name to this path.

.PARAMETER FileName
The name of the file for which to generate a temporary file path.

.EXAMPLE
Get-TempFilePath -FileName "temp.txt"

This will return a string representing a path to a file named "temp.txt" in the temporary files directory.

.NOTES
The returned path does not guarantee that the file exists, it only provides a valid path in the temporary directory.
#>
function Get-TempFilePath {
    param (
        [Parameter(Mandatory = $true)]
        [string]$FileName
    )
    $TempDirectory = [System.IO.Path]::GetTempPath()
    return (Join-Path -Path $TempDirectory -ChildPath $FileName)
}

<#
.SYNOPSIS
Downloads a file from a given URL and starts a process with the downloaded file.

.DESCRIPTION
This function downloads a file from the provided URL to a temporary location, and then starts a process with the downloaded file using the provided arguments.

.PARAMETER DownloadURL
The URL from which to download the file.

.PARAMETER FileName
The name to give to the downloaded file.

.PARAMETER Arguments
The arguments to pass to the process started with the downloaded file.

.EXAMPLE
Download-And-StartProcess -DownloadURL "https://example.com/file.zip" -FileName "file.zip" -Arguments "/s"

This will download the file from "https://example.com/file.zip", save it as "file.zip" in the temporary files directory, and start a process with this file using the "/s" argument.

.NOTES
The process is started in a wait state, meaning that the function will not return until the process has exited.
#>
function Invoke-DownloadAndStartProcess {
    param (
        [string]$DownloadURL,
        [string]$FileName,
        [string]$Arguments
    )
    $ProgressPreference = 'SilentlyContinue'
    $TempFilePath = Get-TempFilePath -FileName $FileName
    Invoke-WebRequest -Uri $DownloadURL -OutFile $TempFilePath
    $ProgressPreference = 'Continue'
    Start-Process $TempFilePath -ArgumentList $Arguments -Wait
}

<#
.SYNOPSIS
    Retrieves the direct download URL for the specified version of the .NET runtime installer.

.DESCRIPTION
    The Get-InstallerUrl function retrieves the direct download URL for the specified version of the .NET runtime installer.
    It uses the Microsoft website to find the URL by parsing the HTML content of the download page.

.PARAMETER version
    The version of the .NET runtime for which to retrieve the installer URL.

.EXAMPLE
    Get-InstallerUrl -version "5.0.3"
    Retrieves the direct download URL for the .NET runtime version 5.0.3.

.INPUTS
    None. You cannot pipe objects to this function.

.OUTPUTS
    System.String
    The direct download URL for the specified version of the .NET runtime installer.

.NOTES
    This function requires an internet connection to access the Microsoft website and retrieve the installer URL.
#>
function Get-InstallerUrl($version) {
    $VersionRootUrl = "https://dotnet.microsoft.com/en-us/download/dotnet/thank-you/runtime-desktop-$version-windows-x64-installer"

    $content = Invoke-WebRequest -Uri $VersionRootUrl -UseBasicParsing
    return ($content.Links | Where-Object { $_.id -eq "directLink" }).href
}

<#
.SYNOPSIS
    Retrieves the version number from a specified URL.

.DESCRIPTION
    The Find-Version function retrieves the version number from a specified URL by making an HTTP request and parsing the response content.

.PARAMETER url
    The URL from which to retrieve the version number.

.OUTPUTS
    System.String
    The version number retrieved from the URL.

.EXAMPLE
    Find-Version -url "https://example.com"
    Retrieves the version number from the specified URL.

#>
function Find-Version($url) {
    $content = Invoke-WebRequest -Uri $url -UseBasicParsing
    $pattern = '<button[^>]*aria-controls="version_0"[^>]*>(.*?)</button>'
    if ($content -match $pattern) {
        $version = $Matches[1]
    }
    return $version
}

<#
.SYNOPSIS
Removes empty .NET Core directories.

.DESCRIPTION
The Remove-EmptyNetCoreDirectories function takes an array of DirectoryInfo objects representing .NET Core directories and removes those that are empty.

.PARAMETER NetCoreDirectories
An array of DirectoryInfo objects representing the .NET Core directories to be checked and potentially removed.

.EXAMPLE
$directories = Get-ChildItem -Path "C:\Projects" -Directory
Remove-EmptyNetCoreDirectories -NetCoreDirectories $directories

This example retrieves all directories in the "C:\Projects" path and removes any that are empty.
#>
function Remove-EmptyNetCoreDirectories {
    param(
        [Parameter(Mandatory = $true)]
        [System.IO.DirectoryInfo[]]$NetCoreDirectories
    )

    foreach ($directory in $NetCoreDirectories) {
        if ($null -eq (Get-ChildItem -Path $directory.FullName)) {
            Remove-Item -Path $directory.FullName -ErrorAction SilentlyContinue
        }
    }
}

Remove-EmptyNetCoreDirectories -NetCoreDirectories $NetCoreDirectories

Invoke-DownloadAndStartProcess -DownloadURL $Uninstaller -FileName "dotnet-core-uninstall.msi" -Arguments "/quiet /norestart"

foreach ($version in $versions) {
    $matchingVersion = $minimumVersions | Where-Object { $_.Major -eq ([System.Version]$Version).Major }
    $latest = Find-Version -url "https://dotnet.microsoft.com/en-us/download/dotnet/$version"

    $url = Get-InstallerUrl -version($latest)
    $fileName = Split-Path $url -Leaf

    Invoke-DownloadAndStartProcess -DownloadURL $url -FileName $fileName -Arguments "/install /quiet /norestart"

    & "C:\Program Files (x86)\dotnet-core-uninstall\dotnet-core-uninstall.exe" remove --all-below $matchingVersion.ToString() --runtime --yes
}

foreach ($Version in $versionFolders) {
    $matchingVersion = $minimumVersions | Where-Object { $_.Major -eq $Version.Major }
    if ($matchingVersion.Build -gt $Version.Build) {
        $currentVersionFolder = $NetCoreDirectories | Where-Object { $_.Name -eq $Version.ToString() }
        Remove-Item -Path $currentVersionFolder.FullName -Recurse -Force -Confirm:$false
    }
}