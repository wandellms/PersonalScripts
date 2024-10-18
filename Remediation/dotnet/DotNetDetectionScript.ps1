$minimumVersions = @(
    [System.Version]"6.0.35",
    [System.Version]"8.0.10"
)


$NetCoreDirectories = Get-ChildItem -Path "$env:ProgramFiles\dotnet\shared\Microsoft.NetCore.App" -Directory -ErrorAction SilentlyContinue | Where-Object { $_.Name -notlike "*preview*" }
$NetCoreDirectories += Get-ChildItem -Path "$(${env:ProgramFiles(x86)})\dotnet\shared\Microsoft.NetCore.App" -Directory -ErrorAction SilentlyContinue | Where-Object { $_.Name -notlike "*preview*" }

[System.Version[]]$versionFolders = $NetCoreDirectories.Name | ForEach-Object { [System.Version]$_ }

function Remove-EmptyNetCoreDirectories {

    foreach ($directory in $NetCoreDirectories) {
        if ($null -eq (Get-ChildItem -Path $directory.FullName)) {
            Write-Host "Removing Empty Directory: $($directory.FullName)"
            Remove-Item -Path $directory.FullName
            $NetCoreDirectories = $NetCoreDirectories | Where-Object { $_.Name -ne $directory.Name }
        }
    }
}

Remove-EmptyNetCoreDirectories

foreach ($Version in $versionFolders) {
    $matchingVersion = $minimumVersions | Where-Object { $_.Major -eq $Version.Major }
    if ($matchingVersion.Build -gt $Version.Build) {
        Write-Host "Minimum Version Not Met - Current Version $Version - Minimum Version $matchingVersion"
        exit 1
    }
}

Write-Host "All Versions Meet Minimum Requirements"
exit 0