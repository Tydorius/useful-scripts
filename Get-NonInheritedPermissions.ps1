#Requires -RunAsAdministrator

<#
.SYNOPSIS
    Gets non-inherited permissions for files and folders.

.DESCRIPTION
    This script retrieves non-inherited permissions for all files and folders within a specified root path.
    It requires administrative privileges to run and outputs the results to a CSV file.

.PARAMETER rootPath
    The root path to start scanning for non-inherited permissions.
    If not provided, the script will prompt for the path.
.Notes
    Author:   Tydorius
    Version:  1.0
    Date:     December 3, 2025
#>
param(
    [Parameter(Mandatory=$false, HelpMessage="Enter the root path to scan.")]
    [string]$rootPath
)

# Check for Administrator Privileges
if (-not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Write-Error "This script must be run with Administrator privileges." -ErrorAction Stop
}

# If rootPath is not provided, prompt the user
if ([string]::IsNullOrEmpty($rootPath)) {
    $rootPath = Read-Host -Prompt "Please enter the root path to scan (e.g., D:\)"
}

# Validate the path
if (-not (Test-Path -Path $rootPath)) {
    Write-Error "The path '$rootPath' does not exist." -ErrorAction Stop
}

$outputFile = Join-Path -Path $PSScriptRoot -ChildPath "NonInheritedPermissions.csv"
$results = @()

# Get all files and folders recursively, then get their ACLs
Get-ChildItem -Path $rootPath -Recurse -ErrorAction SilentlyContinue | ForEach-Object {
    $itemPath = $_.FullName
    try {
        $acl = Get-Acl -Path $itemPath -ErrorAction Stop
        $nonInheritedPermissions = $acl.Access | Where-Object { -not $_.IsInherited }

        if ($nonInheritedPermissions) {
            foreach ($permission in $nonInheritedPermissions) {
                $results += [PSCustomObject]@{
                    FullPath          = $itemPath
                    FileName          = $_.Name
                    IdentityReference = $permission.IdentityReference
                    FileSystemRights  = $permission.FileSystemRights
                    AccessControlType = $permission.AccessControlType
                }
            }
        }
    }
    catch {
        Write-Warning "Could not access ACL for: $itemPath. Error: $_"
    }
}

# Export the results to a CSV file
if ($results) {
    $results | Export-Csv -Path $outputFile -NoTypeInformation
    Write-Host "Non-inherited permissions report saved to: $outputFile"
}
else {
    Write-Host "No non-inherited permissions found in the specified path."
}
