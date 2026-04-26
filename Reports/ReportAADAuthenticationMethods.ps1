#Requires -Modules Microsoft.Graph.Authentication, Microsoft.Graph.Users

<#
.SYNOPSIS
    Reports Windows Hello for Business authentication method registrations for enabled Entra ID users.

.DESCRIPTION
    Connects to Microsoft Graph and retrieves all enabled users. For each user, checks whether
    they have a Windows Hello for Business authentication method registered and exports the results
    to a CSV file.

.PARAMETER S_ReportFilePath
    Path for the output CSV file. Defaults to a timestamped file in the current directory.

.EXAMPLE
    .\ReportAADAuthenticationMethods.ps1

.EXAMPLE
    .\ReportAADAuthenticationMethods.ps1 -S_ReportFilePath "C:\Reports\WHfBReport.csv"
#>

[CmdletBinding()]
param
(
    [Parameter(Mandatory = $false)]
    [string] $S_ReportFilePath
)

$ErrorActionPreference = 'Stop'

if (-not $S_ReportFilePath)
{
    $S_Timestamp = Get-Date -Format 'yyyyMMdd_HHmmss'
    $S_ReportFilePath = Join-Path -Path (Get-Location).Path -ChildPath "ReportAADAuthenticationMethods_$S_Timestamp.csv"
}

$S_RequiredGraphScopes = @(
    'UserAuthenticationMethod.Read.All'
)

$S_GraphRequestDelayMilliseconds = 5

# Connect to Microsoft Graph
$S_ExistingContext = Get-MgContext
if (-not $S_ExistingContext)
{
    Connect-MgGraph -Scopes $S_RequiredGraphScopes -NoWelcome
}

# Get all enabled users in Entra ID
$AllUsers = Get-MgUser -Filter "accountEnabled eq true" -All
$AllUsers = $AllUsers | Sort-Object -Property DisplayName

$AllData = @()

# Loop through each user and retrieve their authentication methods
foreach ($Member in $AllUsers)
{
    $MemberId = $Member.Id
    $MemberName = $Member.DisplayName
    Write-Host "`nWindows Hello for Business Check for $($MemberName):"
    Start-Sleep -Milliseconds $S_GraphRequestDelayMilliseconds
    $WHfBAuthMethods = Get-MgUserAuthenticationMethod -UserId $MemberId | Where-Object {$_.additionalProperties.'@odata.type' -like "*windowsHelloForBusinessAuthenticationMethod*"}
    if ($WHfBAuthMethods)
    {
        foreach ($Method in $WHfBAuthMethods)
        {
            $DeviceName = $Method.additionalProperties.displayName
            $RegisteredDate = $Method.additionalProperties.createdDateTime
            Write-Host "Found on $DeviceName"

            $OBJ = New-Object PSObject
            $OBJ | Add-Member -MemberType NoteProperty -Name "StaffDisplayName" -Value $MemberName
            $OBJ | Add-Member -MemberType NoteProperty -Name "DeviceName" -Value $DeviceName
            $OBJ | Add-Member -MemberType NoteProperty -Name "RegisteredDate" -Value $RegisteredDate

            $AllData += $OBJ
        }
    }
    else
    {
        Write-Host "No Windows Hello for Business Registrations found!`n"
    }
}

$AllData | Export-Csv -Path $S_ReportFilePath -NoTypeInformation
Write-Host "Report exported to: $S_ReportFilePath" -ForegroundColor Green
