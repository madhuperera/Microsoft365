#Requires -Modules Microsoft.Graph.Authentication, Microsoft.Graph.Users

<#
.SYNOPSIS
    Exports Entra ID users who have a manager assigned.

.DESCRIPTION
    Connects to Microsoft Graph and exports a list of users who have a manager assigned,
    including their email address and manager information.

.NOTES
    Author: Madhu Perera
    Requires: Microsoft.Graph.Authentication, Microsoft.Graph.Users modules

.EXAMPLE
    .\ReportUsersWithManagers.ps1
#>

$S_RequiredGraphScopes = @(
    'User.Read.All'
    'Directory.Read.All'
)

$S_GraphRequestDelayMilliseconds = 5

# Install required module if not already installed
# Install-Module Microsoft.Graph.Users -Scope CurrentUser -Force

# Connect to Microsoft Graph
Connect-MgGraph -Scopes $S_RequiredGraphScopes

# Get all users with their manager information
Write-Host "Retrieving licensed and enabled users from Entra ID..." -ForegroundColor Cyan

$users = Get-MgUser -All -Filter "assignedLicenses/`$count ne 0 and accountEnabled eq true" -ConsistencyLevel eventual -CountVariable Records -Property Id, DisplayName, UserPrincipalName, Mail, JobTitle, Department

$usersWithManagers = @()

foreach ($user in $users) {
    Write-Progress -Activity "Processing users" -Status "Checking $($user.DisplayName)" -PercentComplete (($usersWithManagers.Count / $users.Count) * 100)
    
    # Get manager for each user
    try {
        Start-Sleep -Milliseconds $S_GraphRequestDelayMilliseconds
        $manager = Get-MgUserManager -UserId $user.Id -ErrorAction SilentlyContinue
        
        if ($manager) {
            $usersWithManagers += [PSCustomObject]@{
                DisplayName        = $user.DisplayName
                UserPrincipalName  = $user.UserPrincipalName
                Email              = if ($user.Mail) { $user.Mail } else { $user.UserPrincipalName }
                JobTitle           = $user.JobTitle
                Department         = $user.Department
                ManagerName        = $manager.AdditionalProperties.displayName
                ManagerEmail       = $manager.AdditionalProperties.userPrincipalName
            }
        }
    }
    catch {
        # User has no manager assigned, skip
        continue
    }
}

Write-Progress -Activity "Processing users" -Completed

# Export results
$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$exportPath = ".\ReportUsersWithManagers_$timestamp.csv"

$usersWithManagers | Export-Csv -Path $exportPath -NoTypeInformation -Encoding UTF8

Write-Host "`nExport complete!" -ForegroundColor Green
Write-Host "Total users with managers: $($usersWithManagers.Count)" -ForegroundColor Yellow
Write-Host "File saved to: $exportPath" -ForegroundColor Yellow

# Disconnect from Microsoft Graph
Disconnect-MgGraph

# Display sample of results
Write-Host "`nSample of exported data:" -ForegroundColor Cyan
$usersWithManagers | Select-Object -First 5 | Format-Table -AutoSize
