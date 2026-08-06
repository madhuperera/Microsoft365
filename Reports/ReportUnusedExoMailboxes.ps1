#Requires -Modules ExchangeOnlineManagement, Microsoft.Graph.Authentication, Microsoft.Graph.Users

<#
.SYNOPSIS
    Finds and reports Exchange Online mailboxes that have not been actively used.

.DESCRIPTION
    Connects to both Exchange Online and Microsoft Graph to identify user mailboxes
    that show no recent activity. Reports mailbox statistics, last logon times from
    Exchange Online diagnostic logs, and last Entra ID sign-in information.
    Exports results to CSV.

.PARAMETER ReportInExcel
    When set to $true, attempts to export to Excel using the ImportExcel module.
    Falls back to CSV if the module is not available.

.PARAMETER RequiredGridView
    When set to $true, displays the report in an Out-GridView window.

.PARAMETER OutputPath
    Path for the output file (without extension). Defaults to a timestamped file in the current directory.

.EXAMPLE
    .\ReportUnusedExoMailboxes.ps1

.EXAMPLE
    .\ReportUnusedExoMailboxes.ps1 -RequiredGridView $true

.EXAMPLE
    .\ReportUnusedExoMailboxes.ps1 -OutputPath "C:\Reports\UnusedMailboxes"
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [bool]$ReportInExcel = $false,

    [Parameter(Mandatory = $false)]
    [bool]$RequiredGridView = $false,

    [Parameter(Mandatory = $false)]
    [ValidateNotNullOrEmpty()]
    [string]$OutputPath
)

$ErrorActionPreference = 'Stop'

$S_RequiredGraphScopes = @(
    'User.Read.All'
    'AuditLog.Read.All'
)

$S_GraphRequestDelayMilliseconds = 5

if (-not $OutputPath)
{
    $S_Timestamp = Get-Date -Format 'yyyy-MM-dd_HHmm'
    $OutputPath = Join-Path -Path (Get-Location).Path -ChildPath "ReportUnusedExoMailboxes_$S_Timestamp"
}

$S_ReportNameTitle = "Unused Mailboxes Report"
$S_ReportWorksheetName = "UnusedMailboxes"

try
{
  # Connect to the Microsoft Graph PowerShell SDK so that we can read sign in data
  $S_ExistingContext = Get-MgContext
  if ($S_ExistingContext)
  {
    Write-Host "Existing Graph session detected:" -ForegroundColor Yellow
    Write-Host "  Account : $($S_ExistingContext.Account)" -ForegroundColor Yellow
    Write-Host "  TenantId: $($S_ExistingContext.TenantId)" -ForegroundColor Yellow
    Write-Host "  Scopes  : $($S_ExistingContext.Scopes -join ', ')" -ForegroundColor Yellow
    Write-Host ""

    $S_Choice = Read-Host "Use existing session? [Y] Yes  [N] Disconnect and reconnect  (Default: Y)"
    if ($S_Choice -eq 'N')
    {
      Disconnect-MgGraph | Out-Null
      Connect-MgGraph -Scopes $S_RequiredGraphScopes -NoWelcome -ErrorAction Stop
    }
  }
  else
  {
    Connect-MgGraph -Scopes $S_RequiredGraphScopes -NoWelcome -ErrorAction Stop
  }
}
catch
{
  Write-Host "Failed to connect to Microsoft Graph. " -ForegroundColor Red
  exit
}

$S_ActiveContext = Get-MgContext
Write-Host ""
Write-Host "Active Graph context:" -ForegroundColor Cyan
Write-Host "  Account    : $($S_ActiveContext.Account)" -ForegroundColor Cyan
Write-Host "  TenantId   : $($S_ActiveContext.TenantId)" -ForegroundColor Cyan
Write-Host "  Environment: $($S_ActiveContext.Environment)" -ForegroundColor Cyan
Write-Host "  Scopes     : $($S_ActiveContext.Scopes -join ', ')" -ForegroundColor Cyan
Write-Host ""

$S_ContextConfirmation = Read-Host "Proceed with this Graph context? [Y] Yes  [N] No  (Default: N)"
if ([string]::IsNullOrWhiteSpace($S_ContextConfirmation))
{
    $S_ContextConfirmation = 'N'
}
else
{
    $S_ContextConfirmation = $S_ContextConfirmation.ToUpperInvariant()
}
if ($S_ContextConfirmation -ne 'Y')
{
    throw "Operation cancelled. Please reconnect to the correct tenant and account, then run again."
}

try
{
  # Check for Exchange Online
  $S_ModulesLoaded = Get-Module | Select-Object Name
  If (!($S_ModulesLoaded -match "ExchangeOnlineManagement")) 
  {
    Write-Host "Loading Exchange Online PowerShell module" -ForegroundColor Yellow
    Connect-ExchangeOnline -ShowBanner:$False -ErrorAction Stop
  }
}
catch
{
  Write-Host "Failed to connect to Exchange Online. " -ForegroundColor Red
  exit
}


# Find mailboxes and check if they are unused
$S_Now = Get-Date -format s
[int]$S_I = 0
Write-Host "Looking for User Mailboxes..."
[array]$S_Mbx = Get-ExoMailbox -RecipientTypeDetails UserMailbox -ResultSize Unlimited | `
    Select-Object DisplayName, DistinguishedName, UserPrincipalName, ExternalDirectoryObjectId | Sort-Object DisplayName
  Write-Host ("Reporting {0} mailboxes..." -f $S_Mbx.Count)
  $S_Report = [System.Collections.Generic.List[Object]]::new() 
  ForEach ($S_M in $S_Mbx) 
  {
    $S_I++  
    Write-Host ("`n`nProcessing {0} {1}/{2}" -f $S_M.DisplayName, $S_I, $S_Mbx.count) 
    $S_LastActive = $Null
    $S_Log = Export-MailboxDiagnosticLogs -Identity $S_M.DistinguishedName -ExtendedProperties
    $S_Xml = [xml]($S_Log.MailboxLog) 
    $S_LastEMail = $Null; $S_LastCalendar = $Null; $S_LastContacts = $Null; $S_LastFile = $Null
    $S_LastEmail = ($S_Xml.Properties.MailboxTable.Property | Where-Object {$_.Name -eq "LastEmailTimeCurrentValue"}).Value
    $S_LastCalendar = ($S_Xml.Properties.MailboxTable.Property | Where-Object {$_.Name -eq "LastCalendarTimeCurrentValue"}).Value
    $S_LastContacts = ($S_Xml.Properties.MailboxTable.Property | Where-Object {$_.Name -eq "LastContactsTimeCurrentValue"}).Value
    $S_LastFile = ($S_Xml.Properties.MailboxTable.Property | Where-Object {$_.Name -eq "LastFileTimeCurrentValue"}).Value
    $S_LastLogonTime = ($S_Xml.Properties.MailboxTable.Property | Where-Object {$_.Name -eq "LastLogonTime"}).Value 
    $S_LastActive = ($S_Xml.Properties.MailboxTable.Property | Where-Object {$_.Name -eq "LastUserActionWorkloadAggregateTime"}).Value 
    
    # This massaging of dates is to accommodate the different U.S. date format returned by Export-MailboxDiagnosticsData
    [datetime]$S_LastActiveDateTime = Get-Date
    If ([string]::IsNullOrEmpty($S_LastActive)) 
    {
        $S_DaysSinceActive = "N/A"
    }
    If (($S_LastActive.IndexOf("M") -gt -0)) { # U.S. format date with AM or PM in it
        $S_LastActiveDateTime = [datetime]$S_LastActive
    } Else {
        $S_LastActiveDateTime = Get-Date ($S_LastActive) 
    }
    If ($S_LastActiveDateTime) 
    {
        $S_DaysSinceActive = (New-TimeSpan -Start $S_LastActiveDateTime -End $S_Now).Days 
    }
  
    # Get Mailbox statistics
    $S_Stats = (Get-ExoMailboxStatistics -Identity $S_M.DistinguishedName)
    $S_MbxSize = ($S_Stats.TotalItemSize.Value.ToString()).Split("(")[0] 
    # Get last Sign in from Entra ID sign in logs
    $S_LastUserSignIn = $null
    $S_LastUserSignIn = (Get-MgAuditLogSignIn -Filter "UserId eq '$($S_M.ExternalDirectoryObjectId)'" -Top 1).CreatedDateTime
    Start-Sleep -Milliseconds $S_GraphRequestDelayMilliseconds # To avoid throttling from Graph
    If ($S_LastUserSignIn) 
    {
       $S_LastUserSignInDate = Get-Date($S_LastUserSignIn) -format g 
    } 
    else
    {
       $S_LastUserSignInDate = "No sign in records found in last 30 days" 
    }
    # Get account enabled status
    $S_AccountEnabled = (Get-MgUser -UserId $S_M.ExternalDirectoryObjectId -Property AccountEnabled).AccountEnabled
    $S_ReportLine = [PSCustomObject][Ordered]@{ 
        Mailbox         = $S_M.DisplayName 
        UPN             = $S_M.UserPrincipalName
        Enabled         = $S_AccountEnabled
        Items           = $S_Stats.ItemCount 
        Size            = $S_MbxSize 
        LastLogonExo    = $S_LastLogonTime
        LastLogonAD     = $S_LastUserSignInDate
        DaysSinceActive = $S_DaysSinceActive
        LastActive      = $S_LastActive
        LastEmail       = $S_LastEmail
        LastCalendar    = $S_LastCalendar
        LastContacts    = $S_LastContacts
        LastFile        = $S_LastFile } 
    $S_Report.Add($S_ReportLine) 
  } 

if ($RequiredGridView)
{
  Write-Host "`n`nDisplaying report in Grid View..."
  $S_Report | Sort-Object DaysSinceActive -Descending | Out-GridView
}

if ($ReportInExcel) 
{
    If (Get-Module ImportExcel -ListAvailable) 
    { 
        Import-Module ImportExcel -ErrorAction SilentlyContinue 
        $S_ExcelOutputFile = "$OutputPath.xlsx"
        $S_Report | Export-Excel -Path $S_ExcelOutputFile -WorksheetName $S_ReportWorksheetName -Title ("$S_ReportNameTitle {0}" -f (Get-Date -format 'dd-MMM-yyyy')) -TitleBold -TableName $S_ReportWorksheetName
        $S_OutputFile = $S_ExcelOutputFile 
    }
    else 
    { 
        $S_CSVOutputFile = "$OutputPath.csv"
        $S_Report | Export-Csv -Path $S_CSVOutputFile -NoTypeInformation -Encoding Utf8 
        $S_OutputFile = $S_CSVOutputFile 
    } 
}
else
{
    $S_CSVOutputFile = "$OutputPath.csv"
    $S_Report | Export-Csv -Path $S_CSVOutputFile -NoTypeInformation -Encoding Utf8 
    $S_OutputFile = $S_CSVOutputFile 
}
Write-Host ("Output data is available in {0}" -f $S_OutputFile)

$S_DisconnectChoice = Read-Host "`nDisconnect from Microsoft Graph? [Y] Yes  [N] Keep session  (Default: N)"
if ($S_DisconnectChoice -eq 'Y')
{
    Disconnect-MgGraph | Out-Null
    Write-Host "Disconnected from Microsoft Graph." -ForegroundColor Yellow
}
else
{
    Write-Host "Graph session kept alive." -ForegroundColor Green
}
