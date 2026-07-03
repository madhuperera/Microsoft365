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
  $ModulesLoaded = Get-Module | Select-Object Name
  If (!($ModulesLoaded -match "ExchangeOnlineManagement")) 
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
[int]$i = 0
Write-Host "Looking for User Mailboxes..."
[array]$Mbx = Get-ExoMailbox -RecipientTypeDetails UserMailbox -ResultSize Unlimited | `
    Select-Object DisplayName, DistinguishedName, UserPrincipalName, ExternalDirectoryObjectId | Sort-Object DisplayName
  Write-Host ("Reporting {0} mailboxes..." -f $Mbx.Count)
  $Report = [System.Collections.Generic.List[Object]]::new() 
  ForEach ($M in $Mbx) 
  {
    $i++  
    Write-Host ("`n`nProcessing {0} {1}/{2}" -f $M.DisplayName, $i, $Mbx.count) 
    $LastActive = $Null
    $Log = Export-MailboxDiagnosticLogs -Identity $M.DistinguishedName -ExtendedProperties
    $xml = [xml]($Log.MailboxLog) 
    $LastEMail = $Null; $LastCalendar = $Null; $LastContacts = $Null; $LastFile = $Null
    $LastEmail = ($xml.Properties.MailboxTable.Property | Where-Object {$_.Name -eq "LastEmailTimeCurrentValue"}).Value
    $LastCalendar = ($xml.Properties.MailboxTable.Property | Where-Object {$_.Name -eq "LastCalendarTimeCurrentValue"}).Value
    $LastContacts = ($xml.Properties.MailboxTable.Property | Where-Object {$_.Name -eq "LastContactsTimeCurrentValue"}).Value
    $LastFile = ($xml.Properties.MailboxTable.Property | Where-Object {$_.Name -eq "LastFileTimeCurrentValue"}).Value
    $LastLogonTime = ($xml.Properties.MailboxTable.Property | Where-Object {$_.Name -eq "LastLogonTime"}).Value 
    $LastActive = ($xml.Properties.MailboxTable.Property | Where-Object {$_.Name -eq "LastUserActionWorkloadAggregateTime"}).Value 
    
    # This massaging of dates is to accommodate the different U.S. date format returned by Export-MailboxDiagnosticsData
    [datetime]$LastActiveDateTime = Get-Date
    If ([string]::IsNullOrEmpty($LastActive)) 
    {
        $DaysSinceActive = "N/A"
    }
    If (($LastActive.IndexOf("M") -gt -0)) { # U.S. format date with AM or PM in it
        $LastActiveDateTime = [datetime]$LastActive
    } Else {
        $LastActiveDateTime = Get-Date ($LastActive) 
    }
    If ($LastActiveDateTime) 
    {
        $DaysSinceActive = (New-TimeSpan -Start $LastActiveDateTime -End $S_Now).Days 
    }
  
    # Get Mailbox statistics
    $Stats = (Get-ExoMailboxStatistics -Identity $M.DistinguishedName)
    $MbxSize = ($Stats.TotalItemSize.Value.ToString()).Split("(")[0] 
    # Get last Sign in from Entra ID sign in logs
    $LastUserSignIn = $null
    $LastUserSignIn = (Get-MgAuditLogSignIn -Filter "UserId eq '$($M.ExternalDirectoryObjectId)'" -Top 1).CreatedDateTime
    Start-Sleep -Milliseconds $S_GraphRequestDelayMilliseconds # To avoid throttling from Graph
    If ($LastUserSignIn) 
    {
       $LastUserSignInDate = Get-Date($LastUserSignIn) -format g 
    } 
    else
    {
       $LastUserSignInDate = "No sign in records found in last 30 days" 
    }
    # Get account enabled status
    $AccountEnabled = (Get-MgUser -UserId $M.ExternalDirectoryObjectId -Property AccountEnabled).AccountEnabled
    $ReportLine = [PSCustomObject][Ordered]@{ 
        Mailbox         = $M.DisplayName 
        UPN             = $M.UserPrincipalName
        Enabled         = $AccountEnabled
        Items           = $Stats.ItemCount 
        Size            = $MbxSize 
        LastLogonExo    = $LastLogonTime
        LastLogonAD     = $LastUserSignInDate
        DaysSinceActive = $DaysSinceActive
        LastActive      = $LastActive
        LastEmail       = $LastEmail
        LastCalendar    = $LastCalendar
        LastContacts    = $LastContacts
        LastFile        = $LastFile } 
    $Report.Add($ReportLine) 
  } 

if ($RequiredGridView)
{
  # Out-GridView only exists on Windows PowerShell / with the GUI tools installed.
  if (Get-Command Out-GridView -ErrorAction SilentlyContinue)
  {
    Write-Host "`n`nDisplaying report in Grid View..."
    $Report | Sort-Object DaysSinceActive -Descending | Out-GridView
  }
  else
  {
    Write-Host "`n`nOut-GridView is not available on this platform; skipping grid view. Use -ReportInExcel or the CSV output instead." -ForegroundColor Yellow
  }
}

if ($ReportInExcel) 
{
    If (Get-Module ImportExcel -ListAvailable) 
    { 
        Import-Module ImportExcel -ErrorAction SilentlyContinue 
        $ExcelOutputFile = "$OutputPath.xlsx"
        $Report | Export-Excel -Path $ExcelOutputFile -WorksheetName $S_ReportWorksheetName -Title ("$S_ReportNameTitle {0}" -f (Get-Date -format 'dd-MMM-yyyy')) -TitleBold -TableName $S_ReportWorksheetName
        $OutputFile = $ExcelOutputFile 
    }
    else 
    { 
        $CSVOutputFile = "$OutputPath.csv"
        $Report | Export-Csv -Path $CSVOutputFile -NoTypeInformation -Encoding Utf8 
        $Outputfile = $CSVOutputFile 
    } 
}
else
{
    $CSVOutputFile = "$OutputPath.csv"
    $Report | Export-Csv -Path $CSVOutputFile -NoTypeInformation -Encoding Utf8 
    $Outputfile = $CSVOutputFile 
}
Write-Host ("Output data is available in {0}" -f $OutputFile)

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
