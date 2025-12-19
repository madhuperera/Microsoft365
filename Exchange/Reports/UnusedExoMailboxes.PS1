# UnusedExoMailboxes.PS1
# Modified from: https://github.com/12Knocksinna/Office365itpros/blob/master/ReportUnusedExoMailboxes.PS1
# Original script by 12Knocksinna - adapted for custom requirements
# Find and report unused Exchange Online mailboxes
# Needs to connect to Exchange Online and the Microsoft Graph PowerShell SDK

param
(
    [bool]$ReportInExcel = $false,
    [bool]$RequiredGridView = $false
)

$ReportNameTitle = "Unused Mailboxes Report"
$ReportWorksheetName = "UnusedMailboxes"
$ReportOutputName = (Get-Date -Format "yyyy-MM-dd HHmm") + "_" + "UnusedMailboxesReport"

# Function to convert dates to New Zealand time

try
{
  # Connect to the Microsoft Graph PowerShell SDK so that we can read sign in data
  Connect-MgGraph -Scopes User.Read.All, AuditLog.Read.All -NoWelcome -ErrorAction Stop
}
catch
{
  Write-Host "Failed to connect to Microsoft Graph. " -ForegroundColor Red
  exit
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
$Now = Get-Date -format s
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
        $DaysSinceActive = (New-TimeSpan -Start $LastActiveDateTime -End $Now).Days 
    }
  
    # Get Mailbox statistics
    $Stats = (Get-ExoMailboxStatistics -Identity $M.DistinguishedName)
    $MbxSize = ($Stats.TotalItemSize.Value.ToString()).Split("(")[0] 
    # Get last Sign in from Entra ID sign in logs
    $LastUserSignIn = $null
    $LastUserSignIn = (Get-MgAuditLogSignIn -Filter "UserId eq '$($M.ExternalDirectoryObjectId)'" -Top 1).CreatedDateTime
    Start-Sleep -Milliseconds 250 # To avoid throttling from Graph
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
  Write-Host "`n`nDisplaying report in Grid View..."
  $Report | Sort-Object DaysSinceActive -Descending | Out-GridView
}

if ($ReportInExcel) 
{
    If (Get-Module ImportExcel -ListAvailable) 
    { 
        Import-Module ImportExcel -ErrorAction SilentlyContinue 
        $ExcelOutputFile = ((New-Object -ComObject Shell.Application).Namespace('shell:Downloads').Self.Path) + "\$ReportOutputName.xlsx" 
        $Report | Export-Excel -Path $ExcelOutputFile -WorksheetName $ReportWorksheetName -Title ("$ReportNameTitle {0}" -f (Get-Date -format 'dd-MMM-yyyy')) -TitleBold -TableName $ReportWorksheetName
        $OutputFile = $ExcelOutputFile 
    }
    else 
    { 
        $CSVOutputFile = ((New-Object -ComObject Shell.Application).Namespace('shell:Downloads').Self.Path) + "\$ReportOutputName.csv" 
        $Report | Export-Csv -Path $CSVOutputFile -NoTypeInformation -Encoding Utf8 
        $Outputfile = $CSVOutputFile 
    } 
}
else
{
    $CSVOutputFile = ((New-Object -ComObject Shell.Application).Namespace('shell:Downloads').Self.Path) + "\$ReportOutputName.csv" 
    $Report | Export-Csv -Path $CSVOutputFile -NoTypeInformation -Encoding Utf8 
    $Outputfile = $CSVOutputFile 
}
Write-Host ("Output data is available in {0}" -f $OutputFile)
