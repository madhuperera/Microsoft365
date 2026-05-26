#Requires -Modules ExchangeOnlineManagement

<#
.SYNOPSIS
    Reports on mailbox quota usage for all user mailboxes in Exchange Online.

.DESCRIPTION
    Connects to Exchange Online and retrieves mailbox quota statistics for all user mailboxes.
    Calculates the percentage of quota used and flags mailboxes that exceed the warning threshold.
    Exports results to CSV. Optionally exports to Excel using the ImportExcel module.

.PARAMETER Threshold
    Percentage of quota usage to use as the warning level. Defaults to 85.

.PARAMETER ReportInExcel
    When set to $true, attempts to export to Excel using the ImportExcel module.
    Falls back to CSV if the module is not available.

.PARAMETER OutputPath
    Path for the output CSV or Excel file. Defaults to a timestamped file in the current directory.

.EXAMPLE
    .\ReportMailboxQuota.ps1

.EXAMPLE
    .\ReportMailboxQuota.ps1 -Threshold 90

.EXAMPLE
    .\ReportMailboxQuota.ps1 -OutputPath "C:\Reports\MailboxQuota.csv"
#>

[CmdletBinding()]
param (
    [Parameter(Mandatory = $false)]
    [ValidateRange(1, 100)]
    [int]$Threshold = 85,

    [Parameter(Mandatory = $false)]
    [bool]$ReportInExcel = $false,

    [Parameter(Mandatory = $false)]
    [ValidateNotNullOrEmpty()]
    [string]$OutputPath
)

$ErrorActionPreference = 'Stop'

if (-not $OutputPath)
{
    $S_Timestamp = Get-Date -Format 'yyyy-MM-dd_HHmm'
    $OutputPath = Join-Path -Path (Get-Location).Path -ChildPath "ReportMailboxQuota_$S_Timestamp"
}

$S_ReportNameTitle = "Mailbox Quota Report"
$S_ReportWorksheetName = "MailboxQuota"

if (-not (Get-Module -Name ExchangeOnlineManagement -ListAvailable))
{
    throw "ExchangeOnlineManagement module is not installed. Install it using: Install-Module ExchangeOnlineManagement -Scope CurrentUser"
}
Connect-ExchangeOnline -ShowBanner:$false

Write-Host "Finding mailboxes..." 
[array]$S_Mbx = Get-ExoMailbox -RecipientTypeDetails UserMailbox -PropertySet Quota -Properties DisplayName -ResultSize Unlimited 
$S_Report = [System.Collections.Generic.List[Object]]::new() # Create output file for report 
ForEach ($M in $S_Mbx) 
{ 
    # Find current usage 
    Write-Host "`n`nProcessing" $M.DisplayName 
    $ErrorText = $Null 
    $MbxStats = Get-ExoMailboxStatistics -Identity $M.UserPrincipalName | Select-Object ItemCount, TotalItemSize 
    # Return byte count of quota used 
    [INT64]$QuotaUsed = [convert]::ToInt64(((($MbxStats.TotalItemSize.ToString().split("(")[1]).split(")")[0]).split(" ")[0] -replace '[,]', '')) 
    # Byte count for mailbox quota 
    [INT64]$MbxQuota = [convert]::ToInt64(((($M.ProhibitSendReceiveQuota.ToString().split("(")[1]).split(")")[0]).split(" ")[0] -replace '[,]', '')) 
    $MbxQuotaGB = [math]::Round(($MbxQuota / 1GB), 2) 
    $QuotaPercentUsed = [math]::Round(($QuotaUsed / $MbxQuota), 4).ToString("P") 
    $QuotaUsedGB = [math]::Round(($QuotaUsed / 1GB), 2) 
    If ($QuotaPercentUsed -gt $Threshold)
    { 
        Write-Host $M.DisplayName "current mailbox use is above threshold at" $QuotaPercentUsed -Foregroundcolor Red 
        $ErrorText = "Mailbox quota over $QuotaPercentUsed %" 
    } 
    
    # Generate report line for the mailbox 
    $ReportLine = [PSCustomObject]@{  
        Mailbox          = $M.DisplayName  
        MbxQuota         = $MbxQuotaGB 
        Items            = $MbxStats.ItemCount 
        MbxSizeGB        = $QuotaUsedGB 
        QuotaPercentUsed = $QuotaPercentUsed 
        ErrorText        = $ErrorText
    }  
    $S_Report.Add($ReportLine) 
}  

if ($ReportInExcel) 
{
    If (Get-Module ImportExcel -ListAvailable) 
    { 
        Import-Module ImportExcel -ErrorAction SilentlyContinue 
        $ExcelOutputFile = "$OutputPath.xlsx"
        $S_Report | Sort-Object Mailbox | Export-Excel -Path $ExcelOutputFile -WorksheetName $S_ReportWorksheetName -Title ("$S_ReportNameTitle {0}" -f (Get-Date -format 'dd-MMM-yyyy')) -TitleBold -TableName $S_ReportWorksheetName
        $OutputFile = $ExcelOutputFile 
    }
    else 
    { 
        $CSVOutputFile = "$OutputPath.csv"
        $S_Report | Sort-Object Mailbox | Export-Csv -Path $CSVOutputFile -NoTypeInformation -Encoding Utf8 
        $Outputfile = $CSVOutputFile 
    } 
}
else
{
    $CSVOutputFile = "$OutputPath.csv"
    $S_Report | Sort-Object Mailbox | Export-Csv -Path $CSVOutputFile -NoTypeInformation -Encoding Utf8 
    $Outputfile = $CSVOutputFile 
}
Write-Host ("Output data is available in {0}" -f $OutputFile)

