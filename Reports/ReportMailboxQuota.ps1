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

$S_Threshold = $Threshold
$S_ReportInExcel = $ReportInExcel
$S_OutputPath = $OutputPath

if (-not $S_OutputPath)
{
    $S_Timestamp = Get-Date -Format 'yyyy-MM-dd_HHmm'
    $S_OutputPath = Join-Path -Path (Get-Location).Path -ChildPath "ReportMailboxQuota_$S_Timestamp"
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
ForEach ($S_MailboxEntry in $S_Mbx)
{
    # Find current usage
    Write-Host "`n`nProcessing" $S_MailboxEntry.DisplayName
    $S_ErrorText = $null
    $S_MbxStats = Get-ExoMailboxStatistics -Identity $S_MailboxEntry.UserPrincipalName | Select-Object ItemCount, TotalItemSize

    # Return byte count of quota used
    [int64]$S_QuotaUsed = [convert]::ToInt64(((($S_MbxStats.TotalItemSize.ToString().split("(")[1]).split(")")[0]).split(" ")[0] -replace '[,]', ''))

    # Byte count for mailbox quota
    [int64]$S_MbxQuota = [convert]::ToInt64(((($S_MailboxEntry.ProhibitSendReceiveQuota.ToString().split("(")[1]).split(")")[0]).split(" ")[0] -replace '[,]', ''))
    $S_MbxQuotaGB = [math]::Round(($S_MbxQuota / 1GB), 2)
    $S_QuotaPercentUsed = [math]::Round(($S_QuotaUsed / $S_MbxQuota), 4).ToString("P")
    $S_QuotaUsedGB = [math]::Round(($S_QuotaUsed / 1GB), 2)

    if ($S_QuotaPercentUsed -gt $S_Threshold)
    {
        Write-Host $S_MailboxEntry.DisplayName "current mailbox use is above threshold at" $S_QuotaPercentUsed -ForegroundColor Red
        $S_ErrorText = "Mailbox quota over $S_QuotaPercentUsed %"
    }

    # Generate report line for the mailbox
    $S_ReportLine = [PSCustomObject]@{
        Mailbox          = $S_MailboxEntry.DisplayName
        MbxQuota         = $S_MbxQuotaGB
        Items            = $S_MbxStats.ItemCount
        MbxSizeGB        = $S_QuotaUsedGB
        QuotaPercentUsed = $S_QuotaPercentUsed
        ErrorText        = $S_ErrorText
    }
    $S_Report.Add($S_ReportLine)
}

if ($S_ReportInExcel)
{
    if (Get-Module ImportExcel -ListAvailable)
    {
        Import-Module ImportExcel -ErrorAction SilentlyContinue
        $S_ExcelOutputFile = "$S_OutputPath.xlsx"
        $S_Report | Sort-Object Mailbox | Export-Excel -Path $S_ExcelOutputFile -WorksheetName $S_ReportWorksheetName -Title ("$S_ReportNameTitle {0}" -f (Get-Date -Format 'dd-MMM-yyyy')) -TitleBold -TableName $S_ReportWorksheetName
        $S_OutputFile = $S_ExcelOutputFile
    }
    else
    {
        $S_CsvOutputFile = "$S_OutputPath.csv"
        $S_Report | Sort-Object Mailbox | Export-Csv -Path $S_CsvOutputFile -NoTypeInformation -Encoding Utf8
        $S_OutputFile = $S_CsvOutputFile
    }
}
else
{
    $S_CsvOutputFile = "$S_OutputPath.csv"
    $S_Report | Sort-Object Mailbox | Export-Csv -Path $S_CsvOutputFile -NoTypeInformation -Encoding Utf8
    $S_OutputFile = $S_CsvOutputFile
}

Write-Host ("Output data is available in {0}" -f $S_OutputFile)

