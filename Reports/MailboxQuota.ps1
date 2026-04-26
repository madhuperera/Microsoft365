# Set threshold % of quota to use as warning level
param
(
    [int]$Threshold = 85,
    [bool]$ReportInExcel = $false
)

$ReportNameTitle = "Mailbox Quota Report"
$ReportWorksheetName = "MailboxQuota"
$ReportOutputName = (Get-Date -Format "yyyy-MM-dd HHmm") + "_" + "MailboxQuotaReport"

Write-Host "Finding mailboxes..." 
[array]$Mbx = Get-ExoMailbox -RecipientTypeDetails UserMailbox -PropertySet Quota -Properties DisplayName -ResultSize Unlimited 
$Report = [System.Collections.Generic.List[Object]]::new() # Create output file for report 
ForEach ($M in $Mbx) 
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
    $Report.Add($ReportLine) 
}  
# Export to CSV 
$Report | Sort-Object Mailbox | Export-csv -NoTypeInformation MailboxQuotaReport.csv 

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

