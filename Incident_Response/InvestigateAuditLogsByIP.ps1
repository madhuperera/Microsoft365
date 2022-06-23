[cmdletbinding()]

param (
    [Parameter(
            Mandatory=$true)]
            [string]$AttackerIPAddress,
    [Parameter(
            Mandatory=$true)]
            [String]$FromDate,
    [Parameter(
            Mandatory=$true)]
            [String]$ToDate,
    [Parameter(
            Mandatory=$true)]
            [String]$StaffUPN,
    [Parameter(
            Mandatory=$true)]
            [String]$Operation,
    [Parameter(
        Mandatory=$true)]
        [String] $FilePath
)
# -----------------------------------------------------------------------------------------------------------------------------------------------


#-----------------------------------------------------------[Function]---------------------------------------------------------------------------


# -----------------------------------------------------------------------------------------------------------------------------------------------
$Records = (Search-UnifiedAuditLog -StartDate $FromDate -EndDate $ToDate -Operations $Operation `
                 -UserId $StaffUPN -IPAddress $AttackerIPAddress -ResultSize 5000 -ErrorAction Stop)

If ($Records.Count -eq 0)
{
    Write-Output "No Records found"
}
else
{
    # ------------------------- UPDATE | Create ----------------------------------------
    If ($Operation -eq "Update" -or $Operation -eq "Create")
    {
        $AuditLogs = foreach ($Entry in $Records)
        {
            $AuditData = ConvertFrom-Json -InputObject $Entry.AuditData
            $ModifiedItem = $AuditData.Item
            $Attachments = $ModifiedItem.Attachments
            $Subject = $ModifiedItem.Subject
            $ParentFolder = $ModifiedItem.ParentFolder.Path

            $Props = @{
                'EmailFolder'   = $ParentFolder;
                'EmailSubject'  = $Subject;
                'Attachments'   = $Attachments;
                'CreationDate'  = $Entry.CreationDate
        }
        New-Object -TypeName PSObject -Property $Props
        }
    }
    
    # ------------------------- MoveToDeletedItems | HardDelete | SoftDelete ----------------------------------------

    If ($Operation -eq "MoveToDeletedItems" -or $Operation -eq "HardDelete" -or $Operation -eq "SoftDelete")
    {
        $AuditLogs = foreach ($Entry in $Records)
        {
            $AuditData = ConvertFrom-Json -InputObject $Entry.AuditData
            $ModifiedItem = $AuditData.AffectedItems
            $Subject = $ModifiedItem.Subject
            $ParentFolder = $ModifiedItem.ParentFolder.Path

            $Props = @{
                'EmailFolder'   = $ParentFolder;
                'EmailSubject'  = $Subject;
                'CreationDate'  = $Entry.CreationDate
        }
        New-Object -TypeName PSObject -Property $Props
        }
    }

     # ------------------------- Move ----------------------------------------

     If ($Operation -eq "Move")
     {
         $AuditLogs = foreach ($Entry in $Records)
         {
             $AuditData = ConvertFrom-Json -InputObject $Entry.AuditData
             $ModifiedItem = $AuditData.AffectedItems
             $Subject = $ModifiedItem.Subject
             $SourceFolder = $AuditData.Folder.Path
             $DestinationFolder = $AuditData.DestFolder.Path
 
             $Props = @{
                 'DestinationFolder'   = $DestinationFolder;
                 'SourceFolder'     = $SourceFolder;
                 'EmailSubject'  = $Subject;
                 'CreationDate'  = $Entry.CreationDate
         }
         New-Object -TypeName PSObject -Property $Props
         }
     }

    # ------------------------- New-InboxRule ----------------------------------------
    If ($Operation -eq "New-InboxRule")
    {
        $AuditLogs = foreach ($Entry in $Records)
        {
            $AuditData = ConvertFrom-Json -InputObject $Entry.AuditData
            $RuleParameters = $AuditData.Parameters
            $Props = @{
                'RuleContent'   = $RuleParameters;
                'CreationDate'  = $Entry.CreationDate
        }
        New-Object -TypeName PSObject -Property $Props
        }
    }



    $ReportTime = Get-Date -Format "MM-dd-yyyy_HH-mm-ss"

    try
    {
        $CSVFileName = $Operation + "_" + $ReportTime + ".csv"
        $CSVFilePath = $FilePath + "\" + $CSVFileName
        $AuditLogs | Export-CSV -LiteralPath $CSVFilePath -Encoding Unicode -NoTypeInformation -Delimiter "`t" -ErrorAction Stop
    }
    catch 
    {
        Write-Error -Message $_.Exception
    }
}



