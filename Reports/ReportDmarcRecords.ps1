#Requires -Modules ExchangeOnlineManagement

<#
.SYNOPSIS
    Reports on DMARC DNS record configuration for all accepted domains in Exchange Online.

.DESCRIPTION
    Retrieves all accepted domains from Exchange Online and queries DMARC TXT records
    for each non-onmicrosoft.com domain using the specified DNS resolver.
    Outputs results to the console.

.EXAMPLE
    .\ReportDmarcRecords.ps1
#>

[CmdletBinding()]
param ()

$ErrorActionPreference = 'Stop'

$S_DNSServerToUse = "1.1.1.1"
$AllDomains = (Get-AcceptedDomain | Where-Object {$_.DomainName -notlike "*.onmicrosoft.com"}).DomainName

foreach ($Domain in $AllDomains)
{
    Write-Output "`nReport on $Domain"

    Resolve-DnsName -Type TXT -Server $S_DNSServerToUse -Name $("_dmarc." + $Domain)
}