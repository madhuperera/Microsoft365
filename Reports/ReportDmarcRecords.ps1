#Requires -Modules ExchangeOnlineManagement

<#
.SYNOPSIS
    Reports on DMARC DNS record configuration for all accepted domains in Exchange Online.

.DESCRIPTION
    Connects to Exchange Online and retrieves all accepted domains, then queries DMARC TXT records
    for each non-onmicrosoft.com domain using the specified DNS resolver.
    Outputs results to the console.

.EXAMPLE
    .\ReportDmarcRecords.ps1
#>

[CmdletBinding()]
param()

$ErrorActionPreference = 'Stop'

$S_DNSServerToUse = "1.1.1.1"

if (-not (Get-Module -Name ExchangeOnlineManagement -ListAvailable))
{
    throw "ExchangeOnlineManagement module is not installed. Install it using: Install-Module ExchangeOnlineManagement -Scope CurrentUser"
}
Connect-ExchangeOnline -ShowBanner:$false

$S_AllDomains = (Get-AcceptedDomain | Where-Object {$_.DomainName -notlike "*.onmicrosoft.com"}).DomainName

foreach ($S_Domain in $S_AllDomains)
{
    Write-Output "`nReport on $S_Domain"

    Resolve-DnsName -Type TXT -Server $S_DNSServerToUse -Name $("_dmarc." + $S_Domain)
}
