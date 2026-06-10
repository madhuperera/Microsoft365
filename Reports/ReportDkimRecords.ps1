#Requires -Modules ExchangeOnlineManagement

<#
.SYNOPSIS
    Reports on DKIM DNS record configuration for all accepted domains in Exchange Online.

.DESCRIPTION
    Connects to Exchange Online, retrieves DKIM signing configuration, and verifies DKIM DNS records
    for all non-onmicrosoft.com domains using the specified DNS resolver.
    Outputs results to the console.

.EXAMPLE
    .\ReportDkimRecords.ps1
#>

[CmdletBinding()]
param ()

$ErrorActionPreference = 'Stop'

$S_DNSServerToUse = "1.1.1.1"

if (-not (Get-Module -Name ExchangeOnlineManagement -ListAvailable))
{
    throw "ExchangeOnlineManagement module is not installed. Install it using: Install-Module ExchangeOnlineManagement -Scope CurrentUser"
}
Connect-ExchangeOnline -ShowBanner:$false

$S_DkimConfiguration = Get-DkimSigningConfig | Where-Object {$_.Domain -notlike "*.onmicrosoft.com"}

foreach ($S_DkimEntry in $S_DkimConfiguration)
{
    [String] $S_DomainName = $S_DkimEntry.Domain
    [bool] $S_DkimEnabled = $S_DkimEntry.Enabled
    Write-Output "`nReport on $S_DomainName"
    if ($S_DkimEnabled)
    {
        Write-Output "Dkim is enabled for the domain: $S_DomainName"
        try
        {
            # Selector 1
            $S_Selector1DNSName = "selector1._domainkey.$S_DomainName"
            $S_Result1 = Resolve-DnsName -Type CNAME -Server $S_DNSServerToUse -Name $S_Selector1DNSName -ErrorAction SilentlyContinue
            if ($S_Result1)
            {
                $S_Selector1DnsHostName = $S_Result1.NameHost
                if (Resolve-DnsName -Type CNAME -Server $S_DNSServerToUse -Name $S_Selector1DnsHostName -ErrorAction SilentlyContinue)
                {
                    Write-Output "Successfully verified Selector 1 for $S_DomainName"
                }
                else
                {
                    Write-Output "Error validating Selector 1 for $S_DomainName. Please rotate the Dkim Keys and try again!"
                }
            }
            else
            {
                Write-Output "Error validating DNS for $S_Selector1DNSName. Please either update $S_DNSServerToUse or check your DNS Provider"
            }

            # Selector 2
            $S_Selector2DNSName = "selector2._domainkey.$S_DomainName"
            $S_Result2 = Resolve-DnsName -Type CNAME -Server $S_DNSServerToUse -Name $S_Selector2DNSName -ErrorAction SilentlyContinue
            if ($S_Result2)
            {
                $S_Selector2DnsHostName = $S_Result2.NameHost
                if (Resolve-DnsName -Type CNAME -Server $S_DNSServerToUse -Name $S_Selector2DnsHostName -ErrorAction SilentlyContinue)
                {
                    Write-Output "Successfully verified Selector 2 for $S_DomainName"
                }
                else
                {
                    Write-Output "Error validating Selector 2 for $S_DomainName. Please rotate the Dkim Keys and try again!"
                }
            }
            else
            {
                Write-Output "Error validating DNS for $S_Selector2DNSName. Please either update $S_DNSServerToUse or check your DNS Provider"
            }
        }
        catch
        {
            Write-Output "Something went wrong"
        }
    }
    else
    {
        Write-Output "Dkim is NOT enabled for the domain: $S_DomainName"
    }
}