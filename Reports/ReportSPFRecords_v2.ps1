#Requires -Modules ExchangeOnlineManagement

<#
.SYNOPSIS
    Reports on SPF DNS record configuration for all accepted domains in Exchange Online.

.DESCRIPTION
    Connects to Exchange Online and retrieves all accepted domains, then queries SPF TXT records
    for each non-onmicrosoft.com domain.

    Cross-platform variant (_v2): DNS lookups use DNS-over-HTTPS (DoH) via Google and Cloudflare
    instead of the Windows-only Resolve-DnsName cmdlet, so the script runs on macOS, Linux and
    Windows PowerShell 7+.

    Outputs results to the console.

.EXAMPLE
    .\ReportSPFRecords_v2.ps1
#>

[CmdletBinding()]
param()

$ErrorActionPreference = 'Stop'

$S_DohTimeoutSeconds = 10

# ---------------------------------------------------------------------------
# Helper: cross-platform DNS resolution via DNS-over-HTTPS (DoH)
# Returns an object with .Records (Resolve-DnsName-like), .Server and .Error.
# ---------------------------------------------------------------------------
function Resolve-DnsRecordDoH
{
    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory = $true)]
        [string]$Name,

        [Parameter(Mandatory = $true)]
        [ValidateSet('A', 'AAAA', 'CNAME', 'MX', 'NS', 'TXT', 'SOA')]
        [string]$Type
    )

    # DNS record type numbers used to interpret the DoH JSON "Answer" entries.
    $F_TypeNumToName = @{ 1 = 'A'; 2 = 'NS'; 5 = 'CNAME'; 6 = 'SOA'; 15 = 'MX'; 16 = 'TXT'; 28 = 'AAAA' }

    # DoH providers tried in order. Google takes a plain query; Cloudflare needs the JSON accept header.
    $F_Providers = @(
        [pscustomobject]@{ Name = 'dns.google';         Uri = 'https://dns.google/resolve';          Headers = @{} }
        [pscustomobject]@{ Name = 'cloudflare-dns.com'; Uri = 'https://cloudflare-dns.com/dns-query'; Headers = @{ 'accept' = 'application/dns-json' } }
    )

    foreach ($F_Provider in $F_Providers)
    {
        try
        {
            $F_Query = '{0}?name={1}&type={2}' -f $F_Provider.Uri, [uri]::EscapeDataString($Name), $Type
            $F_Response = Invoke-RestMethod -Uri $F_Query -Headers $F_Provider.Headers `
                -TimeoutSec $S_DohTimeoutSeconds -ErrorAction Stop

            # Status 0 = NOERROR, 3 = NXDOMAIN (a valid "no record" answer). Anything else: try next provider.
            if ($null -eq $F_Response.Status) { continue }
            if ($F_Response.Status -ne 0 -and $F_Response.Status -ne 3) { continue }

            $F_Records = New-Object System.Collections.Generic.List[object]
            foreach ($F_Answer in @($F_Response.Answer))
            {
                if (-not $F_Answer) { continue }
                $F_AnsTypeName = $F_TypeNumToName[[int]$F_Answer.type]
                if ($F_AnsTypeName -ne $Type) { continue }

                $F_Data = [string]$F_Answer.data
                $F_Record = [ordered]@{
                    Name         = [string]$F_Answer.name
                    Type         = $F_AnsTypeName
                    TTL          = [int]$F_Answer.TTL
                    NameHost     = $null
                    NameExchange = $null
                    Preference   = $null
                    Strings      = $null
                    IPAddress    = $null
                }

                switch ($F_AnsTypeName)
                {
                    'CNAME' { $F_Record.NameHost = $F_Data.TrimEnd('.') }
                    'NS'    { $F_Record.NameHost = $F_Data.TrimEnd('.') }
                    'SOA'   { $F_Record.NameHost = $F_Data }
                    'MX'    {
                        $F_Parts = $F_Data.Split(' ', 2)
                        if ($F_Parts.Count -eq 2)
                        {
                            $F_Record.Preference   = [int]$F_Parts[0]
                            $F_Record.NameExchange = $F_Parts[1].TrimEnd('.')
                        }
                        else
                        {
                            $F_Record.NameExchange = $F_Data.TrimEnd('.')
                        }
                    }
                    'TXT'   {
                        # DoH returns TXT data wrapped in quotes; long records arrive as
                        # multiple quoted chunks ("a" "b"). Reassemble to the raw value.
                        $F_Clean = $F_Data -replace '^"', '' -replace '"$', ''
                        $F_Clean = $F_Clean -replace '" "', ''
                        $F_Record.Strings = @($F_Clean)
                    }
                    default { $F_Record.IPAddress = $F_Data }
                }

                $F_Records.Add([pscustomobject]$F_Record)
            }

            return [pscustomobject]@{
                Records = $F_Records.ToArray()
                Server  = $F_Provider.Name
                Error   = $null
            }
        }
        catch
        {
            continue
        }
    }

    return [pscustomobject]@{
        Records = @()
        Server  = $null
        Error   = ('No response from DoH resolvers for {0} {1}' -f $Type, $Name)
    }
}

if (-not (Get-Module -Name ExchangeOnlineManagement -ListAvailable))
{
    throw "ExchangeOnlineManagement module is not installed. Install it using: Install-Module ExchangeOnlineManagement -Scope CurrentUser"
}
Connect-ExchangeOnline -ShowBanner:$false

$S_AllDomains = (Get-AcceptedDomain | Where-Object {$_.DomainName -notlike "*.onmicrosoft.com"}).DomainName

foreach ($S_Domain in $S_AllDomains)
{
    Write-Output "`nReport on $S_Domain"

    $S_SpfLookup = Resolve-DnsRecordDoH -Type TXT -Name $S_Domain
    $S_SpfRecords = $S_SpfLookup.Records | Where-Object { ($_.Strings -join '') -like "*spf*" }

    if ($S_SpfRecords)
    {
        $S_SpfRecords | ForEach-Object { Write-Output ($_.Strings -join '') }
    }
    elseif ($S_SpfLookup.Error)
    {
        Write-Output $S_SpfLookup.Error
    }
    else
    {
        Write-Output "No SPF TXT record found for $S_Domain"
    }
}
