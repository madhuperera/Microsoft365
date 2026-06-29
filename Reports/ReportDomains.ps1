#Requires -Modules Microsoft.Graph.Authentication, Microsoft.Graph.Identity.DirectoryManagement

<#
.SYNOPSIS
    Generates a Microsoft 365 Domain and DNS HTML report for the connected tenant.

.DESCRIPTION
    Read-only reporting script. Connects to Microsoft Graph and discovers all
    verified domains in the current Microsoft 365 tenant. For every domain
    (including the tenant's *.onmicrosoft.com namespace) it queries DNS to
    gather:

      * Authoritative name servers (used to infer the DNS hosting provider)
      * SPF (TXT v=spf1) records
      * MX records with priority

    For tenant-owned (non-onmicrosoft.com) domains it additionally collects:

      * DKIM CNAME records for Microsoft 365 selector1 and selector2
      * DMARC (_dmarc TXT) records, with policy state

    DNS records are resolved against the domain's own authoritative name
    servers first. If the authoritative lookup fails, the script falls back
    to Cloudflare (1.1.1.1) and then Google (8.8.8.8). DNS lookup failures
    do not stop the script; they are recorded in the report as review items.

    The script does not make any tenant or DNS changes.

    A standalone HTML report is written to a Reports/Output sub-folder of
    the script directory, with a timestamped filename.

    Required Microsoft Graph delegated scopes (read-only):
        Domain.Read.All
        Organization.Read.All

.PARAMETER OutputPath
    Folder or HTML file path for the report. When a folder is supplied, a
    timestamped filename is generated automatically. Defaults to the
    current working directory.

.PARAMETER IncludeOnMicrosoftDomains
    When specified, *.onmicrosoft.com domains are also included in the DKIM
    and DMARC sections of the report. They are always included in the
    Domain Inventory, SPF and MX sections.

.PARAMETER SkipRegistrarLookup
    When specified, RDAP and WHOIS port-43 lookups for the Registrar
    column are skipped. Useful on locked-down hosts that block outbound
    HTTPS to RDAP endpoints or TCP/43 to WHOIS servers.

.EXAMPLE
    .\ReportDomains.ps1

    Connects to Microsoft Graph (or reuses an existing context), discovers
    tenant domains, performs DNS checks and writes a timestamped HTML
    report under .\Output.

.EXAMPLE
    .\ReportDomains.ps1 -OutputPath 'C:\Reports'

    Writes the timestamped HTML report into C:\Reports.

.NOTES
    Read-only. No tenant or DNS changes are made.
    Required modules: Microsoft.Graph.Authentication,
                      Microsoft.Graph.Identity.DirectoryManagement
    Resolve-DnsName (Windows DnsClient module) is required for DNS lookups.
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [ValidateNotNullOrEmpty()]
    [string]$OutputPath,

    [Parameter(Mandatory = $false)]
    [switch]$IncludeOnMicrosoftDomains,

    [Parameter(Mandatory = $false)]
    [switch]$SkipRegistrarLookup
)

$ErrorActionPreference = 'Stop'

# ---------------------------------------------------------------------------
# Script-level configuration
# ---------------------------------------------------------------------------
$S_RequiredGraphScopes = @(
    'Domain.Read.All'
    'Organization.Read.All'
)

$S_GraphRequestDelayMilliseconds   = 5
$S_PublicResolvers                 = @('1.1.1.1', '8.8.8.8')
$S_DnsQueryTimeoutSeconds          = 5
$S_RequireGraphContextConfirmation = $true

# Registrar lookup configuration
$S_SkipRegistrarLookup    = $false
$S_RdapBootstrapUrl       = 'https://data.iana.org/rdap/dns.json'
$S_RdapTimeoutSeconds     = 10
$S_WhoisTimeoutSeconds    = 8
$S_WhoisPort              = 43
# Per-TLD WHOIS port-43 fallback servers used when RDAP is unavailable.
# Only TLDs with a known stable port-43 endpoint should be listed here.
$S_WhoisServerOverrides   = @{
    'nz' = 'whois.srs.net.nz'
}

$S_Timestamp = Get-Date -Format 'yyyyMMdd_HHmmss'
$S_ScriptRoot = if ($PSScriptRoot) { $PSScriptRoot } else { (Get-Location).Path }

if ($SkipRegistrarLookup)
{
    $S_SkipRegistrarLookup = $true
}

# Cache for the IANA RDAP bootstrap document so we only download it once per run.
$S_RdapBootstrapCache = $null

# ---------------------------------------------------------------------------
# Helper: look up the RDAP base URL for a TLD via the IANA bootstrap document
# ---------------------------------------------------------------------------
function Get-RdapBaseUrlForTld
{
    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory = $true)]
        [string]$Tld
    )

    if (-not $script:S_RdapBootstrapCache)
    {
        try
        {
            $script:S_RdapBootstrapCache = Invoke-RestMethod -Uri $S_RdapBootstrapUrl `
                -Method Get -TimeoutSec $S_RdapTimeoutSeconds -ErrorAction Stop
        }
        catch
        {
            Write-Verbose ('IANA RDAP bootstrap fetch failed: {0}' -f $_.Exception.Message)
            return $null
        }
    }

    $F_Tld = $Tld.ToLowerInvariant()
    foreach ($F_Service in $script:S_RdapBootstrapCache.services)
    {
        $F_Tlds = @($F_Service[0])
        $F_Urls = @($F_Service[1])
        if ($F_Tlds -contains $F_Tld -and $F_Urls.Count -gt 0)
        {
            return ($F_Urls[0].TrimEnd('/'))
        }
    }
    return $null
}

# ---------------------------------------------------------------------------
# Helper: query RDAP for a domain and extract the registrar entity name
# ---------------------------------------------------------------------------
function Get-RegistrarFromRdap
{
    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory = $true)]
        [string]$DomainName
    )

    $F_Tld = ($DomainName -split '\.')[-1]
    $F_Base = Get-RdapBaseUrlForTld -Tld $F_Tld
    if (-not $F_Base)
    {
        return [pscustomobject]@{
            Registrar = $null
            Source    = 'RDAP'
            Error     = ('No RDAP service found for .{0}' -f $F_Tld)
        }
    }

    $F_Url = ('{0}/domain/{1}' -f $F_Base, $DomainName)
    try
    {
        $F_Response = Invoke-RestMethod -Uri $F_Url -Method Get `
            -TimeoutSec $S_RdapTimeoutSeconds -ErrorAction Stop
    }
    catch
    {
        return [pscustomobject]@{
            Registrar = $null
            Source    = 'RDAP'
            Error     = ('RDAP request failed for {0}: {1}' -f $DomainName, $_.Exception.Message)
        }
    }

    $F_Registrar = $null
    if ($F_Response.entities)
    {
        foreach ($F_Entity in $F_Response.entities)
        {
            if ($F_Entity.roles -and ($F_Entity.roles -contains 'registrar'))
            {
                # Prefer the structured fn property from vCard array if present
                if ($F_Entity.vcardArray -and $F_Entity.vcardArray.Count -ge 2)
                {
                    foreach ($F_VCardItem in $F_Entity.vcardArray[1])
                    {
                        if ($F_VCardItem.Count -ge 4 -and $F_VCardItem[0] -eq 'fn')
                        {
                            $F_Registrar = [string]$F_VCardItem[3]
                            break
                        }
                    }
                }

                if (-not $F_Registrar -and $F_Entity.handle)
                {
                    $F_Registrar = [string]$F_Entity.handle
                }
                break
            }
        }
    }

    if (-not $F_Registrar)
    {
        return [pscustomobject]@{
            Registrar = $null
            Source    = 'RDAP'
            Error     = ('RDAP response had no registrar entity for {0}' -f $DomainName)
        }
    }

    return [pscustomobject]@{
        Registrar = $F_Registrar
        Source    = 'RDAP'
        Error     = $null
    }
}

# ---------------------------------------------------------------------------
# Helper: legacy WHOIS port-43 lookup, used as a fallback for TLDs not in
# the IANA RDAP bootstrap (notably .nz). Returns the registrar name only.
# ---------------------------------------------------------------------------
function Get-RegistrarFromWhois
{
    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory = $true)]
        [string]$DomainName,

        [Parameter(Mandatory = $true)]
        [string]$WhoisServer
    )

    $F_Client = $null
    $F_Stream = $null
    $F_Reader = $null
    try
    {
        $F_Client = New-Object System.Net.Sockets.TcpClient
        $F_Client.SendTimeout    = $S_WhoisTimeoutSeconds * 1000
        $F_Client.ReceiveTimeout = $S_WhoisTimeoutSeconds * 1000

        $F_AsyncResult = $F_Client.BeginConnect($WhoisServer, $S_WhoisPort, $null, $null)
        if (-not $F_AsyncResult.AsyncWaitHandle.WaitOne([TimeSpan]::FromSeconds($S_WhoisTimeoutSeconds)))
        {
            $F_Client.Close()
            return [pscustomobject]@{
                Registrar = $null
                Source    = ('WHOIS:{0}' -f $WhoisServer)
                Error     = ('WHOIS connect timed out after {0}s' -f $S_WhoisTimeoutSeconds)
            }
        }
        $F_Client.EndConnect($F_AsyncResult)

        $F_Stream = $F_Client.GetStream()
        $F_QueryBytes = [System.Text.Encoding]::ASCII.GetBytes(("{0}`r`n" -f $DomainName))
        $F_Stream.Write($F_QueryBytes, 0, $F_QueryBytes.Length)

        $F_Reader = New-Object System.IO.StreamReader($F_Stream, [System.Text.Encoding]::ASCII)
        $F_Body = $F_Reader.ReadToEnd()
    }
    catch
    {
        return [pscustomobject]@{
            Registrar = $null
            Source    = ('WHOIS:{0}' -f $WhoisServer)
            Error     = ('WHOIS lookup failed for {0}: {1}' -f $DomainName, $_.Exception.Message)
        }
    }
    finally
    {
        if ($F_Reader) { $F_Reader.Dispose() }
        if ($F_Stream) { $F_Stream.Dispose() }
        if ($F_Client) { $F_Client.Close() }
    }

    if ([string]::IsNullOrWhiteSpace($F_Body))
    {
        return [pscustomobject]@{
            Registrar = $null
            Source    = ('WHOIS:{0}' -f $WhoisServer)
            Error     = 'WHOIS response was empty'
        }
    }

    # Common registrar field labels across registries.
    $F_Patterns = @(
        '(?im)^\s*Registrar\s+Name\s*:\s*(?<v>.+)$',
        '(?im)^\s*Registrar\s*:\s*(?<v>.+)$',
        '(?im)^\s*Sponsoring\s+Registrar\s*:\s*(?<v>.+)$'
    )
    foreach ($F_Pattern in $F_Patterns)
    {
        $F_Match = [regex]::Match($F_Body, $F_Pattern)
        if ($F_Match.Success)
        {
            return [pscustomobject]@{
                Registrar = $F_Match.Groups['v'].Value.Trim()
                Source    = ('WHOIS:{0}' -f $WhoisServer)
                Error     = $null
            }
        }
    }

    return [pscustomobject]@{
        Registrar = $null
        Source    = ('WHOIS:{0}' -f $WhoisServer)
        Error     = 'No registrar field found in WHOIS response'
    }
}

# ---------------------------------------------------------------------------
# Helper: combined RDAP-first, WHOIS-fallback registrar lookup
# ---------------------------------------------------------------------------
function Get-DomainRegistrar
{
    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory = $true)]
        [string]$DomainName
    )

    $F_Tld = ($DomainName -split '\.')[-1].ToLowerInvariant()

    # Try RDAP first
    $F_Rdap = Get-RegistrarFromRdap -DomainName $DomainName
    if ($F_Rdap.Registrar)
    {
        return $F_Rdap
    }

    # Fall back to known WHOIS server for this TLD (e.g. .nz)
    if ($S_WhoisServerOverrides.ContainsKey($F_Tld))
    {
        $F_Whois = Get-RegistrarFromWhois -DomainName $DomainName -WhoisServer $S_WhoisServerOverrides[$F_Tld]
        if ($F_Whois.Registrar)
        {
            return $F_Whois
        }
        return [pscustomobject]@{
            Registrar = $null
            Source    = $F_Whois.Source
            Error     = ('{0}; WHOIS fallback: {1}' -f $F_Rdap.Error, $F_Whois.Error)
        }
    }

    return $F_Rdap
}

# ---------------------------------------------------------------------------
# Helper: infer DNS hosting provider from name server hostnames
# ---------------------------------------------------------------------------
function Get-DnsHostingProvider
{
    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory = $true)]
        [AllowNull()]
        [string[]]$NameServers
    )

    if (-not $NameServers -or $NameServers.Count -eq 0)
    {
        return 'N/A'
    }

    $F_Sample = ($NameServers | Select-Object -First 1).ToLowerInvariant()

    $F_Map = [ordered]@{
        'cloudflare.com'        = 'Cloudflare'
        'azure-dns.'            = 'Azure DNS'
        'awsdns-'               = 'AWS Route 53'
        'domaincontrol.com'     = 'GoDaddy'
        'godaddy.com'           = 'GoDaddy'
        'dnsmadeeasy.com'       = 'DNS Made Easy'
        'dnsimple.com'          = 'DNSimple'
        'nsone.net'             = 'NS1'
        'ultradns.'             = 'UltraDNS'
        'akam.net'              = 'Akamai'
        'akamaitech.net'        = 'Akamai'
        'name-services.com'     = 'Enom'
        'registrar-servers.com' = 'Namecheap'
        'namecheaphosting.com'  = 'Namecheap'
        'wixdns.net'            = 'Wix'
        'squarespacedns.com'    = 'Squarespace'
        'shopify.com'           = 'Shopify'
        'crazydomains.com'      = 'Crazy Domains'
        'syd1.zonomi.com'       = 'Zonomi'
        'zonomi.com'            = 'Zonomi'
        'freeparking.'          = 'Freeparking'
        '1stdomains.'           = '1st Domains'
        'discountdomains.'      = 'Discount Domains'
        'domains4less.'         = 'Domains4Less'
        'iwantmyname.com'       = 'iwantmyname'
        'voyager.net.nz'        = 'Voyager'
        'umbrellahosting.'      = 'Umbrellar'
        'sitehost.co.nz'        = 'SiteHost'
        'cyon.ch'               = 'Cyon'
        'name.com'              = 'Name.com'
        'gandi.net'             = 'Gandi'
        'hover.com'             = 'Hover'
        'hostinger.com'         = 'Hostinger'
        'bluehost.com'          = 'Bluehost'
        'hostgator.com'         = 'HostGator'
        'register.com'          = 'Register.com'
        'tucows.com'            = 'Tucows'
        'crsnic.net'            = 'Network Solutions'
        'worldnic.com'          = 'Network Solutions'
        'dynect.net'            = 'Dyn (Oracle)'
        'oracledns.com'         = 'Oracle DNS'
        'googledomains.com'     = 'Google Domains'
        'google.com'            = 'Google'
    }

    foreach ($F_Key in $F_Map.Keys)
    {
        if ($F_Sample -like "*$F_Key*")
        {
            return $F_Map[$F_Key]
        }
    }

    return ('Unknown ({0})' -f $F_Sample)
}

# ---------------------------------------------------------------------------
# Helper: resolve DNS with authoritative-first then public-resolver fallback
# ---------------------------------------------------------------------------
function Resolve-DnsRecordSafe
{
    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory = $true)]
        [string]$Name,

        [Parameter(Mandatory = $true)]
        [ValidateSet('A', 'AAAA', 'CNAME', 'MX', 'NS', 'TXT', 'SOA')]
        [string]$Type,

        [Parameter(Mandatory = $false)]
        [string[]]$PreferredServers
    )

    $F_Servers = @()
    if ($PreferredServers)
    {
        # Resolve NS hostnames to IPs so the query goes straight to the
        # authoritative server. Querying by hostname can route through a
        # caching resolver, which makes TTLs count down between runs.
        foreach ($F_Pref in $PreferredServers)
        {
            if ($F_Pref -match '^[0-9a-fA-F:.]+$')
            {
                $F_Servers += $F_Pref
            }
            else
            {
                try
                {
                    $F_Ips = Resolve-DnsName -Name $F_Pref -Type A -DnsOnly -NoHostsFile -QuickTimeout -ErrorAction Stop |
                        Where-Object { $_.Type -eq 'A' } | Select-Object -ExpandProperty IPAddress
                    if ($F_Ips) { $F_Servers += $F_Ips }
                }
                catch
                {
                    # Fall back to the hostname itself if resolution fails
                    $F_Servers += $F_Pref
                }
            }
        }
    }
    $F_Servers += $S_PublicResolvers
    $F_Servers = $F_Servers | Where-Object { $_ } | Select-Object -Unique

    foreach ($F_Server in $F_Servers)
    {
        try
        {
            $F_Result = Resolve-DnsName -Name $Name -Type $Type -Server $F_Server `
                -DnsOnly -NoHostsFile -QuickTimeout -ErrorAction Stop
            if ($F_Result)
            {
                return [pscustomobject]@{
                    Records = @($F_Result | Where-Object { $_.Type -eq $Type -or ($Type -eq 'CNAME' -and $_.Type -eq 'CNAME') })
                    AllRecords = @($F_Result)
                    Server  = $F_Server
                    Error   = $null
                }
            }
        }
        catch
        {
            # Try the next resolver
            continue
        }
    }

    return [pscustomobject]@{
        Records    = @()
        AllRecords = @()
        Server     = $null
        Error      = ('No response from authoritative or public resolvers for {0} {1}' -f $Type, $Name)
    }
}

# ---------------------------------------------------------------------------
# Microsoft Graph connection
# ---------------------------------------------------------------------------
Write-Host 'Inspecting existing Microsoft Graph context...' -ForegroundColor Cyan

$S_ExistingContext = $null
try
{
    $S_ExistingContext = Get-MgContext
}
catch
{
    $S_ExistingContext = $null
}

if ($S_ExistingContext)
{
    Write-Host ''
    Write-Host 'An existing Microsoft Graph connection was found:' -ForegroundColor Cyan
    Write-Host ('  Account     : {0}' -f $S_ExistingContext.Account)
    Write-Host ('  Tenant ID   : {0}' -f $S_ExistingContext.TenantId)
    Write-Host ('  Environment : {0}' -f $S_ExistingContext.Environment)
    Write-Host ('  Scopes      : {0}' -f (($S_ExistingContext.Scopes | Sort-Object) -join ', '))
    Write-Host ''

    if ($S_RequireGraphContextConfirmation)
    {
        $S_Answer = Read-Host 'Continue using this existing Microsoft Graph connection? [Y/N]'
        if ($S_Answer -notmatch '^(Y|y)')
        {
            Write-Host 'Disconnecting existing Microsoft Graph session...' -ForegroundColor Yellow
            try { Disconnect-MgGraph | Out-Null } catch { Write-Warning $_.Exception.Message }
            $S_ExistingContext = $null
        }
    }
}

if (-not $S_ExistingContext)
{
    Write-Host 'Connecting to Microsoft Graph with the read-only scope set...' -ForegroundColor Cyan
    try
    {
        Connect-MgGraph -Scopes $S_RequiredGraphScopes -NoWelcome
    }
    catch
    {
        throw ('Connect-MgGraph failed: {0}' -f $_.Exception.Message)
    }
}

$S_Context = Get-MgContext
if (-not $S_Context)
{
    throw 'Unable to retrieve Microsoft Graph context after Connect-MgGraph.'
}

Write-Host ''
Write-Host 'Active Microsoft Graph context:' -ForegroundColor Green
Write-Host ('  Account     : {0}' -f $S_Context.Account)
Write-Host ('  Tenant ID   : {0}' -f $S_Context.TenantId)
Write-Host ('  Environment : {0}' -f $S_Context.Environment)
Write-Host ''

# ---------------------------------------------------------------------------
# Tenant + domain discovery
# ---------------------------------------------------------------------------
Write-Host 'Retrieving tenant organisation details...' -ForegroundColor Cyan
$S_Org = $null
try
{
    $S_Org = Get-MgOrganization -ErrorAction Stop | Select-Object -First 1
}
catch
{
    Write-Warning ('Get-MgOrganization failed: {0}' -f $_.Exception.Message)
}
Start-Sleep -Milliseconds $S_GraphRequestDelayMilliseconds

Write-Host 'Retrieving tenant domains...' -ForegroundColor Cyan
$S_AllDomains = @()
try
{
    $S_AllDomains = Get-MgDomain -All -ErrorAction Stop
}
catch
{
    throw ('Get-MgDomain failed: {0}' -f $_.Exception.Message)
}
Start-Sleep -Milliseconds $S_GraphRequestDelayMilliseconds

if (-not $S_AllDomains -or $S_AllDomains.Count -eq 0)
{
    throw 'No domains were returned from Microsoft Graph.'
}

Write-Host ('  Found {0} domain(s) in the tenant.' -f $S_AllDomains.Count) -ForegroundColor Green

# ---------------------------------------------------------------------------
# Per-domain DNS discovery
# ---------------------------------------------------------------------------
$S_DomainResults = New-Object System.Collections.Generic.List[object]

foreach ($S_Domain in ($S_AllDomains | Sort-Object Id))
{
    $S_DomainName = $S_Domain.Id
    $S_IsOnMicrosoft = $S_DomainName -like '*.onmicrosoft.com'

    Write-Host ('Processing domain: {0}' -f $S_DomainName) -ForegroundColor Cyan

    $S_NameServers = @()
    $S_NsError     = $null
    $S_HostingProvider = 'N/A'
    $S_Registrar       = 'N/A'
    $S_RegistrarSource = $null
    $S_RegistrarError  = $null
    $S_SpfRecord   = $null
    $S_SpfTtl      = $null
    $S_SpfError    = $null
    $S_Dkim1Status = 'N/A'
    $S_Dkim2Status = 'N/A'
    $S_Dkim1Length = 'N/A'
    $S_Dkim2Length = 'N/A'
    $S_DmarcRecord = $null
    $S_DmarcPolicy = 'N/A'
    $S_DmarcTtl    = $null
    $S_DmarcError  = $null
    $S_MxRecords   = @()
    $S_MxError     = $null

    # NS, SPF and MX are reported for every domain, including *.onmicrosoft.com
    # (Microsoft 365 mail flow uses the tenant's onmicrosoft.com namespace too).
    # DKIM and DMARC are reported for tenant-owned domains only by default;
    # for *.onmicrosoft.com they are surfaced only when -IncludeOnMicrosoftDomains
    # is supplied (see HTML section filters below).

    # Name servers
    $F_NsLookup = Resolve-DnsRecordSafe -Name $S_DomainName -Type 'NS'
        if ($F_NsLookup.Records.Count -gt 0)
        {
            $S_NameServers = @($F_NsLookup.Records | ForEach-Object { $_.NameHost } | Sort-Object -Unique)
            $S_HostingProvider = Get-DnsHostingProvider -NameServers $S_NameServers
        }
        else
        {
            $S_NsError = $F_NsLookup.Error
        }

        # *.onmicrosoft.com is owned and operated by Microsoft - hard-code the
        # Registrar and DNS Hosting Provider so the inventory is unambiguous.
        if ($S_IsOnMicrosoft)
        {
            $S_Registrar       = 'Microsoft'
            $S_RegistrarSource = 'Microsoft-owned namespace'
            $S_RegistrarError  = $null
            $S_HostingProvider = 'Microsoft'
        }

        # Registrar (RDAP first, WHOIS fallback for known TLDs).
        if ($S_IsOnMicrosoft)
        {
            # Already set above.
        }
        elseif (-not $S_SkipRegistrarLookup)
        {
            try
            {
                $F_RegLookup = Get-DomainRegistrar -DomainName $S_DomainName
                if ($F_RegLookup.Registrar)
                {
                    $S_Registrar       = $F_RegLookup.Registrar
                    $S_RegistrarSource = $F_RegLookup.Source
                }
                else
                {
                    $S_RegistrarError = $F_RegLookup.Error
                    $S_RegistrarSource = $F_RegLookup.Source
                }
            }
            catch
            {
                $S_RegistrarError = ('Registrar lookup threw: {0}' -f $_.Exception.Message)
            }
        }
        else
        {
            $S_RegistrarError = 'Registrar lookup skipped (-SkipRegistrarLookup)'
        }

        # SPF
        $F_TxtLookup = Resolve-DnsRecordSafe -Name $S_DomainName -Type 'TXT' -PreferredServers $S_NameServers
        if ($F_TxtLookup.Records.Count -gt 0)
        {
            $F_SpfMatch = $F_TxtLookup.Records | Where-Object {
                ($_.Strings -join '') -match '^\s*v=spf1'
            } | Select-Object -First 1
            if ($F_SpfMatch)
            {
                $S_SpfRecord = ($F_SpfMatch.Strings -join '')
                $S_SpfTtl    = $F_SpfMatch.TTL
            }
            else
            {
                $S_SpfError = 'No v=spf1 TXT record found'
            }
        }
        else
        {
            $S_SpfError = $F_TxtLookup.Error
        }

        # DKIM and DMARC are only meaningful for tenant-owned domains.
        # For *.onmicrosoft.com we report a dash so the row count stays consistent
        # if -IncludeOnMicrosoftDomains is supplied.
        if (-not $S_IsOnMicrosoft)
        {
            # DKIM (Microsoft 365 selectors)
            # Per-selector status values:
            #   Healthy   - CNAME found, target TXT published with p=, key >= 2048-bit
            #   Weak      - CNAME found, target TXT published with p=, key < 2048-bit
            #   Unhealthy - CNAME found but target TXT does not resolve / has no p=
            #   Missing   - No CNAME at the selector
            foreach ($F_Selector in 1, 2)
            {
                $F_DkimName = ('selector{0}._domainkey.{1}' -f $F_Selector, $S_DomainName)
                $F_DkimLookup = Resolve-DnsRecordSafe -Name $F_DkimName -Type 'CNAME' -PreferredServers $S_NameServers
                $F_Cname = $F_DkimLookup.Records | Select-Object -First 1
                $F_LengthValue = 'N/A'
                $F_KeyBits     = 0
                $F_StatusValue = 'Missing'

                if ($F_Cname -and $F_Cname.NameHost)
                {
                    $F_TargetTxtLookup = Resolve-DnsRecordSafe -Name $F_Cname.NameHost -Type 'TXT'
                    $F_TargetTxt = $F_TargetTxtLookup.Records | Where-Object {
                        ($_.Strings -join '') -match 'p='
                    } | Select-Object -First 1

                    if ($F_TargetTxt)
                    {
                        $F_PMatch = [regex]::Match(($F_TargetTxt.Strings -join ''), 'p=([A-Za-z0-9+/=]+)')
                        if ($F_PMatch.Success -and -not [string]::IsNullOrWhiteSpace($F_PMatch.Groups[1].Value))
                        {
                            $F_KeyBytes = $F_PMatch.Groups[1].Value
                            try
                            {
                                $F_KeyBlob = [Convert]::FromBase64String($F_KeyBytes)
                                if     ($F_KeyBlob.Length -ge 380) { $F_KeyBits = 4096; $F_LengthValue = '4096-bit' }
                                elseif ($F_KeyBlob.Length -ge 280) { $F_KeyBits = 2048; $F_LengthValue = '2048-bit' }
                                elseif ($F_KeyBlob.Length -ge 140) { $F_KeyBits = 1024; $F_LengthValue = '1024-bit' }
                                else                                { $F_KeyBits = 0;    $F_LengthValue = ('{0} bytes' -f $F_KeyBlob.Length) }

                                if ($F_KeyBits -ge 2048)      { $F_StatusValue = 'Healthy' }
                                elseif ($F_KeyBits -gt 0)     { $F_StatusValue = 'Weak' }
                                else                          { $F_StatusValue = 'Unhealthy' }
                            }
                            catch
                            {
                                $F_LengthValue = 'Unparseable'
                                $F_StatusValue = 'Unhealthy'
                            }
                        }
                        else
                        {
                            $F_StatusValue = 'Unhealthy'
                        }
                    }
                    else
                    {
                        $F_StatusValue = 'Unhealthy'
                    }
                }
                else
                {
                    $F_StatusValue = 'Missing'
                }

                if ($F_Selector -eq 1)
                {
                    $S_Dkim1Status = $F_StatusValue
                    $S_Dkim1Length = $F_LengthValue
                }
                else
                {
                    $S_Dkim2Status = $F_StatusValue
                    $S_Dkim2Length = $F_LengthValue
                }
            }

            # DMARC
            $F_DmarcLookup = Resolve-DnsRecordSafe -Name ('_dmarc.{0}' -f $S_DomainName) -Type 'TXT' -PreferredServers $S_NameServers
            if ($F_DmarcLookup.Records.Count -gt 0)
            {
                $F_DmarcMatch = $F_DmarcLookup.Records | Where-Object {
                    ($_.Strings -join '') -match '^\s*v=DMARC1'
                } | Select-Object -First 1
                if ($F_DmarcMatch)
                {
                    $S_DmarcRecord = ($F_DmarcMatch.Strings -join '')
                    $S_DmarcTtl    = $F_DmarcMatch.TTL
                    $F_PolicyMatch = [regex]::Match($S_DmarcRecord, '(?i)\bp\s*=\s*(none|quarantine|reject)')
                    if ($F_PolicyMatch.Success)
                    {
                        $S_DmarcPolicy = $F_PolicyMatch.Groups[1].Value.ToLowerInvariant()
                    }
                }
                else
                {
                    $S_DmarcError = 'No v=DMARC1 TXT record found'
                }
            }
            else
            {
                $S_DmarcError = $F_DmarcLookup.Error
            }
        }

        # MX
        $F_MxLookup = Resolve-DnsRecordSafe -Name $S_DomainName -Type 'MX' -PreferredServers $S_NameServers
        if ($F_MxLookup.Records.Count -gt 0)
        {
            $S_MxRecords = @($F_MxLookup.Records | Sort-Object Preference | ForEach-Object {
                [pscustomobject]@{
                    Exchange   = $_.NameExchange
                    Preference = $_.Preference
                    TTL        = $_.TTL
                }
            })
        }
        else
        {
            $S_MxError = $F_MxLookup.Error
        }

    $S_DomainResults.Add([pscustomobject]@{
        DomainName        = $S_DomainName
        IsDefault         = [bool]$S_Domain.IsDefault
        IsInitial         = [bool]$S_Domain.IsInitial
        IsVerified        = [bool]$S_Domain.IsVerified
        IsOnMicrosoft     = $S_IsOnMicrosoft
        SupportedServices = @($S_Domain.SupportedServices)
        AuthenticationType = $S_Domain.AuthenticationType
        NameServers       = $S_NameServers
        NameServerError   = $S_NsError
        HostingProvider   = $S_HostingProvider
        Registrar         = $S_Registrar
        RegistrarSource   = $S_RegistrarSource
        RegistrarError    = $S_RegistrarError
        SpfRecord         = $S_SpfRecord
        SpfTtl            = $S_SpfTtl
        SpfError          = $S_SpfError
        Dkim1Status       = $S_Dkim1Status
        Dkim2Status       = $S_Dkim2Status
        Dkim1Length       = $S_Dkim1Length
        Dkim2Length       = $S_Dkim2Length
        DmarcRecord       = $S_DmarcRecord
        DmarcPolicy       = $S_DmarcPolicy
        DmarcTtl          = $S_DmarcTtl
        DmarcError        = $S_DmarcError
        MxRecords         = $S_MxRecords
        MxError           = $S_MxError
    })
}

# ---------------------------------------------------------------------------
# Resolve output path
# ---------------------------------------------------------------------------
if (-not $OutputPath)
{
    $OutputPath = (Get-Location).Path
}

$S_OutputFolder = $null
$S_OutputFile   = $null

if (Test-Path -LiteralPath $OutputPath -PathType Container)
{
    $S_OutputFolder = (Resolve-Path -LiteralPath $OutputPath).Path
}
elseif ([System.IO.Path]::GetExtension($OutputPath) -ieq '.html')
{
    $S_OutputFile   = $OutputPath
    $S_OutputFolder = Split-Path -Parent $OutputPath
    if (-not $S_OutputFolder) { $S_OutputFolder = $S_ScriptRoot }
}
else
{
    $S_OutputFolder = $OutputPath
}

if (-not (Test-Path -LiteralPath $S_OutputFolder))
{
    New-Item -ItemType Directory -Path $S_OutputFolder -Force | Out-Null
}

if (-not $S_OutputFile)
{
    $S_OutputFile = Join-Path -Path $S_OutputFolder -ChildPath ("ReportDomains_{0}.html" -f $S_Timestamp)
}

# ---------------------------------------------------------------------------
# Build HTML report
# ---------------------------------------------------------------------------
function ConvertTo-HtmlSafe
{
    param([Parameter(Mandatory = $true)][AllowNull()][object]$Value)
    if ($null -eq $Value -or ($Value -is [string] -and [string]::IsNullOrWhiteSpace($Value)))
    {
        return 'N/A'
    }
    return [System.Net.WebUtility]::HtmlEncode([string]$Value)
}

function Format-TtlFriendly
{
    param([AllowNull()][object]$Seconds)
    if ($null -eq $Seconds) { return 'N/A' }
    $F_Total = 0
    if (-not [int]::TryParse([string]$Seconds, [ref]$F_Total)) { return 'N/A' }
    if ($F_Total -le 0) { return '0s' }
    if ($F_Total -lt 3600)
    {
        $F_Mins = [math]::Round($F_Total / 60, 1)
        return ('{0} min ({1}s)' -f $F_Mins, $F_Total)
    }
    if ($F_Total -lt 86400)
    {
        $F_Hours = [math]::Round($F_Total / 3600, 1)
        return ('{0} h ({1}s)' -f $F_Hours, $F_Total)
    }
    $F_Days = [math]::Round($F_Total / 86400, 1)
    return ('{0} d ({1}s)' -f $F_Days, $F_Total)
}

$S_TenantDisplay = if ($S_Org -and $S_Org.DisplayName) { $S_Org.DisplayName } else { 'Unknown Tenant' }
$S_TenantId      = if ($S_Context.TenantId)            { $S_Context.TenantId }   else { 'Unknown' }
$S_PrimaryDomain = ($S_AllDomains | Where-Object { $_.IsDefault } | Select-Object -First 1).Id
$S_InitialDomain = ($S_AllDomains | Where-Object { $_.IsInitial } | Select-Object -First 1).Id
$S_ReportDate    = Get-Date -Format 'dd MMM yyyy HH:mm'

# --- Section 1: Domain Inventory ---
$S_InventoryRows = ($S_DomainResults | ForEach-Object {
    $F_Workloads = if ($_.SupportedServices -and $_.SupportedServices.Count -gt 0) { ($_.SupportedServices | Sort-Object) -join ', ' } else { 'None' }
    $F_NsList    = if ($_.NameServers -and $_.NameServers.Count -gt 0) { ($_.NameServers -join '<br/>') } else { 'N/A' }
    $F_Notes = @()
    if (-not $_.IsVerified)    { $F_Notes += 'Unverified' }
    if ($_.IsDefault)          { $F_Notes += 'Primary / Default' }
    if ($_.IsInitial)          { $F_Notes += 'Initial onmicrosoft.com' }
    if ($_.NameServerError)    { $F_Notes += 'NS lookup failed' }
    if ($_.RegistrarError -and $_.Registrar -eq 'N/A')
    {
        $F_Notes += ('Registrar: {0}' -f $_.RegistrarError)
    }
    elseif ($_.RegistrarSource)
    {
        $F_Notes += ('Registrar source: {0}' -f $_.RegistrarSource)
    }
    $F_NotesText = if ($F_Notes.Count -gt 0) { $F_Notes -join '; ' } else { '' }

    $F_RowClass = ''
    if (-not $_.IsVerified -or $_.NameServerError) { $F_RowClass = ' class="warn"' }

    "<tr$F_RowClass><td>{0}</td><td>{1}</td><td>{2}</td><td>{3}</td><td>{4}</td><td>{5}</td></tr>" -f `
        (ConvertTo-HtmlSafe $_.DomainName), `
        (ConvertTo-HtmlSafe $F_Workloads), `
        (ConvertTo-HtmlSafe $_.Registrar), `
        (ConvertTo-HtmlSafe $_.HostingProvider), `
        $F_NsList, `
        (ConvertTo-HtmlSafe $F_NotesText)
}) -join "`n"

# --- Section 2: Primary Tenant Domain (Field/Value) ---
$S_PrimaryDomainDisplay = if ($S_PrimaryDomain) { $S_PrimaryDomain } else { 'N/A' }
$S_InitialDomainDisplay = if ($S_InitialDomain) { $S_InitialDomain } else { 'N/A' }
$S_VerifiedDomainCount  = @($S_AllDomains | Where-Object { $_.IsVerified }).Count

$S_PrimaryRows = @(
    [pscustomobject]@{ Field = 'Tenant Display Name';   Value = $S_TenantDisplay }
    [pscustomobject]@{ Field = 'Tenant ID';             Value = $S_TenantId }
    [pscustomobject]@{ Field = 'Primary Domain';        Value = $S_PrimaryDomainDisplay }
    [pscustomobject]@{ Field = 'Initial Domain';        Value = $S_InitialDomainDisplay }
    [pscustomobject]@{ Field = 'Verified Domain Count'; Value = $S_VerifiedDomainCount }
    [pscustomobject]@{ Field = 'Total Domain Count';    Value = $S_AllDomains.Count }
    [pscustomobject]@{ Field = 'Generated At (Local)';  Value = $S_ReportDate }
) | ForEach-Object {
    "<tr><th>{0}</th><td>{1}</td></tr>" -f (ConvertTo-HtmlSafe $_.Field), (ConvertTo-HtmlSafe $_.Value)
}
$S_PrimaryRowsHtml = ($S_PrimaryRows -join "`n")

# --- Section 3a: SPF ---
$S_SpfRows = ($S_DomainResults | ForEach-Object {
    $F_RowClass = ''
    if ($_.SpfError)
    {
        $F_RowClass = ' class="bad"'
    }
    elseif ($_.SpfRecord -and $_.SpfRecord -match '\+all')
    {
        $F_RowClass = ' class="bad"'
    }
    elseif ($_.SpfRecord -and ($_.SpfRecord -notmatch 'include:spf\.protection\.outlook\.com' -or $_.SpfRecord -match '\?all'))
    {
        $F_RowClass = ' class="warn"'
    }

    "<tr$F_RowClass><td>{0}</td><td><code>{1}</code></td><td>{2}</td></tr>" -f `
        (ConvertTo-HtmlSafe $_.DomainName), `
        (ConvertTo-HtmlSafe $_.SpfRecord), `
        (ConvertTo-HtmlSafe (Format-TtlFriendly $_.SpfTtl))
}) -join "`n"

# --- Section 3b: DKIM ---
function Get-DkimStatusBadge
{
    param([Parameter(Mandatory = $true)][string]$Status)
    switch ($Status)
    {
        'Healthy'   { return '<span class="badge badge-good">Healthy</span>' }
        'Weak'      { return '<span class="badge badge-warn">Weak</span>' }
        'Unhealthy' { return '<span class="badge badge-bad">Unhealthy</span>' }
        'Missing'   { return '<span class="badge badge-bad">Missing</span>' }
        default     { return [System.Net.WebUtility]::HtmlEncode($Status) }
    }
}

$S_DkimRows = ($S_DomainResults | ForEach-Object {
    if ($_.IsOnMicrosoft)
    {
        "<tr><td>{0}</td><td>N/A</td><td>N/A</td><td>N/A</td><td>N/A</td></tr>" -f `
            (ConvertTo-HtmlSafe $_.DomainName)
        return
    }

    $F_BadStatuses  = @('Missing', 'Unhealthy')
    $F_WarnStatuses = @('Weak')

    $F_RowClass = ''
    if (($F_BadStatuses -contains $_.Dkim1Status) -or ($F_BadStatuses -contains $_.Dkim2Status))
    {
        $F_RowClass = ' class="bad"'
    }
    elseif (($F_WarnStatuses -contains $_.Dkim1Status) -or ($F_WarnStatuses -contains $_.Dkim2Status))
    {
        $F_RowClass = ' class="warn"'
    }

    "<tr$F_RowClass><td>{0}</td><td>{1}</td><td>{2}</td><td>{3}</td><td>{4}</td></tr>" -f `
        (ConvertTo-HtmlSafe $_.DomainName), `
        (Get-DkimStatusBadge -Status $_.Dkim1Status), `
        (Get-DkimStatusBadge -Status $_.Dkim2Status), `
        (ConvertTo-HtmlSafe $_.Dkim1Length), `
        (ConvertTo-HtmlSafe $_.Dkim2Length)
}) -join "`n"

# --- Section 3c: DMARC ---
function Get-DmarcPolicyBadge
{
    param([string]$Policy)
    $F_Lower = ([string]$Policy).Trim().ToLowerInvariant()
    switch ($F_Lower)
    {
        'reject'     { return '<span class="badge badge-good">reject</span>' }
        'quarantine' { return '<span class="badge badge-good">quarantine</span>' }
        ''           { return '<span class="badge badge-bad">none/missing</span>' }
        default      { return ('<span class="badge badge-bad">{0}</span>' -f [System.Net.WebUtility]::HtmlEncode($F_Lower)) }
    }
}

$S_DmarcRows = ($S_DomainResults | ForEach-Object {
    if ($_.IsOnMicrosoft)
    {
        "<tr><td>{0}</td><td>N/A</td><td>N/A</td><td>N/A</td></tr>" -f `
            (ConvertTo-HtmlSafe $_.DomainName)
        return
    }

    $F_RowClass = ''
    if ($_.DmarcError -or [string]::IsNullOrWhiteSpace($_.DmarcPolicy) -or ($_.DmarcPolicy -eq 'none'))
    {
        $F_RowClass = ' class="bad"'
    }

    "<tr$F_RowClass><td>{0}</td><td><code>{1}</code></td><td>{2}</td><td>{3}</td></tr>" -f `
        (ConvertTo-HtmlSafe $_.DomainName), `
        (ConvertTo-HtmlSafe $_.DmarcRecord), `
        (Get-DmarcPolicyBadge -Policy $_.DmarcPolicy), `
        (ConvertTo-HtmlSafe (Format-TtlFriendly $_.DmarcTtl))
}) -join "`n"

# --- Section 4: MX ---
$S_MxRows = ($S_DomainResults | ForEach-Object {
    $F_DomainResult = $_
    if ($F_DomainResult.MxError)
    {
        "<tr class=""bad""><td>{0}</td><td>{1}</td><td>N/A</td><td>N/A</td></tr>" -f `
            (ConvertTo-HtmlSafe $F_DomainResult.DomainName), `
            (ConvertTo-HtmlSafe $F_DomainResult.MxError)
    }
    elseif (-not $F_DomainResult.MxRecords -or $F_DomainResult.MxRecords.Count -eq 0)
    {
        "<tr class=""warn""><td>{0}</td><td>No MX records returned</td><td>N/A</td><td>N/A</td></tr>" -f `
            (ConvertTo-HtmlSafe $F_DomainResult.DomainName)
    }
    else
    {
        $F_DomainResult.MxRecords | ForEach-Object {
            $F_Mx = $_
            $F_RowClass = ''
            if ($F_Mx.Exchange -notmatch 'mail\.protection\.outlook\.com$')
            {
                $F_RowClass = ' class="warn"'
            }
            "<tr$F_RowClass><td>{0}</td><td>{1}</td><td>{2}</td><td>{3}</td></tr>" -f `
                (ConvertTo-HtmlSafe $F_DomainResult.DomainName), `
                (ConvertTo-HtmlSafe $F_Mx.Exchange), `
                (ConvertTo-HtmlSafe $F_Mx.Preference), `
                (ConvertTo-HtmlSafe (Format-TtlFriendly $F_Mx.TTL))
        }
    }
}) -join "`n"

$S_Html = @"
<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Microsoft 365 Domain and DNS Report - $([System.Net.WebUtility]::HtmlEncode($S_TenantDisplay))</title>
<style>
  * { margin: 0; padding: 0; box-sizing: border-box; }
  body { font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; background: #f0f2f5; color: #2c3e50; padding: 30px; }
  .header { background: linear-gradient(135deg, #1a1a2e, #16213e); color: #fff; padding: 28px 36px; border-radius: 12px; margin-bottom: 28px; }
  .header h1 { font-size: 1.6em; margin-bottom: 6px; }
  .header p  { font-size: 0.9em; opacity: 0.85; }
  .section { background: #fff; border-radius: 10px; padding: 28px 30px; box-shadow: 0 2px 8px rgba(0,0,0,0.06); margin-bottom: 24px; overflow-x: auto; }
  .section h2 { font-size: 1.15em; margin-bottom: 14px; color: #1a1a2e; border-bottom: 2px solid #e6e9ef; padding-bottom: 8px; }
  .section h3 { font-size: 1.0em; margin: 18px 0 10px 0; color: #1a1a2e; }
  table { width: 100%; border-collapse: collapse; font-size: 0.88em; }
  th { background: #1a1a2e; color: #fff; padding: 10px 12px; text-align: left; font-weight: 600; }
  td, th.kv { padding: 9px 12px; border-bottom: 1px solid #eef0f3; vertical-align: top; }
  tr:hover td { background: #f8f9fa; }
  tr.warn td { background: #fff8e1; }
  tr.bad  td { background: #fdecea; }
  tr.warn:hover td { background: #fff3c4; }
  tr.bad:hover td  { background: #fbd9d4; }
  code { font-family: Consolas, 'Courier New', monospace; font-size: 0.92em; word-break: break-all; }
  .legend { font-size: 0.82em; color: #666; margin-top: 10px; }
  .legend span.swatch { display: inline-block; width: 14px; height: 14px; border-radius: 3px; vertical-align: middle; margin: 0 6px 0 14px; border: 1px solid #ccc; }
  .badge { display: inline-block; padding: 2px 10px; border-radius: 12px; font-size: 0.82em; font-weight: 600; line-height: 1.4; }
  .badge-good { background: #e7f7ec; color: #1e7e34; border: 1px solid #b6e2c0; }
  .badge-warn { background: #fff3cd; color: #8a6d1c; border: 1px solid #ffe48c; }
  .badge-bad  { background: #f8d7da; color: #a4202b; border: 1px solid #f1b0b7; }
  .footer { text-align: center; font-size: 0.8em; color: #999; margin-top: 18px; }
  table.kv th { width: 30%; background: #f4f6fa; color: #1a1a2e; text-align: left; }
</style>
</head>
<body>

<div class="header">
  <h1>Microsoft 365 Domain and DNS Report</h1>
  <p>Tenant: $([System.Net.WebUtility]::HtmlEncode($S_TenantDisplay)) &nbsp;|&nbsp; Tenant ID: $([System.Net.WebUtility]::HtmlEncode($S_TenantId)) &nbsp;|&nbsp; Generated: $([System.Net.WebUtility]::HtmlEncode($S_ReportDate))</p>
</div>

<div class="section">
  <h2>1. Domain Inventory</h2>
  <table>
    <thead>
      <tr>
        <th>Domain Name</th>
        <th>Enabled Workloads</th>
        <th>Registrar</th>
        <th>DNS Hosting Provider</th>
        <th>Name Servers</th>
        <th>Notes</th>
      </tr>
    </thead>
    <tbody>
$S_InventoryRows
    </tbody>
  </table>
  <p class="legend">Registrar discovered via RDAP (IANA bootstrap) with WHOIS port-43 fallback for supported TLDs (e.g. .nz). Update manually where the lookup returned N/A.
    <span class="swatch" style="background:#fff8e1"></span>Review
    <span class="swatch" style="background:#fdecea"></span>Action required
  </p>
</div>

<div class="section">
  <h2>2. Primary Tenant Domain</h2>
  <table class="kv">
    <thead><tr><th class="kv">Field</th><th>Value</th></tr></thead>
    <tbody>
$S_PrimaryRowsHtml
    </tbody>
  </table>
</div>

<div class="section">
  <h2>3. Email Authentication Records</h2>

  <h3>SPF</h3>
  <table>
    <thead><tr><th>Domain</th><th>SPF Record</th><th>TTL</th></tr></thead>
    <tbody>
$S_SpfRows
    </tbody>
  </table>

  <h3>DKIM</h3>
  <table>
    <thead>
      <tr>
        <th>Domain</th>
        <th>DKIM Key 1 Status</th>
        <th>DKIM Key 2 Status</th>
        <th>Key 1 Length</th>
        <th>Key 2 Length</th>
      </tr>
    </thead>
    <tbody>
$S_DkimRows
    </tbody>
  </table>

  <h3>DMARC</h3>
  <table>
    <thead>
      <tr>
        <th>Domain</th>
        <th>DMARC Record</th>
        <th>Policy State</th>
        <th>TTL</th>
      </tr>
    </thead>
    <tbody>
$S_DmarcRows
    </tbody>
  </table>
</div>

<div class="section">
  <h2>4. MX Records</h2>
  <table>
    <thead>
      <tr>
        <th>Domain</th>
        <th>MX Record</th>
        <th>Priority</th>
        <th>TTL</th>
      </tr>
    </thead>
    <tbody>
$S_MxRows
    </tbody>
  </table>
</div>

<div class="footer">Generated by ReportDomains.ps1 - read-only Microsoft 365 tenant discovery report.</div>

</body>
</html>
"@

$S_Html | Out-File -FilePath $S_OutputFile -Encoding UTF8 -Force

Write-Host ''
Write-Host 'Report generation complete.' -ForegroundColor Green
Write-Host ('  Output file : {0}' -f $S_OutputFile)
Write-Host ('  Domains processed : {0}' -f $S_DomainResults.Count)
Write-Host ''
