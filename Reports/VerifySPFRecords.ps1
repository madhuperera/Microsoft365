$DNSServerToUse = "1.1.1.1"
$AllDomains = (Get-AcceptedDomain | Where-Object {$_.DomainName -notlike "*.onmicrosoft.com"}).DomainName

foreach ($Domain in $AllDomains)
{
    Write-Output "`nReport on $Domain"

    Resolve-DnsName -Type TXT -Server $DNSServerToUse -Name $Domain |`
        Where-Object {$_.Strings -like "*spf*"}
}