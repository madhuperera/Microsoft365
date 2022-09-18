$DNSServerToUse = "1.1.1.1"
$DkimConfiguration = Get-DkimSigningConfig | Where-Object {$_.Domain -notlike "*.onmicrosoft.com"}

foreach ($DkimEntry in $DkimConfiguration)
{
    [String] $DomainName = $DkimEntry.Domain
    [bool] $DkimEnabled = $DkimEntry.Enabled
    Write-Output "`nReport on $DomainName"
    if ($DkimEnabled)
    {
        Write-Output "Dkim is enabled for the domain: $DomainName"
        try
        {
            # Selector 1
            $Selector1DNSName = "selector1._domainkey.$DomainName"
            $Result1 = Resolve-DnsName -Type CNAME -Server $DNSServerToUse -Name $Selector1DNSName -ErrorAction SilentlyContinue
            if ($Result1)
            {
                $Selector1DnsHostName = $Result1.NameHost
                if (Resolve-DnsName -Type CNAME -Server $DNSServerToUse -Name $Selector1DnsHostName -ErrorAction SilentlyContinue)
                {
                    Write-Output "Successfully verified Selector 1 for $DomainName"
                }
                else
                {
                    Write-Output "Error validating Selector 1 for $DomainName. Please rotate the Dkim Keys and try again!"
                }
            }
            else
            {
                Write-Output "Error validating DNS for $Selector1DNSName. Please either update $DNSServerToUse or check your DNS Provider"
            }

            # Selector 2
            $Selector2DNSName = "selector2._domainkey.$DomainName"
            $Result2 = Resolve-DnsName -Type CNAME -Server $DNSServerToUse -Name $Selector2DNSName -ErrorAction SilentlyContinue
            if ($Result2)
            {
                $Selector2DnsHostName = $Result2.NameHost
                if (Resolve-DnsName -Type CNAME -Server $DNSServerToUse -Name $Selector2DnsHostName -ErrorAction SilentlyContinue)
                {
                    Write-Output "Successfully verified Selector 2 for $DomainName"
                }
                else
                {
                    Write-Output "Error validating Selector 2 for $DomainName. Please rotate the Dkim Keys and try again!"
                }
            }
            else
            {
                Write-Output "Error validating DNS for $Selector1DNSName. Please either update $DNSServerToUse or check your DNS Provider"
            }

            

        }
        catch
        {
            Write-Output "Something went wrong"
        }
    }
    else
    {
        Write-Output "Dkim is NOT enabled for the domain: $DomainName"
    }
}