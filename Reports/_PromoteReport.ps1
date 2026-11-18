<#
.SYNOPSIS
    Promotes the development copy of a report script to production and
    retires the outgoing production script as a frozen versioned snapshot.

.DESCRIPTION
    Reports that follow the two-name lifecycle keep exactly two live files:

      <ReportName>_prod.ps1   The released script. Other scripts and projects
                              reference this filename only.
      <ReportName>_dev.ps1    The working copy where new features are built.

    This script performs a promotion in three steps:

      1. Validates that both files exist, that the development script parses
         without syntax errors, and that it differs from production.
      2. Copies the current production script to the next frozen snapshot,
         <ReportName>_vN.ps1, where N is one higher than the highest existing
         snapshot number in the folder.
      3. Writes the development script over the production script, replacing
         references to its own _dev filename with the _prod filename.

    The development script is left in place as the starting point for the
    next round of changes. No tenant connection is made and no git commands
    are run; review, commit and tag the result yourself.

.PARAMETER ReportName
    Base name of the report without the _prod / _dev suffix or extension,
    for example ReportMemberMFA.

.PARAMETER Path
    Folder that holds the report scripts. Defaults to the folder this script
    is in.

.EXAMPLE
    .\_PromoteReport.ps1 -ReportName ReportMemberMFA -WhatIf

    Shows what a promotion would do without changing any files.

.EXAMPLE
    .\_PromoteReport.ps1 -ReportName ReportMemberMFA

    Retires ReportMemberMFA_prod.ps1 to the next ReportMemberMFA_vN.ps1 and
    promotes ReportMemberMFA_dev.ps1 to ReportMemberMFA_prod.ps1.
#>

[CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'High')]
param(
    [Parameter(Mandatory = $true)]
    [ValidatePattern('^(Report|Review)[A-Za-z0-9]+$')]
    [string]$ReportName,

    [Parameter(Mandatory = $false)]
    [ValidateNotNullOrEmpty()]
    [string]$Path
)

$ErrorActionPreference = 'Stop'

# ── Resolve paths ──────────────────────────────────────────────────────────────
if (-not $Path)
{
    $Path = if ($PSScriptRoot) { $PSScriptRoot } else { (Get-Location).Path }
}

$S_ProdName = "${ReportName}_prod.ps1"
$S_DevName  = "${ReportName}_dev.ps1"
$S_ProdPath = Join-Path -Path $Path -ChildPath $S_ProdName
$S_DevPath  = Join-Path -Path $Path -ChildPath $S_DevName

foreach ($S_Required in @($S_ProdPath, $S_DevPath))
{
    if (-not (Test-Path -Path $S_Required -PathType Leaf))
    {
        throw "Required script not found: $S_Required"
    }
}

# ── Validate the development script ────────────────────────────────────────────
$S_ParseTokens = $null
$S_ParseErrors = $null
[void][System.Management.Automation.Language.Parser]::ParseFile($S_DevPath, [ref]$S_ParseTokens, [ref]$S_ParseErrors)
if ($S_ParseErrors -and $S_ParseErrors.Count -gt 0)
{
    $S_ParseErrors | ForEach-Object { Write-Warning ("{0} (line {1})" -f $_.Message, $_.Extent.StartLineNumber) }
    throw "$S_DevName has $($S_ParseErrors.Count) syntax error(s). Promotion cancelled."
}

# ── Build the new production content ───────────────────────────────────────────
# Bytes are handled directly so the file's UTF-8 encoding (with or without a
# byte order mark) and line endings are preserved on every PowerShell version.
$S_DevBytes = [System.IO.File]::ReadAllBytes($S_DevPath)
$S_HasBom = (
    $S_DevBytes.Length -ge 3 -and
    $S_DevBytes[0] -eq 0xEF -and $S_DevBytes[1] -eq 0xBB -and $S_DevBytes[2] -eq 0xBF
)
$S_Encoding   = New-Object System.Text.UTF8Encoding($S_HasBom)
$S_BomLength  = if ($S_HasBom) { 3 } else { 0 }
$S_DevText    = $S_Encoding.GetString($S_DevBytes, $S_BomLength, $S_DevBytes.Length - $S_BomLength)
$S_NewProdText  = $S_DevText.Replace($S_DevName, $S_ProdName)
$S_NewProdBytes = [byte[]]($S_Encoding.GetPreamble() + $S_Encoding.GetBytes($S_NewProdText))

$S_ProdBytes = [System.IO.File]::ReadAllBytes($S_ProdPath)
if ([System.Convert]::ToBase64String($S_ProdBytes) -eq [System.Convert]::ToBase64String($S_NewProdBytes))
{
    throw "$S_DevName has no changes compared with $S_ProdName. Nothing to promote."
}

# ── Work out the next frozen snapshot number ───────────────────────────────────
$S_SnapshotPattern = '^' + [regex]::Escape($ReportName) + '_v(\d+)\.ps1$'
$S_HighestSnapshot = 0
foreach ($S_File in Get-ChildItem -Path $Path -Filter "${ReportName}_v*.ps1" -File)
{
    if ($S_File.Name -match $S_SnapshotPattern)
    {
        $S_HighestSnapshot = [math]::Max($S_HighestSnapshot, [int]$Matches[1])
    }
}
$S_RetiredName = "${ReportName}_v$($S_HighestSnapshot + 1).ps1"
$S_RetiredPath = Join-Path -Path $Path -ChildPath $S_RetiredName

# ── Promote ────────────────────────────────────────────────────────────────────
Write-Host "Promotion plan for $ReportName" -ForegroundColor Cyan
Write-Host "  Retire : $S_ProdName -> $S_RetiredName (frozen snapshot)"
Write-Host "  Promote: $S_DevName -> $S_ProdName"
Write-Host "  Keep   : $S_DevName (unchanged, baseline for the next round)"

if ($PSCmdlet.ShouldProcess($S_ProdName, "Retire to $S_RetiredName and replace with $S_DevName"))
{
    Copy-Item -Path $S_ProdPath -Destination $S_RetiredPath
    [System.IO.File]::WriteAllBytes($S_ProdPath, $S_NewProdBytes)

    Write-Host "Promotion complete." -ForegroundColor Green
    Write-Host "Next steps:"
    Write-Host "  1. Update the report's changelog and README entries for the new release."
    Write-Host "  2. Review the changes with git diff, then commit."
    Write-Host "  3. Tag the release if the report uses release tags."
}
