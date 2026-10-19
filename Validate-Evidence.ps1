[CmdletBinding()]
param([Parameter(Mandatory)][string]$RunDirectory)

$ErrorActionPreference = 'Stop'
Import-Module (Join-Path $PSScriptRoot 'scripts\EvidenceValidation.psm1') -Force
$result = Test-OfficeRunEvidence -RunDirectory $RunDirectory
$result | ConvertTo-Json -Depth 12
if (-not $result.Passed) { throw 'Evidence validation failed; inspect Errors in the JSON output.' }
