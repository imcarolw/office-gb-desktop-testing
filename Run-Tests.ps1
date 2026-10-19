#Requires -Version 7.0
[CmdletBinding()]
param(
    [string[]]$Application = @('All'),
    [int[]]$CaseId = @(),
    [switch]$Interactive,
    [switch]$ListTests,
    [switch]$Preview
)

$ErrorActionPreference = 'Stop'
Import-Module (Join-Path $PSScriptRoot 'scripts\TestRouting.psm1') -Force
Import-Module (Join-Path $PSScriptRoot 'scripts\EvidenceValidation.psm1') -Force
$catalog = @(Read-OfficeTestCatalog -SolutionRoot $PSScriptRoot)
$selected = @(Select-OfficeTestCases -Catalog $catalog -Application $Application -CaseId $CaseId)
if ($ListTests -and $Preview) { throw 'Use either ListTests or Preview, not both.' }
if ($ListTests) {
    ConvertTo-Json -InputObject $selected -Depth 8
    return
}
if ($CaseId.Count -gt 0 -and -not ($CaseId.Count -eq 1 -and $CaseId[0] -eq 0) -and -not $Interactive) {
    throw 'A case selector requires -Interactive (also when previewing).'
}
$plan = New-OfficeTestPlan -Cases $selected -Interactive $Interactive.IsPresent
if ($Preview) {
    $plan | ConvertTo-Json -Depth 10
    return
}

$previousGate = $env:RUN_OFFICE_UI_TESTS
$previousRunId = $env:OFFICE_UI_RUN_ID
$previousRoot = $env:OFFICE_TEST_ROOT
try {
    $env:OFFICE_TEST_ROOT = $PSScriptRoot
    $env:RUN_OFFICE_UI_TESTS = if ($Interactive) { '1' } else { $null }
    $runId = '{0}-{1}' -f (Get-Date -Format 'yyyyMMdd-HHmmss-fff'), ([guid]::NewGuid().ToString('N'))
    if ($Interactive) {
        $settings = Get-Content -LiteralPath (Join-Path $PSScriptRoot 'settings\office-test-settings.json') -Raw | ConvertFrom-Json
        if ([string]::IsNullOrWhiteSpace($settings.artifactRoot)) { throw 'artifactRoot must not be empty.' }
        $artifactRoot = [IO.Path]::GetFullPath($settings.artifactRoot, $PSScriptRoot)
        $office = @(Get-Process WINWORD,EXCEL,POWERPNT -ErrorAction SilentlyContinue)
        if ($office.Count -gt 0) {
            $details = ($office | ForEach-Object { "$($_.ProcessName) PID $($_.Id)" }) -join ', '
            $office | ForEach-Object { $_.Dispose() }
            throw "Close existing Office sessions before interactive execution: $details. No sessions were closed."
        }
    }
    $failed = $false
    $env:OFFICE_UI_RUN_ID = $runId
    $runDirectory = if ($Interactive) { Join-Path $artifactRoot $runId } else { $null }
    $manifest = [ordered]@{
        RunId = $runId; StartedUtc = [datetime]::UtcNow.ToString('O')
        CompletedUtc = $null; State = 'Running'; Cases = @($plan.Cases)
        Executions = @(); CleanupErrors = @()
    }
    if ($Interactive) {
        New-Item -ItemType Directory -Path $runDirectory -Force | Out-Null
        $manifest | ConvertTo-Json -Depth 12 | Set-Content (Join-Path $runDirectory 'run-manifest.json') -Encoding utf8
    }
    foreach ($project in $plan.Projects) {
        $results = if ($Interactive) { Join-Path $runDirectory "test-results\$($project.Application)" }
            else { Join-Path $PSScriptRoot "TestResults\$runId\$($project.Application)" }
        $exitCode = -1
        $executionError = $null
        try {
            & dotnet test (Join-Path $PSScriptRoot $project.Project) `
                --filter $project.Filter --logger 'console;verbosity=normal' `
                --logger 'trx;LogFileName=results.trx' --results-directory $results
            $exitCode = $LASTEXITCODE
        } catch {
            $executionError = $_.Exception.ToString()
            Write-Warning "Test project could not execute: $executionError"
        }
        $manifest.Executions += [pscustomobject]@{
            Project = $project.Project; ExitCode = $exitCode; Error = $executionError
        }
        if ($exitCode -ne 0 -or $executionError) { $failed = $true }
        if ($Interactive) {
            $manifest | ConvertTo-Json -Depth 12 | Set-Content (Join-Path $runDirectory 'run-manifest.json') -Encoding utf8
        }
    }
    if ($Interactive) {
        $remaining = @(Get-Process WINWORD,EXCEL,POWERPNT -ErrorAction SilentlyContinue)
        foreach ($process in $remaining) {
            $manifest.CleanupErrors += "Office process remains: $($process.ProcessName) PID $($process.Id)."
            $process.Dispose()
        }
        $manifest.State = 'Completed'
        $manifest.CompletedUtc = [datetime]::UtcNow.ToString('O')
        $manifest | ConvertTo-Json -Depth 12 | Set-Content (Join-Path $runDirectory 'run-manifest.json') -Encoding utf8
        $validation = Test-OfficeRunEvidence -RunDirectory $runDirectory
        $validation | ConvertTo-Json -Depth 12 | Set-Content (Join-Path $runDirectory 'evidence-validation.json') -Encoding utf8
        ConvertTo-Json -InputObject @($validation.Cases) -Depth 10 |
            Set-Content (Join-Path $runDirectory 'run-summary.json') -Encoding utf8
        $summaryPath = Write-OfficeRunSummary -Validation $validation
        $validation.Cases | Format-Table Application,CaseId,Status,Steps,Screenshots,EvidenceValid
        foreach ($validationError in $validation.Errors) { Write-Warning $validationError }
        if (-not $validation.Passed) { $failed = $true }
        Write-Host "Run evidence: $runDirectory"
        Write-Host "Results summary: $summaryPath"
        Write-Host 'Automated execution and evidence checks completed. Glyph rendering still requires visual review.'
    }
    if ($failed) { throw 'Test execution or evidence validation failed. Inspect the retained reports and evidence.' }
} finally {
    $env:RUN_OFFICE_UI_TESTS = $previousGate
    $env:OFFICE_UI_RUN_ID = $previousRunId
    $env:OFFICE_TEST_ROOT = $previousRoot
}
