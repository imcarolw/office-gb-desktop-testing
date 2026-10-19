#Requires -Version 7.0
[CmdletBinding()]
param()

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest
$root = Split-Path $PSScriptRoot
Import-Module (Join-Path $PSScriptRoot 'TestRouting.psm1') -Force
Import-Module (Join-Path $PSScriptRoot 'EvidenceValidation.psm1') -Force
$catalog = @(Read-OfficeTestCatalog -SolutionRoot $root)
$sandbox = Join-Path ([IO.Path]::GetTempPath()) ('OfficeGbWorkflow-' + [guid]::NewGuid().ToString('N'))

function Assert-True([bool]$Condition, [string]$Message) {
    if (-not $Condition) { throw "Assertion failed: $Message" }
}

function Assert-Throws([scriptblock]$Action) {
    $threw = $false
    try { & $Action | Out-Null } catch { $threw = $true }
    Assert-True $threw 'Invalid selection must fail'
}

function Write-Json($Value, [string]$Path) {
    ConvertTo-Json -InputObject $Value -Depth 12 | Set-Content -LiteralPath $Path -Encoding utf8
}

try {
    New-Item -ItemType Directory -Path $sandbox | Out-Null
    $excel = @(Select-OfficeTestCases -Catalog $catalog -Application excel)
    Assert-True ($excel.Count -eq 2 -and @($excel | Where-Object application -ne Excel).Count -eq 0) 'Excel selection'
    $case = $catalog | Where-Object { 'save as' -in $_.aliases -and $_.application -eq 'Word' }
    $selected = @(Select-OfficeTestCases -Catalog $catalog -Application Word -CaseId $case.caseId)
    $plan = New-OfficeTestPlan -Cases $selected -Interactive $true
    Assert-True ($selected.Count -eq 1 -and $plan.Projects[0].Filter -eq "TestCategory=Interactive&(TestCategory=Case$($case.caseId))") 'Save As scenario selection'
    Write-Host 'PASS: Application and catalog scenario select the correct cases'

    Assert-Throws { Select-OfficeTestCases -Catalog $catalog -Application Outlook }
    Assert-Throws { Select-OfficeTestCases -Catalog $catalog -CaseId 9999 }
    Assert-Throws { Select-OfficeTestCases -Catalog $catalog -Application All,Excel }
    Assert-Throws { Select-OfficeTestCases -Catalog $catalog -Application Excel -CaseId $case.caseId }
    Write-Host 'PASS: Unknown and conflicting selections fail'

    $runDirectory = Join-Path $sandbox 'evidence-fixture'
    $caseDirectory = Join-Path $runDirectory "$($case.application)\$($case.caseId)"
    $trxDirectory = Join-Path $runDirectory "test-results\$($case.application)"
    New-Item -ItemType Directory -Path $caseDirectory, $trxDirectory | Out-Null
    Add-Type -AssemblyName System.Drawing
    $start = [datetime]::UtcNow
    $steps = @(
        for ($number = 1; $number -le $case.expectedSteps; $number++) {
            $stepStart = [datetime]::UtcNow
            $name = 'step-{0:D2}-01-after.png' -f $number
            $path = Join-Path $caseDirectory $name
            $image = [Drawing.Bitmap]::new(120, 100)
            try { $image.Save($path, [Drawing.Imaging.ImageFormat]::Png) }
            finally { $image.Dispose() }
            (Get-Item -LiteralPath $path).LastWriteTimeUtc = $stepStart
            @{
                Number = $number; Status = 'AutomatedChecksPassed'; Error = $null; ScreenshotError = $null
                StartedUtc = $stepStart.ToString('O'); CompletedUtc = [datetime]::UtcNow.ToString('O')
                Screenshots = @($name)
            }
        }
    )
    $end = [datetime]::UtcNow
    $report = @{
        Application = $case.application; CaseId = $case.caseId; ExpectedStepCount = $case.expectedSteps
        Status = 'NeedsVisualReview'; VisualStatus = 'ReviewRequired'; Error = $null; Steps = $steps
        StartedUtc = $start.ToString('O'); CompletedUtc = $end.ToString('O')
    }
    $reportPath = Join-Path $caseDirectory 'result.json'
    Write-Json $report $reportPath
    Write-Json @{
        RunId = 'evidence-fixture'; State = 'Completed'; Cases = @($case); CleanupErrors = @()
        StartedUtc = $start.ToString('O'); CompletedUtc = $end.ToString('O')
        Executions = @(@{ Project = $case.project; ExitCode = 0; Error = $null })
    } (Join-Path $runDirectory 'run-manifest.json')
    "<TestRun><Results><UnitTestResult testName=`"$($case.testName)`" outcome=`"Passed`" /></Results></TestRun>" |
        Set-Content -LiteralPath (Join-Path $trxDirectory 'results.trx') -Encoding utf8
    $validation = Test-OfficeRunEvidence -RunDirectory $runDirectory
    Assert-True $validation.Passed ($validation.Errors -join '; ')
    $summary = Get-Content -LiteralPath (Write-OfficeRunSummary -Validation $validation) -Raw
    Assert-True ($summary.Contains('**Automated result: Passed**') -and
        $summary.Contains("[Report]($($case.application)/$($case.caseId)/result.json)") -and
        $summary.Contains('Visual glyph review is still required')) 'Readable summary and relative evidence links'
    $report.Status = 'Failed'
    $report.Error = 'Synthetic test failure'
    Write-Json $report $reportPath
    $failed = Test-OfficeRunEvidence -RunDirectory $runDirectory
    Assert-True (-not $failed.Passed) 'Failed case must fail validation'
    $summary = Get-Content -LiteralPath (Write-OfficeRunSummary -Validation $failed) -Raw
    Assert-True ($summary.Contains('**Automated result: Failed**') -and $summary.Contains('## Failures')) 'Failure is visible in summary'
    $report.Status = 'NeedsVisualReview'
    $report.Error = $null
    Write-Json $report $reportPath
    Remove-Item -LiteralPath (Join-Path $caseDirectory 'step-01-01-after.png')
    Assert-True (-not (Test-OfficeRunEvidence -RunDirectory $runDirectory).Passed) 'Missing screenshot must fail validation'
    Write-Host 'PASS: Valid evidence and readable summary pass; failed cases and missing screenshots fail'

    # Exercise the real run-ID and environment lifecycle through the no-UI path.
    Copy-Item -LiteralPath (Join-Path $root 'Run-Tests.ps1') -Destination $sandbox
    New-Item -ItemType Directory -Path (Join-Path $sandbox 'scripts'), (Join-Path $sandbox 'data') | Out-Null
    Copy-Item -LiteralPath (Join-Path $PSScriptRoot 'TestRouting.psm1'),
        (Join-Path $PSScriptRoot 'EvidenceValidation.psm1') -Destination (Join-Path $sandbox 'scripts')
    Write-Json $catalog (Join-Path $sandbox 'data\test-catalog.json')
    foreach ($project in @($catalog.project | Select-Object -Unique)) {
        $path = Join-Path $sandbox $project
        New-Item -ItemType Directory -Path (Split-Path $path) -Force | Out-Null
        Set-Content -LiteralPath $path -Value '<Project />'
    }
    function global:dotnet {
        param([Parameter(ValueFromRemainingArguments)][object[]]$CommandArguments)
        if ($env:RUN_OFFICE_UI_TESTS -eq '1') { throw 'Infrastructure run must not enable Office UI' }
        $results = $CommandArguments[[array]::IndexOf($CommandArguments, '--results-directory') + 1]
        New-Item -ItemType Directory -Path $results -Force | Out-Null
        Set-Content -LiteralPath (Join-Path $results 'fixture.txt') -Value $env:OFFICE_UI_RUN_ID
        $global:LASTEXITCODE = 0
    }
    $runner = Join-Path $sandbox 'Run-Tests.ps1'
    & $runner -Application Word | Out-Null
    $first = @(Get-ChildItem -LiteralPath (Join-Path $sandbox 'TestResults') -Directory)
    $marker = Join-Path $first[0].FullName 'Word\fixture.txt'
    $original = Get-Content -LiteralPath $marker -Raw
    & $runner -Application Word | Out-Null
    $runs = @(Get-ChildItem -LiteralPath (Join-Path $sandbox 'TestResults') -Directory)
    Assert-True ($first.Count -eq 1 -and $runs.Count -eq 2 -and $runs[0].Name -ne $runs[1].Name) 'Unique IDs for separate invocations'
    Assert-True ((Get-Content -LiteralPath $marker -Raw) -ceq $original) 'Earlier run remains unchanged'
    Write-Host 'PASS: Manual invocations use unique run folders and preserve earlier results'
    Write-Host '4 focused workflow checks passed; no Office applications were launched.'
} finally {
    Remove-Item Function:\dotnet -ErrorAction SilentlyContinue
    if (Test-Path -LiteralPath $sandbox) { Remove-Item -LiteralPath $sandbox -Recurse -Force }
}
